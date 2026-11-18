import type { ToolCall, ToolDefinition } from "./mistral";
import {
  getPresentationText,
  getSelection,
  insertDeck,
  replaceSelection,
  type DeckSpec,
  type Selection,
} from "./powerpoint";
import {
  addShapes,
  addTable,
  deleteShapes,
  deleteSlides,
  getSelectedSlidesAndShapes,
  getSlideDetails,
  moveSlide,
  updateShapes,
  type NewShape,
  type ShapeUpdate,
} from "./shapes";

export const SYSTEM_PROMPT = `You are Mistral, an assistant built into Microsoft PowerPoint. You can read and change the user's presentation with tools.

Reading:
- get_presentation_text: text of all slides. Use for summaries, reviews, speaker notes, questions about content.
- get_selection: selected text, plus the selected slide numbers and shape ids. Use when the user says "this", "this slide", "the selection", "the selected box", etc.
- get_slide_details: every shape on one slide with id, type, position, size, text, fill color and table cells. ALWAYS call this before changing the layout of an existing slide.

Writing:
- replace_selection: replace the selected text.
- add_slides: append slides with title + bullets (uses the template layouts).
- add_shapes: add text boxes, lines and geometric shapes (Rectangle, RoundRectangle, Ellipse, Triangle, Diamond, Hexagon, Pentagon, Octagon, HomePlate, Chevron, RightArrow, LeftArrow, Donut, Teardrop, Plaque, ...), with fill/line colors and text formatting.
- update_shapes: move, resize, recolor, restyle or change the text of existing shapes (by id).
- delete_shapes: remove shapes (by id).
- add_table: insert a table.
- delete_slides / move_slide: remove or reorder slides.

Layout rules:
- Coordinates are in points from the top-left corner. A standard 16:9 slide is 960 x 540 points (4:3 is 720 x 540); infer the size from existing shapes if unsure.
- Keep ~40pt margins, align elements to a common grid, leave room below the title (title area is usually the top ~110pt), and avoid overlapping shapes unless intentional.
- For columns/cards: compute equal widths with consistent gaps (e.g. three columns: width = (880 - 2*20) / 3).
- Use colors as "#RRGGBB". Prefer a restrained palette that fits the existing deck; ensure text contrasts with its fill.
- Slide content: short titles (max ~8 words), 3-5 concise bullets (max ~15 words), no bullet characters or numbering in text.

Behaviour:
- Act directly instead of asking for confirmation, then briefly say what you did.
- You cannot insert images, charts, SmartArt, animations or change themes/backgrounds; say so if asked and offer the closest alternative.
- When rewriting text, keep one line per paragraph and preserve the language unless asked to translate.
- Answer in the user's language.`;

const color = { type: "string", description: 'Hex color "#RRGGBB", or "none".' };
const styleProps = {
  text: { type: "string", description: "Text content. Use newlines for separate paragraphs." },
  fill_color: color,
  fill_transparency: { type: "number", description: "0 (opaque) to 1 (fully transparent)." },
  line_color: { ...color, description: 'Outline color "#RRGGBB", or "none" to hide the outline.' },
  line_width: { type: "number", description: "Outline width in points." },
  font_color: color,
  font_size: { type: "number" },
  font_name: { type: "string" },
  bold: { type: "boolean" },
  italic: { type: "boolean" },
  align: { type: "string", enum: ["Left", "Center", "Right", "Justify"] },
  vertical_align: { type: "string", enum: ["Top", "Middle", "Bottom"] },
  bullets: { type: "boolean", description: "Show bullet points for the paragraphs." },
};
const boxProps = {
  left: { type: "number" },
  top: { type: "number" },
  width: { type: "number" },
  height: { type: "number" },
};
const slideNumber = { type: "integer", description: "1-based slide number." };

function tool(name: string, description: string, properties: Record<string, unknown> = {}, required: string[] = []): ToolDefinition {
  return { type: "function", function: { name, description, parameters: { type: "object", properties, required } } };
}

export const TOOLS: ToolDefinition[] = [
  tool("get_presentation_text", "Returns the text of every slide in the presentation, labelled by slide number."),
  tool(
    "get_selection",
    "Returns the selected text (or text of the single selected text box), the selected slide numbers and selected shape ids.",
  ),
  tool("get_slide_details", "Lists all shapes on a slide with id, name, type, position, size, text, fill color and table cells.", { slide: slideNumber }, ["slide"]),
  tool(
    "replace_selection",
    "Replaces the currently selected text (or selected text box contents) with the given text.",
    { text: { type: "string", description: "The new text. Use newlines to separate paragraphs/bullets." } },
    ["text"],
  ),
  tool(
    "add_slides",
    "Appends slides to the end of the presentation using the template layouts, optionally starting with a title slide.",
    {
      title_slide: {
        type: "object",
        description: "Optional title slide to insert first.",
        properties: { title: { type: "string" }, subtitle: { type: "string" } },
        required: ["title"],
      },
      slides: {
        type: "array",
        description: "Content slides in order.",
        items: {
          type: "object",
          properties: { title: { type: "string" }, bullets: { type: "array", items: { type: "string" } } },
          required: ["title", "bullets"],
        },
      },
    },
    ["slides"],
  ),
  tool(
    "add_shapes",
    "Adds text boxes, lines or geometric shapes to a slide. Returns the new shape ids.",
    {
      slide: slideNumber,
      shapes: {
        type: "array",
        items: {
          type: "object",
          properties: {
            kind: {
              type: "string",
              description:
                '"text_box", "line" (from left/top to left+width/top+height), or a geometric shape type such as "Rectangle", "RoundRectangle", "Ellipse", "Triangle", "Diamond", "Hexagon", "Chevron", "HomePlate", "RightArrow".',
            },
            name: { type: "string" },
            ...boxProps,
            ...styleProps,
          },
          required: ["kind", "left", "top", "width", "height"],
        },
      },
    },
    ["slide", "shapes"],
  ),
  tool(
    "update_shapes",
    "Changes existing shapes on a slide: position, size, text, colors and text formatting. Only given properties change.",
    {
      slide: slideNumber,
      updates: {
        type: "array",
        items: {
          type: "object",
          properties: { shape_id: { type: "string" }, name: { type: "string" }, ...boxProps, ...styleProps },
          required: ["shape_id"],
        },
      },
    },
    ["slide", "updates"],
  ),
  tool(
    "delete_shapes",
    "Deletes shapes from a slide by id.",
    { slide: slideNumber, shape_ids: { type: "array", items: { type: "string" } } },
    ["slide", "shape_ids"],
  ),
  tool(
    "add_table",
    "Inserts a table on a slide. The first row is treated as the header when header colors are given.",
    {
      slide: slideNumber,
      rows: { type: "array", description: "Table cells, row by row.", items: { type: "array", items: { type: "string" } } },
      ...boxProps,
      style: {
        type: "string",
        description: 'Optional built-in table style, e.g. "MediumStyle2Accent1", "LightStyle1Accent2", "NoStyleTableGrid".',
      },
      header_fill_color: color,
      header_font_color: color,
      font_size: { type: "number" },
    },
    ["slide", "rows"],
  ),
  tool(
    "delete_slides",
    "Deletes slides by number.",
    { slides: { type: "array", items: { type: "integer" }, description: "1-based slide numbers." } },
    ["slides"],
  ),
  tool(
    "move_slide",
    "Moves a slide to a new position.",
    { from: slideNumber, to: { type: "integer", description: "1-based target position." } },
    ["from", "to"],
  ),
];

const LABELS: Record<string, string> = {
  get_presentation_text: "Reading presentation",
  get_selection: "Reading selection",
  get_slide_details: "Inspecting slide",
  replace_selection: "Replacing selection",
  add_slides: "Adding slides",
  add_shapes: "Adding shapes",
  update_shapes: "Updating shapes",
  delete_shapes: "Deleting shapes",
  add_table: "Adding table",
  delete_slides: "Deleting slides",
  move_slide: "Moving slide",
};

/** Human-readable label for the activity line shown in the chat. */
export function describeToolCall(call: ToolCall): string {
  return LABELS[call.function.name] ?? call.function.name;
}

const clean = (s: unknown) => String(s ?? "").replace(/^\s*(?:[-•*]|\d+[.)])\s+/, "").trim();
const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "s"}`;

function toDeck(args: Record<string, any>): DeckSpec {
  const slides = Array.isArray(args.slides) ? args.slides : [];
  return {
    title: args.title_slide?.title ? String(args.title_slide.title) : undefined,
    subtitle: args.title_slide?.subtitle ? String(args.title_slide.subtitle) : undefined,
    slides: slides.map((s: any) => ({
      title: clean(s?.title),
      bullets: Array.isArray(s?.bullets) ? s.bullets.map(clean).filter(Boolean) : [],
    })),
  };
}

function slideArg(args: Record<string, any>, key = "slide"): number {
  const n = Number(args[key]);
  if (!Number.isInteger(n) || n < 1) throw new Error(`"${key}" must be a slide number starting at 1.`);
  return n;
}

interface ToolOutcome {
  result: string;
  summary: string;
  ok: boolean;
}

/** Executes tool calls against PowerPoint. Remembers the last selection so replace_selection targets it. */
export class ToolRunner {
  private selection: Selection | null = null;

  /** Runs a tool and returns the result for the model plus a short summary for the UI. Errors go back to the model. */
  async run(call: ToolCall): Promise<ToolOutcome> {
    try {
      const args = call.function.arguments ? JSON.parse(call.function.arguments) : {};
      return await this.dispatch(call.function.name, args);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return { result: `Error: ${message}`, summary: `Failed: ${message}`, ok: false };
    }
  }

  private async dispatch(name: string, args: Record<string, any>): Promise<ToolOutcome> {
    switch (name) {
      case "get_presentation_text": {
        const text = await getPresentationText();
        const count = (text.match(/^Slide \d+:/gm) ?? []).length;
        return { result: text || "(empty presentation)", summary: `Read ${plural(count, "slide")}`, ok: true };
      }
      case "get_selection": {
        const selection = await getSelection();
        const where = await getSelectedSlidesAndShapes();
        this.selection = selection;
        return {
          result: JSON.stringify({
            selected_text: selection?.text ?? null,
            selected_slides: where.slides,
            selected_shape_ids: where.shapeIds,
          }),
          summary: selection ? "Read selection" : where.slides.length ? `On slide ${where.slides.join(", ")}` : "No selection",
          ok: true,
        };
      }
      case "get_slide_details": {
        const n = slideArg(args);
        return { result: await getSlideDetails(n), summary: `Inspected slide ${n}`, ok: true };
      }
      case "replace_selection": {
        const selection = this.selection ?? (await getSelection());
        if (!selection) return { result: "Nothing is selected, so nothing was replaced.", summary: "No selection", ok: false };
        await replaceSelection(selection, String(args.text ?? ""));
        this.selection = null;
        return { result: "Selection replaced.", summary: "Replaced selection", ok: true };
      }
      case "add_slides": {
        const deck = toDeck(args);
        if (!deck.slides.length && !deck.title) throw new Error("No slides given.");
        const added = await insertDeck(deck, Boolean(deck.title));
        return { result: `Added ${added} slides at the end of the presentation.`, summary: `Added ${plural(added, "slide")}`, ok: true };
      }
      case "add_shapes": {
        const n = slideArg(args);
        const specs = (Array.isArray(args.shapes) ? args.shapes : []) as NewShape[];
        if (!specs.length) throw new Error("No shapes given.");
        const ids = await addShapes(n, specs);
        return { result: JSON.stringify({ created_shape_ids: ids }), summary: `Added ${plural(ids.length, "shape")} to slide ${n}`, ok: true };
      }
      case "update_shapes": {
        const n = slideArg(args);
        const updates = (Array.isArray(args.updates) ? args.updates : []) as ShapeUpdate[];
        const count = await updateShapes(n, updates);
        return { result: `Updated ${count} shapes.`, summary: `Updated ${plural(count, "shape")} on slide ${n}`, ok: true };
      }
      case "delete_shapes": {
        const n = slideArg(args);
        const count = await deleteShapes(n, (args.shape_ids ?? []).map(String));
        return { result: `Deleted ${count} shapes.`, summary: `Deleted ${plural(count, "shape")} on slide ${n}`, ok: true };
      }
      case "add_table": {
        const n = slideArg(args);
        const id = await addTable(n, { ...args, rows: args.rows ?? [] } as Parameters<typeof addTable>[1]);
        return { result: JSON.stringify({ table_shape_id: id }), summary: `Added table to slide ${n}`, ok: true };
      }
      case "delete_slides": {
        const numbers = (Array.isArray(args.slides) ? args.slides : []).map(Number);
        const count = await deleteSlides(numbers);
        return { result: `Deleted ${count} slides. Remaining slides have been renumbered.`, summary: `Deleted ${plural(count, "slide")}`, ok: true };
      }
      case "move_slide": {
        const from = slideArg(args, "from");
        const to = slideArg(args, "to");
        await moveSlide(from, to);
        return { result: `Moved slide ${from} to position ${to}.`, summary: `Moved slide ${from} → ${to}`, ok: true };
      }
      default:
        throw new Error(`Unknown tool ${name}`);
    }
  }
}
