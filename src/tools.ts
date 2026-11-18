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
  applyLayout,
  deleteAiContent,
  deleteShapes,
  deleteSlides,
  getSelectedSlidesAndShapes,
  getSlideDetails,
  getSlideImage,
  groupShapes,
  insertImage,
  insertSlidesFromFile,
  listLayouts,
  moveSlide,
  setZOrder,
  ungroupShape,
  updateShapes,
  type NewShape,
  type ShapeUpdate,
  type ZOrder,
} from "./shapes";
import { getAttachment } from "./attachments";
import {
  THEME_COLORS,
  addHyperlink,
  getDesign,
  setBackground,
  setThemeColors,
  updateTable,
  type BackgroundSpec,
  type TableEdit,
  type ThemeColorName,
} from "./design";

export const SYSTEM_PROMPT = `You are Mistral, an assistant built into Microsoft PowerPoint. You can read and change the user's presentation with tools.

Reading:
- get_presentation_text: text of all slides. Use for summaries, reviews, speaker notes, questions about content.
- get_selection: selected text, plus the selected slide numbers and shape ids. Use when the user says "this", "this slide", "the selection", "the selected box", etc.
- get_slide_details: every shape on one slide with id, type, position, size, text, fill color, table cells and whether you created it. ALWAYS call this before changing the layout of an existing slide.
- view_slide: a rendered picture of a slide. Use it to check how a slide actually looks.
- list_layouts: the layout names available in the template.
- get_design: the real slide size, the theme colours and a slide's background. Call it before designing or recolouring.

Writing:
- replace_selection: replace the selected text.
- add_slides: append slides with title + bullets (uses the template layouts; each slide may name a layout from list_layouts).
- apply_layout: switch an existing slide to another template layout (e.g. "Two Content", "Comparison", "Section Header").
- add_shapes: add text boxes, lines and geometric shapes (Rectangle, RoundRectangle, Ellipse, Triangle, Diamond, Hexagon, Pentagon, Octagon, HomePlate, Chevron, RightArrow, LeftArrow, Donut, Teardrop, Plaque, ...), with fill/line colors and text formatting.
- update_shapes: move, resize, recolor, restyle or change the text of existing shapes (by id).
- delete_shapes: remove shapes (by id).
- add_table: insert a table.
- insert_image: place an image the user attached (ids like att_1) on a slide.
- insert_slides_from_file: insert all slides of an attached .pptx file.
- group_shapes / ungroup_shape / set_z_order: group shapes and change their stacking order.
- delete_slides / move_slide: remove or reorder slides.
- delete_ai_content: remove the slides and/or shapes you created earlier (they are tagged). Use it for "undo what you added" or "remove the AI stuff". The add-in asks the user to confirm in a dialog, so do not ask for confirmation in chat; if they cancel, accept it.
- update_table: edit an existing table (cell text and colours, insert/delete rows and columns, merge cells, column widths, style, banding).
- add_hyperlink: link a shape, or a word/phrase inside it, to a URL or mailto: address.
- set_background: solid, gradient, pattern or picture background for some or all slides, or reset to the template.
- set_theme_colors: change the deck's theme colours (e.g. brand colours). This affects the whole presentation — only do it when asked.

Layout rules:
- Slide numbers are always as shown in PowerPoint: 1 is the first slide (never 0).
- Coordinates are in points from the top-left corner. Get the real slide size from get_design or get_slide_details (slide_size); a standard 16:9 slide is 960 x 540.
- Keep ~40pt margins, align elements to a common grid, leave room below the title (title area is usually the top ~110pt), and avoid overlapping shapes unless intentional.
- For columns/cards: compute equal widths with consistent gaps (e.g. three columns: width = (880 - 2*20) / 3).
- Prefer template layouts (add_slides with layout, apply_layout) over drawing placeholder-like boxes yourself.
- After creating or rearranging a slide's layout, call view_slide once to check it; fix overlaps, overflowing or clipped text and poor contrast, then stop (at most two correction rounds).
- When placing images keep their aspect ratio (give width only, height is derived) unless asked otherwise.
- Group shapes that belong together (e.g. a card's box + text) after creating them.
- Use colors as "#RRGGBB". Take colours from the theme (get_design) so new elements match the deck; ensure text contrasts with its fill.
- Give inserted images a short alt_text for accessibility. Use list_style "numbered" for steps/sequences.
- Slide content: short titles (max ~8 words), 3-5 concise bullets (max ~15 words), no bullet characters or numbering in text.

Behaviour:
- Act directly instead of asking for confirmation, then briefly say what you did.
- You can only insert images the user attached to the chat (paperclip or paste). You cannot create charts, SmartArt, animations or speaker notes; say so if asked and offer the closest alternative (e.g. a table or shapes instead of a chart).
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
  list_style: { type: "string", enum: ["bullets", "numbered", "none"], description: "Paragraph list style." },
  indent_level: { type: "integer", description: "Paragraph indent level (0 = top level)." },
  rotation: { type: "number", description: "Rotation in degrees, clockwise." },
  alt_text: { type: "string", description: "Alternative text for screen readers." },
};
const boxProps = {
  left: { type: "number" },
  top: { type: "number" },
  width: { type: "number" },
  height: { type: "number" },
};
const slideNumber = { type: "integer", description: "Slide number as shown in PowerPoint: 1 = first slide (not 0-based)." };

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
          properties: {
            title: { type: "string" },
            bullets: { type: "array", items: { type: "string" } },
            layout: { type: "string", description: 'Optional layout name from list_layouts. Default "Title and Content".' },
          },
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
  tool("view_slide", "Returns a rendered picture of a slide so you can see its actual look.", { slide: slideNumber }, ["slide"]),
  tool("list_layouts", "Lists the names of the slide layouts available in the presentation's template."),
  tool(
    "apply_layout",
    "Changes an existing slide to another layout of the template. Existing content is moved into the new layout's placeholders.",
    { slide: slideNumber, layout: { type: "string", description: "Exact layout name from list_layouts." } },
    ["slide", "layout"],
  ),
  tool(
    "insert_image",
    "Places an attached image on a slide. If only width (or only height) is given the other is derived from the aspect ratio.",
    {
      slide: slideNumber,
      attachment_id: { type: "string", description: 'Attachment id like "att_1".' },
      ...boxProps,
      name: { type: "string" },
      alt_text: { type: "string", description: "Short description of the image for screen readers." },
    },
    ["slide", "attachment_id", "left", "top"],
  ),
  tool(
    "insert_slides_from_file",
    "Inserts all slides of an attached .pptx file into the presentation.",
    {
      attachment_id: { type: "string" },
      after_slide: { type: "integer", description: "Insert after this slide number. Omit to append at the end." },
      keep_source_formatting: {
        type: "boolean",
        description: "true keeps the file's own design; false (default) adopts this presentation's theme.",
      },
    },
    ["attachment_id"],
  ),
  tool(
    "group_shapes",
    "Groups two or more shapes on a slide into one group. Returns the group's shape id.",
    { slide: slideNumber, shape_ids: { type: "array", items: { type: "string" } }, name: { type: "string" } },
    ["slide", "shape_ids"],
  ),
  tool("ungroup_shape", "Ungroups a group shape.", { slide: slideNumber, shape_id: { type: "string" } }, ["slide", "shape_id"]),
  tool(
    "set_z_order",
    "Changes the stacking order of shapes.",
    {
      slide: slideNumber,
      shape_ids: { type: "array", items: { type: "string" } },
      position: { type: "string", enum: ["BringToFront", "BringForward", "SendToBack", "SendBackward"] },
    },
    ["slide", "shape_ids", "position"],
  ),
  tool(
    "delete_ai_content",
    "Deletes slides and/or shapes that were created by this assistant (they carry a tag).",
    {
      scope: { type: "string", enum: ["slides", "shapes", "all"], description: "What to remove." },
      slide: { ...slideNumber, description: "Only look at this slide. Omit for the whole presentation." },
    },
    ["scope"],
  ),
  tool(
    "get_design",
    "Returns the slide size (points), the theme colours and optionally the background of one slide.",
    { slide: { ...slideNumber, description: "Optional slide whose background to report." } },
  ),
  tool(
    "set_theme_colors",
    "Changes theme colours of the presentation. Everything that uses theme colours updates.",
    {
      colors: {
        type: "object",
        description: `Map of theme colour name to "#RRGGBB". Names: ${THEME_COLORS.join(", ")}.`,
        properties: Object.fromEntries(THEME_COLORS.map((c) => [c, { type: "string" }])),
      },
    },
    ["colors"],
  ),
  tool(
    "set_background",
    "Sets the background of slides.",
    {
      slides: { type: "array", items: { type: "integer" }, description: "1-based slide numbers. Omit for all slides." },
      fill: { type: "string", enum: ["solid", "gradient", "pattern", "picture", "reset"], description: '"reset" restores the template background.' },
      color: { ...color, description: 'Colour for "solid".' },
      transparency: { type: "number", description: "0..1 for solid or picture." },
      gradient_type: { type: "string", enum: ["Linear", "Radial", "Rectangular", "Path", "ShadeFromTitle"] },
      pattern: { type: "string", description: 'Pattern name, e.g. "Percent20", "LightDownwardDiagonal", "SmallGrid", "DottedGrid".' },
      pattern_foreground: color,
      pattern_background: color,
      attachment_id: { type: "string", description: 'Attached image for "picture".' },
      hide_background_graphics: { type: "boolean", description: "Hide the template's background graphics on these slides." },
    },
    ["fill"],
  ),
  tool(
    "add_hyperlink",
    "Adds a hyperlink to a shape, or to a phrase inside the shape's text.",
    {
      slide: slideNumber,
      shape_id: { type: "string" },
      address: { type: "string", description: "URL (https://...) or mailto: address." },
      text: { type: "string", description: "Phrase in the shape's text to link. Omit to link the whole shape." },
      screen_tip: { type: "string" },
    },
    ["slide", "shape_id", "address"],
  ),
  tool(
    "update_table",
    "Edits an existing table. Rows and columns are 1-based. Structural changes are applied before cell edits.",
    {
      slide: slideNumber,
      shape_id: { type: "string", description: "The table's shape id." },
      insert_rows: { type: "object", properties: { at: { type: "integer" }, count: { type: "integer" } }, required: ["at"] },
      insert_columns: { type: "object", properties: { at: { type: "integer" }, count: { type: "integer" } }, required: ["at"] },
      delete_rows: { type: "array", items: { type: "integer" } },
      delete_columns: { type: "array", items: { type: "integer" } },
      merge: {
        type: "array",
        items: {
          type: "object",
          properties: {
            row: { type: "integer" },
            column: { type: "integer" },
            row_count: { type: "integer" },
            column_count: { type: "integer" },
          },
          required: ["row", "column", "row_count", "column_count"],
        },
      },
      cells: {
        type: "array",
        items: {
          type: "object",
          properties: {
            row: { type: "integer" },
            column: { type: "integer" },
            text: { type: "string" },
            fill_color: color,
            font_color: color,
            font_size: { type: "number" },
            bold: { type: "boolean" },
            align: { type: "string", enum: ["Left", "Center", "Right", "Justify"] },
          },
          required: ["row", "column"],
        },
      },
      column_widths: {
        type: "array",
        items: { type: "object", properties: { column: { type: "integer" }, width: { type: "number" } }, required: ["column", "width"] },
      },
      style: { type: "string", description: 'Built-in table style, e.g. "MediumStyle2Accent1".' },
      header_row: { type: "boolean" },
      banded_rows: { type: "boolean" },
      first_column: { type: "boolean" },
    },
    ["slide", "shape_id"],
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
  view_slide: "Looking at slide",
  list_layouts: "Reading layouts",
  apply_layout: "Changing layout",
  insert_image: "Inserting image",
  insert_slides_from_file: "Importing slides",
  group_shapes: "Grouping shapes",
  ungroup_shape: "Ungrouping",
  set_z_order: "Reordering layers",
  delete_ai_content: "Removing AI content",
  get_design: "Reading design",
  set_theme_colors: "Changing theme colours",
  set_background: "Changing background",
  add_hyperlink: "Adding link",
  update_table: "Editing table",
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
      layout: s?.layout ? String(s.layout) : undefined,
    })),
  };
}

function slideArg(args: Record<string, any>, key = "slide"): number {
  const raw = args[key];
  const n = Number(raw);
  if (raw == null || raw === "" || !Number.isInteger(n) || n < 1) {
    const got = raw == null ? "nothing" : JSON.stringify(raw);
    throw new Error(
      `"${key}" must be a slide number as shown in PowerPoint (1 = first slide), but got ${got}.` +
        (n === 0 ? " Slide numbers are not 0-based." : ""),
    );
  }
  return n;
}

export interface ToolOutcome {
  result: string;
  summary: string;
  ok: boolean;
  /** Image (data URL) to show the model after the tool results, e.g. a slide screenshot. */
  image?: string;
}

/** Executes tool calls against PowerPoint. Remembers the last selection so replace_selection targets it. */
export interface ConfirmRequest {
  title: string;
  message: string;
  confirmLabel: string;
}

/** Asks the user to confirm a destructive action; resolves true to proceed. */
export type Confirm = (request: ConfirmRequest) => Promise<boolean>;

/**
 * Counts the content the assistant created, asks the user to confirm, then deletes it.
 * Used by the delete_ai_content tool and directly (without Mistral) by the Actions menu.
 */
export async function removeAiContent(
  confirm: Confirm,
  scope: "slides" | "shapes" | "all" = "all",
  only?: number,
): Promise<ToolOutcome> {
  const found = await deleteAiContent(scope, only, true);
  if (!found.slides && !found.shapes) {
    return { result: "Nothing created by the assistant was found, so nothing was removed.", summary: "Nothing to remove", ok: true };
  }
  const parts = [found.slides ? plural(found.slides, "slide") : "", found.shapes ? plural(found.shapes, "item") : ""].filter(Boolean);
  const ok = await confirm({
    title: "Remove AI content?",
    message:
      `This removes all items and slides created by Mistral${only ? ` on slide ${only}` : ""}: ${parts.join(" and ")}. ` +
      "Anything you added to those slides is removed with them. " +
      "Changes Mistral made to your existing content (text, colours, layouts, backgrounds) are not undone.",
    confirmLabel: "Remove",
  });
  if (!ok) {
    return {
      result: "The user cancelled the removal. Nothing was deleted. Do not try again unless the user asks.",
      summary: "Removal cancelled",
      ok: false,
    };
  }
  const { slides, shapes } = await deleteAiContent(scope, only);
  return {
    result: `Deleted ${slides} slides and ${shapes} shapes created by the assistant.`,
    summary: `Removed ${plural(slides, "slide")}, ${plural(shapes, "item")}`,
    ok: true,
  };
}

export class ToolRunner {
  private selection: Selection | null = null;

  constructor(private readonly confirm: Confirm = async () => true) {}

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
      case "view_slide": {
        const n = slideArg(args);
        const image = await getSlideImage(n);
        return { result: `The picture of slide ${n} follows in the next message.`, summary: `Looked at slide ${n}`, ok: true, image };
      }
      case "list_layouts": {
        const names = await listLayouts();
        return { result: JSON.stringify(names), summary: `Found ${plural(names.length, "layout")}`, ok: true };
      }
      case "apply_layout": {
        const n = slideArg(args);
        await applyLayout(n, String(args.layout ?? ""));
        return { result: `Slide ${n} now uses the "${args.layout}" layout.`, summary: `Slide ${n} → ${args.layout}`, ok: true };
      }
      case "insert_image": {
        const n = slideArg(args);
        const att = getAttachment(String(args.attachment_id));
        if (att.kind !== "image") throw new Error(`${att.id} is not an image.`);
        const ratio = (att.height ?? 1) / (att.width ?? 1);
        let width = Number(args.width) || 0;
        let height = Number(args.height) || 0;
        if (!width && !height) width = Math.min(400, att.width ?? 400);
        if (!height) height = width * ratio;
        if (!width) width = height / ratio;
        const id = await insertImage(n, {
          base64: att.base64,
          mime: att.mime,
          left: Number(args.left) || 0,
          top: Number(args.top) || 0,
          width,
          height,
          name: args.name ? String(args.name) : att.name,
          altText: args.alt_text ? String(args.alt_text) : undefined,
        });
        return {
          result: JSON.stringify({ shape_id: id, width: Math.round(width), height: Math.round(height) }),
          summary: `Inserted ${att.name} on slide ${n}`,
          ok: true,
        };
      }
      case "insert_slides_from_file": {
        const att = getAttachment(String(args.attachment_id));
        if (att.kind !== "pptx") throw new Error(`${att.id} is not a PowerPoint file.`);
        const after = args.after_slide != null ? slideArg(args, "after_slide") : undefined;
        const count = await insertSlidesFromFile(att.base64, after, Boolean(args.keep_source_formatting));
        return { result: `Inserted ${count} slides from ${att.name}.`, summary: `Imported ${plural(count, "slide")} from ${att.name}`, ok: true };
      }
      case "group_shapes": {
        const n = slideArg(args);
        const id = await groupShapes(n, (args.shape_ids ?? []).map(String), args.name ? String(args.name) : undefined);
        return { result: JSON.stringify({ group_shape_id: id }), summary: `Grouped shapes on slide ${n}`, ok: true };
      }
      case "ungroup_shape": {
        const n = slideArg(args);
        await ungroupShape(n, String(args.shape_id));
        return { result: "Ungrouped. Call get_slide_details for the new shape ids.", summary: `Ungrouped on slide ${n}`, ok: true };
      }
      case "set_z_order": {
        const n = slideArg(args);
        await setZOrder(n, (args.shape_ids ?? []).map(String), args.position as ZOrder);
        return { result: "Layer order changed.", summary: `Reordered layers on slide ${n}`, ok: true };
      }
      case "delete_ai_content": {
        const scope = ["slides", "shapes", "all"].includes(args.scope) ? args.scope : "all";
        const only = args.slide != null ? slideArg(args) : undefined;
        return removeAiContent(this.confirm, scope, only);
      }
      case "get_design": {
        const n = args.slide != null ? slideArg(args) : undefined;
        return { result: JSON.stringify(await getDesign(n)), summary: "Read design", ok: true };
      }
      case "set_theme_colors": {
        const changed = await setThemeColors((args.colors ?? {}) as Partial<Record<ThemeColorName, string>>);
        return { result: `Changed theme colours: ${changed.join(", ")}.`, summary: `Changed ${plural(changed.length, "theme colour")}`, ok: true };
      }
      case "set_background": {
        const spec: BackgroundSpec = { ...(args as BackgroundSpec) };
        if (spec.fill === "picture") {
          const att = getAttachment(String(args.attachment_id));
          if (att.kind !== "image") throw new Error(`${att.id} is not an image.`);
          spec.image_base64 = att.base64;
        }
        const numbers = Array.isArray(args.slides) && args.slides.length ? args.slides.map(Number) : undefined;
        const count = await setBackground(numbers, spec);
        return { result: `Background set on ${count} slides.`, summary: `Background on ${plural(count, "slide")}`, ok: true };
      }
      case "add_hyperlink": {
        const n = slideArg(args);
        const target = await addHyperlink(
          n,
          String(args.shape_id),
          String(args.address),
          args.text ? String(args.text) : undefined,
          args.screen_tip ? String(args.screen_tip) : undefined,
        );
        return { result: `Linked ${target} to ${args.address}.`, summary: `Linked ${target} on slide ${n}`, ok: true };
      }
      case "update_table": {
        const n = slideArg(args);
        const done = await updateTable(n, String(args.shape_id), args as TableEdit);
        return {
          result: done.length ? `Table updated: ${done.join(", ")}.` : "Nothing to change.",
          summary: `Edited table on slide ${n}`,
          ok: true,
        };
      }
      default:
        throw new Error(`Unknown tool ${name}`);
    }
  }
}
