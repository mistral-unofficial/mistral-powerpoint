/** PowerPoint helpers built on Office.js. */

export interface SlideSpec {
  title: string;
  bullets: string[];
  /** Optional layout name, e.g. "Two Content". Defaults to "Title and Content". */
  layout?: string;
}

/** Tag key put on every slide and shape the add-in creates, so AI content can be found and removed. */
export const AI_TAG = "MISTRAL_AI";

export interface DeckSpec {
  title?: string;
  subtitle?: string;
  slides: SlideSpec[];
}

const supports = (version: string) => Office.context.requirements.isSetSupported("PowerPointApi", version);

// ---------------------------------------------------------------------------
// Selection
// ---------------------------------------------------------------------------

export type Selection =
  | { kind: "text"; text: string }
  | { kind: "shape"; text: string; shapeId: string; slideId?: string };

function getSelectedTextCommon(): Promise<string> {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? String(result.value ?? "") : "");
    });
  });
}

/** Returns the selected text, or the text of the single selected shape if no text is highlighted. */
export async function getSelection(): Promise<Selection | null> {
  const text = (await getSelectedTextCommon()).trim();
  if (text) return { kind: "text", text };

  if (!supports("1.5")) return null;
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await context.sync();
    if (shapes.items.length !== 1) return null;
    const shape = shapes.items[0];
    const range = shape.textFrame.textRange;
    range.load("text");
    try {
      await context.sync();
    } catch {
      return null; // shape has no text frame (picture, chart, ...)
    }
    return range.text.trim() ? { kind: "shape" as const, text: range.text, shapeId: shape.id } : null;
  });
}

export async function replaceSelection(selection: Selection, newText: string): Promise<void> {
  if (selection.kind === "text") {
    await new Promise<void>((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(newText, { coercionType: Office.CoercionType.Text }, (result) =>
        result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(result.error.message)),
      );
    });
    return;
  }
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await context.sync();
    const shape = shapes.items.find((s) => s.id === selection.shapeId);
    if (!shape) throw new Error("The shape is no longer selected. Select it again and retry.");
    shape.textFrame.textRange.text = newText;
    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Reading the deck
// ---------------------------------------------------------------------------

const TEXT_SHAPE_TYPES = new Set<string>(["GeometricShape", "TextBox", "Placeholder"]);

/** Extracts the text of every slide, formatted as "Slide N:\n…". */
export async function getPresentationText(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const out: string[] = [];
    for (let i = 0; i < slides.items.length; i++) {
      const shapes = slides.items[i].shapes;
      shapes.load("items/id,items/type");
      await context.sync();

      const candidates = shapes.items.filter((s) => TEXT_SHAPE_TYPES.has(s.type));
      const ranges = candidates.map((s) => {
        const r = s.textFrame.textRange;
        r.load("text");
        return r;
      });
      const texts: string[] = [];
      try {
        await context.sync();
        ranges.forEach((r) => texts.push(r.text));
      } catch {
        // Some placeholder (e.g. picture) has no text frame: retry one by one.
        for (const s of candidates) {
          const r = s.textFrame.textRange;
          r.load("text");
          try {
            await context.sync();
            texts.push(r.text);
          } catch {
            /* skip */
          }
        }
      }
      const body = texts.map((t) => t.trim()).filter(Boolean).join("\n");
      out.push(`Slide ${i + 1}:\n${body || "(no text)"}`);
    }
    return out.join("\n\n");
  });
}

// ---------------------------------------------------------------------------
// Creating slides
// ---------------------------------------------------------------------------

interface Layouts {
  title?: string;
  content?: string;
  all: { id: string; name: string }[];
}

async function findLayouts(context: PowerPoint.RequestContext): Promise<Layouts> {
  const masters = context.presentation.slideMasters;
  masters.load("items/id");
  await context.sync();
  if (!masters.items.length) return { all: [] };
  const layouts = masters.items[0].layouts;
  layouts.load("items/id,items/name");
  await context.sync();

  const byName = (re: RegExp) => layouts.items.find((l) => re.test(l.name))?.id;
  return {
    // English and German default names, falling back to the standard Office ordering.
    title: byName(/^(title slide|titelfolie)$/i) ?? layouts.items[0]?.id,
    content: byName(/^(title and content|titel und inhalt)$/i) ?? layouts.items[1]?.id ?? layouts.items[0]?.id,
    all: layouts.items.map((l) => ({ id: l.id, name: l.name })),
  };
}

async function addSlide(context: PowerPoint.RequestContext, layoutId: string | undefined): Promise<PowerPoint.Slide> {
  context.presentation.slides.add(layoutId ? { layoutId } : undefined);
  await context.sync();
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const slide = slides.items[slides.items.length - 1];
  if (supports("1.3")) slide.tags.add(AI_TAG, "1");
  return slide;
}

/** Fills the slide's title/body placeholders, adding text boxes when the layout lacks them. */
async function fillSlide(
  context: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  title: string,
  body: string,
): Promise<void> {
  const shapes = slide.shapes;
  shapes.load("items/id,items/type,items/name");
  await context.sync();

  const placeholders = shapes.items.filter((s) => s.type === "Placeholder");
  let titleShape: PowerPoint.Shape | undefined;
  let bodyShape: PowerPoint.Shape | undefined;

  if (supports("1.8") && placeholders.length) {
    placeholders.forEach((p) => p.placeholderFormat.load("type"));
    await context.sync();
    const t = (p: PowerPoint.Shape) => String(p.placeholderFormat.type);
    titleShape = placeholders.find((p) => t(p) === "Title" || t(p) === "CenterTitle");
    bodyShape = placeholders.find((p) => ["Body", "Subtitle", "Content"].includes(t(p)));
  } else {
    [titleShape, bodyShape] = placeholders;
  }

  if (titleShape) titleShape.textFrame.textRange.text = title;
  else {
    const box = shapes.addTextBox(title, { left: 40, top: 30, width: 880, height: 70 });
    box.textFrame.textRange.font.size = 36;
    box.textFrame.textRange.font.bold = true;
  }

  if (body) {
    if (bodyShape) bodyShape.textFrame.textRange.text = body;
    else {
      const box = shapes.addTextBox(body, { left: 40, top: 120, width: 880, height: 380 });
      box.textFrame.textRange.font.size = 22;
    }
  }
  await context.sync();
}

/** Appends the generated deck to the end of the presentation and returns the number of slides added. */
export async function insertDeck(deck: DeckSpec, withTitleSlide: boolean): Promise<number> {
  return PowerPoint.run(async (context) => {
    const layouts = await findLayouts(context);
    let added = 0;

    if (withTitleSlide && deck.title) {
      const slide = await addSlide(context, layouts.title);
      await fillSlide(context, slide, deck.title, deck.subtitle ?? "");
      added++;
    }
    for (const spec of deck.slides) {
      const named = spec.layout
        ? layouts.all.find((l) => l.name.toLowerCase() === spec.layout!.toLowerCase())?.id
        : undefined;
      const slide = await addSlide(context, named ?? layouts.content);
      await fillSlide(context, slide, spec.title, spec.bullets.join("\n"));
      added++;
    }

    // Jump to the first new slide when supported.
    if (added && supports("1.5")) {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      context.presentation.setSelectedSlides([slides.items[slides.items.length - added].id]);
      await context.sync();
    }
    return added;
  });
}
