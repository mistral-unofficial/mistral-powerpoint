/** Office.js helpers for inspecting and editing shapes, tables and slide order. Slide numbers are 1-based. */

const supports = (version: string) => Office.context.requirements.isSetSupported("PowerPointApi", version);

function requireApi(version: string, feature: string): void {
  if (!supports(version)) {
    throw new Error(`${feature} needs PowerPointApi ${version}, which this PowerPoint version doesn't support.`);
  }
}

export interface ShapeStyle {
  text?: string;
  fill_color?: string; // "#RRGGBB" or "none"
  fill_transparency?: number; // 0..1
  line_color?: string; // "#RRGGBB" or "none"
  line_width?: number; // points
  font_color?: string;
  font_size?: number;
  font_name?: string;
  bold?: boolean;
  italic?: boolean;
  align?: "Left" | "Center" | "Right" | "Justify";
  vertical_align?: "Top" | "Middle" | "Bottom";
  bullets?: boolean;
}

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface NewShape extends ShapeStyle, Box {
  kind: string; // "text_box", "line", or a geometric shape type like "Rectangle"
  name?: string;
}

export interface ShapeUpdate extends ShapeStyle, Partial<Box> {
  shape_id: string;
  name?: string;
}

export interface NewTable {
  rows: string[][];
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  style?: string;
  header_fill_color?: string;
  header_font_color?: string;
  font_size?: number;
}

async function getSlide(context: PowerPoint.RequestContext, slideNumber: number): Promise<PowerPoint.Slide> {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  const slide = slides.items[slideNumber - 1];
  if (!slide) throw new Error(`Slide ${slideNumber} doesn't exist (the presentation has ${slides.items.length} slides).`);
  return slide;
}

const round = (n: number) => Math.round(n * 10) / 10;
const isNone = (v?: string) => v?.toLowerCase() === "none" || v?.toLowerCase() === "transparent";

function applyStyle(shape: PowerPoint.Shape, style: ShapeStyle, isLine = false): void {
  if (!isLine) {
    if (style.fill_color) {
      if (isNone(style.fill_color)) shape.fill.clear();
      else shape.fill.setSolidColor(style.fill_color);
    }
    if (style.fill_transparency != null) shape.fill.transparency = style.fill_transparency;
  }
  if (style.line_color) {
    if (isNone(style.line_color)) shape.lineFormat.visible = false;
    else {
      shape.lineFormat.visible = true;
      shape.lineFormat.color = style.line_color;
    }
  }
  if (style.line_width != null) shape.lineFormat.weight = style.line_width;
  if (isLine) return;

  const range = shape.textFrame.textRange;
  if (style.text != null) range.text = style.text;
  if (style.font_color) range.font.color = style.font_color;
  if (style.font_size) range.font.size = style.font_size;
  if (style.font_name) range.font.name = style.font_name;
  if (style.bold != null) range.font.bold = style.bold;
  if (style.italic != null) range.font.italic = style.italic;
  if (style.align) range.paragraphFormat.horizontalAlignment = style.align;
  if (style.bullets != null) range.paragraphFormat.bulletFormat.visible = style.bullets;
  if (style.vertical_align) shape.textFrame.verticalAlignment = style.vertical_align;
}

/** Returns slide numbers and shape ids of the current selection (empty when unsupported). */
export async function getSelectedSlidesAndShapes(): Promise<{ slides: number[]; shapeIds: string[] }> {
  if (!supports("1.5")) return { slides: [], shapeIds: [] };
  return PowerPoint.run(async (context) => {
    const all = context.presentation.slides;
    const selectedSlides = context.presentation.getSelectedSlides();
    const selectedShapes = context.presentation.getSelectedShapes();
    all.load("items/id");
    selectedSlides.load("items/id");
    selectedShapes.load("items/id");
    await context.sync();
    const ids = all.items.map((s) => s.id);
    return {
      slides: selectedSlides.items.map((s) => ids.indexOf(s.id) + 1).filter((n) => n > 0),
      shapeIds: selectedShapes.items.map((s) => s.id),
    };
  });
}

/** Describes every shape on a slide: id, name, type, position/size in points, text, fill and table cells. */
export async function getSlideDetails(slideNumber: number): Promise<string> {
  return PowerPoint.run(async (context) => {
    const slide = await getSlide(context, slideNumber);
    const shapes = slide.shapes;
    shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    await context.sync();

    const details = shapes.items.map((s) => ({
      id: s.id,
      name: s.name,
      type: String(s.type),
      left: round(s.left),
      top: round(s.top),
      width: round(s.width),
      height: round(s.height),
    })) as Array<Record<string, unknown>>;

    // Text, fill and table contents are loaded per shape because not every shape type has them.
    for (const [i, s] of shapes.items.entries()) {
      const type = String(s.type);
      if (["GeometricShape", "TextBox", "Placeholder"].includes(type)) {
        const range = s.textFrame.textRange;
        range.load("text");
        try {
          await context.sync();
          if (range.text.trim()) details[i].text = range.text;
        } catch {
          /* no text frame */
        }
        if (type !== "Placeholder") {
          s.fill.load("type,foregroundColor");
          try {
            await context.sync();
            if (String(s.fill.type) === "Solid") details[i].fill_color = s.fill.foregroundColor;
          } catch {
            /* no fill */
          }
        }
      } else if (type === "Table" && supports("1.8")) {
        const table = s.getTable();
        table.load("values");
        try {
          await context.sync();
          details[i].rows = table.values;
        } catch {
          /* ignore */
        }
      }
    }
    return JSON.stringify({ slide: slideNumber, shapes: details });
  });
}

export async function addShapes(slideNumber: number, specs: NewShape[]): Promise<string[]> {
  return PowerPoint.run(async (context) => {
    const slide = await getSlide(context, slideNumber);
    const created: PowerPoint.Shape[] = [];
    for (const spec of specs) {
      const box = { left: spec.left, top: spec.top, width: spec.width, height: spec.height };
      const kind = spec.kind.replace(/[\s_-]/g, "").toLowerCase();
      let shape: PowerPoint.Shape;
      if (kind === "textbox" || kind === "text") {
        shape = slide.shapes.addTextBox(spec.text ?? "", box);
      } else if (kind === "line") {
        shape = slide.shapes.addLine(PowerPoint.ConnectorType.straight, box);
      } else {
        shape = slide.shapes.addGeometricShape(spec.kind as PowerPoint.GeometricShapeType, box);
      }
      if (spec.name) shape.name = spec.name;
      applyStyle(shape, spec, kind === "line");
      shape.load("id");
      created.push(shape);
    }
    await context.sync();
    return created.map((s) => s.id);
  });
}

export async function updateShapes(slideNumber: number, updates: ShapeUpdate[]): Promise<number> {
  return PowerPoint.run(async (context) => {
    const slide = await getSlide(context, slideNumber);
    for (const u of updates) {
      const shape = slide.shapes.getItem(u.shape_id);
      if (u.left != null) shape.left = u.left;
      if (u.top != null) shape.top = u.top;
      if (u.width != null) shape.width = u.width;
      if (u.height != null) shape.height = u.height;
      if (u.name) shape.name = u.name;
      shape.load("type");
      await context.sync();
      applyStyle(shape, u, String(shape.type) === "Line");
    }
    await context.sync();
    return updates.length;
  });
}

export async function deleteShapes(slideNumber: number, shapeIds: string[]): Promise<number> {
  return PowerPoint.run(async (context) => {
    const slide = await getSlide(context, slideNumber);
    shapeIds.forEach((id) => slide.shapes.getItem(id).delete());
    await context.sync();
    return shapeIds.length;
  });
}

export async function addTable(slideNumber: number, spec: NewTable): Promise<string> {
  requireApi("1.8", "Inserting tables");
  const rows = spec.rows.filter((r) => Array.isArray(r) && r.length);
  if (!rows.length) throw new Error("The table has no rows.");
  const columns = Math.max(...rows.map((r) => r.length));
  const values = rows.map((r) => Array.from({ length: columns }, (_, c) => String(r[c] ?? "")));

  return PowerPoint.run(async (context) => {
    const slide = await getSlide(context, slideNumber);
    const options: PowerPoint.TableAddOptions = { values };
    if (spec.left != null) options.left = spec.left;
    if (spec.top != null) options.top = spec.top;
    if (spec.width != null) options.width = spec.width;
    if (spec.height != null) options.height = spec.height;
    if (spec.style) options.style = spec.style as PowerPoint.TableStyle;
    if (spec.font_size) options.uniformCellProperties = { font: { size: spec.font_size } };
    if (spec.header_fill_color || spec.header_font_color) {
      options.specificCellProperties = values.map((row, r) =>
        row.map(() =>
          r === 0
            ? {
                ...(spec.header_fill_color ? { fill: { color: spec.header_fill_color } } : {}),
                font: { bold: true, ...(spec.header_font_color ? { color: spec.header_font_color } : {}) },
              }
            : {},
        ),
      );
    }
    const shape = slide.shapes.addTable(values.length, columns, options);
    shape.load("id");
    await context.sync();
    return shape.id;
  });
}

export async function deleteSlides(slideNumbers: number[]): Promise<number> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    const targets = [...new Set(slideNumbers)].map((n) => {
      const s = slides.items[n - 1];
      if (!s) throw new Error(`Slide ${n} doesn't exist.`);
      return s;
    });
    targets.forEach((s) => s.delete());
    await context.sync();
    return targets.length;
  });
}

export async function moveSlide(from: number, to: number): Promise<void> {
  requireApi("1.8", "Reordering slides");
  await PowerPoint.run(async (context) => {
    const slide = await getSlide(context, from);
    slide.moveTo(to - 1);
    await context.sync();
  });
}
