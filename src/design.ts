/** PowerPointApi 1.9 / 1.10 helpers: slide size, theme colours, backgrounds, hyperlinks, table editing. */

const supports = (version: string) => Office.context.requirements.isSetSupported("PowerPointApi", version);

function requireApi(version: string, feature: string): void {
  if (!supports(version)) {
    throw new Error(`${feature} needs PowerPointApi ${version}, which this PowerPoint version doesn't support.`);
  }
}

async function getSlides(context: PowerPoint.RequestContext, numbers?: number[]): Promise<PowerPoint.Slide[]> {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  if (!numbers?.length) return slides.items;
  return numbers.map((n) => {
    const s = slides.items[n - 1];
    if (!s) throw new Error(`Slide ${n} doesn't exist (the presentation has ${slides.items.length} slides).`);
    return s;
  });
}

// ---------------------------------------------------------------------------
// Slide size and theme
// ---------------------------------------------------------------------------

export const THEME_COLORS = [
  "Dark1",
  "Light1",
  "Dark2",
  "Light2",
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
  "Hyperlink",
] as const;
export type ThemeColorName = (typeof THEME_COLORS)[number];

/** Slide size in points (null when unsupported). */
export async function getSlideSize(): Promise<{ width: number; height: number } | null> {
  if (!supports("1.10")) return null;
  return PowerPoint.run(async (context) => {
    const setup = context.presentation.pageSetup;
    setup.load("slideWidth,slideHeight");
    await context.sync();
    return { width: setup.slideWidth, height: setup.slideHeight };
  });
}

/** Slide size, theme colours of the first slide master and the background type of a slide. */
export async function getDesign(slideNumber?: number): Promise<Record<string, unknown>> {
  requireApi("1.10", "Reading the design");
  return PowerPoint.run(async (context) => {
    const setup = context.presentation.pageSetup;
    setup.load("slideWidth,slideHeight");
    const masters = context.presentation.slideMasters;
    masters.load("items/id,items/name");
    await context.sync();

    const result: Record<string, unknown> = {
      slide_width: setup.slideWidth,
      slide_height: setup.slideHeight,
    };
    const master = masters.items[0];
    if (master) {
      const colors = THEME_COLORS.map((c) => [c, master.themeColorScheme.getThemeColor(c)] as const);
      await context.sync();
      result.theme_colors = Object.fromEntries(colors.map(([name, value]) => [name, value.value]));
      result.master = master.name;
    }
    if (slideNumber) {
      const [slide] = await getSlides(context, [slideNumber]);
      const bg = slide.background;
      bg.load("isMasterBackgroundFollowed,areBackgroundGraphicsHidden");
      bg.fill.load("type");
      const solid = bg.fill.getSolidFillOrNullObject();
      solid.load("color");
      await context.sync();
      result.slide_background = {
        slide: slideNumber,
        type: bg.fill.type,
        follows_master: bg.isMasterBackgroundFollowed,
        background_graphics_hidden: bg.areBackgroundGraphicsHidden,
        ...(solid.isNullObject ? {} : { color: solid.color }),
      };
    }
    return result;
  });
}

/** Changes theme colours on every slide master (affects all shapes and text that use theme colours). */
export async function setThemeColors(colors: Partial<Record<ThemeColorName, string>>): Promise<string[]> {
  requireApi("1.10", "Changing theme colours");
  const entries = Object.entries(colors).filter(([name, value]) => THEME_COLORS.includes(name as ThemeColorName) && value);
  if (!entries.length) throw new Error(`Give at least one of: ${THEME_COLORS.join(", ")}.`);
  return PowerPoint.run(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id");
    await context.sync();
    if (!masters.items.length) throw new Error("The presentation has no slide master.");
    for (const master of masters.items) {
      for (const [name, value] of entries) master.themeColorScheme.setThemeColor(name as ThemeColorName, String(value));
    }
    await context.sync();
    return entries.map(([name]) => name);
  });
}

// ---------------------------------------------------------------------------
// Backgrounds
// ---------------------------------------------------------------------------

export interface BackgroundSpec {
  fill: "solid" | "gradient" | "pattern" | "picture" | "reset";
  color?: string;
  transparency?: number;
  gradient_type?: "Linear" | "Radial" | "Rectangular" | "Path" | "ShadeFromTitle";
  pattern?: string;
  pattern_foreground?: string;
  pattern_background?: string;
  image_base64?: string;
  hide_background_graphics?: boolean;
}

/** Sets the background of the given slides (all slides when none given). Returns the number changed. */
export async function setBackground(slideNumbers: number[] | undefined, spec: BackgroundSpec): Promise<number> {
  requireApi("1.10", "Changing slide backgrounds");
  return PowerPoint.run(async (context) => {
    const slides = await getSlides(context, slideNumbers);
    for (const slide of slides) {
      const bg = slide.background;
      switch (spec.fill) {
        case "reset":
          bg.reset();
          break;
        case "solid":
          if (!spec.color) throw new Error('A solid background needs "color".');
          bg.fill.setSolidFill({ color: spec.color, transparency: spec.transparency ?? 0 });
          break;
        case "gradient":
          bg.fill.setGradientFill({ type: spec.gradient_type ?? "Linear" });
          break;
        case "pattern":
          bg.fill.setPatternFill({
            pattern: (spec.pattern ?? "Percent20") as PowerPoint.SlideBackgroundPatternFillType,
            ...(spec.pattern_foreground ? { foregroundColor: spec.pattern_foreground } : {}),
            ...(spec.pattern_background ? { backgroundColor: spec.pattern_background } : {}),
          });
          break;
        case "picture":
          if (!spec.image_base64) throw new Error("A picture background needs an attached image.");
          bg.fill.setPictureOrTextureFill({ imageBase64: spec.image_base64, transparency: spec.transparency ?? 0 });
          break;
      }
      if (spec.hide_background_graphics != null) bg.areBackgroundGraphicsHidden = spec.hide_background_graphics;
    }
    await context.sync();
    return slides.length;
  });
}

// ---------------------------------------------------------------------------
// Hyperlinks
// ---------------------------------------------------------------------------

/** Links a whole shape, or the first occurrence of `text` inside it, to a URL / mailto: address. */
export async function addHyperlink(
  slideNumber: number,
  shapeId: string,
  address: string,
  text?: string,
  screenTip?: string,
): Promise<string> {
  requireApi("1.10", "Adding hyperlinks");
  return PowerPoint.run(async (context) => {
    const [slide] = await getSlides(context, [slideNumber]);
    const shape = slide.shapes.getItem(shapeId);
    const options: PowerPoint.HyperlinkAddOptions = { address, ...(screenTip ? { screenTip } : {}) };
    if (!text) {
      shape.setHyperlink(options);
      await context.sync();
      return "shape";
    }
    const range = shape.textFrame.textRange;
    range.load("text");
    await context.sync();
    let index = range.text.indexOf(text);
    if (index < 0) index = range.text.toLowerCase().indexOf(text.toLowerCase());
    if (index < 0) throw new Error(`"${text}" was not found in the shape's text.`);
    range.getSubstring(index, text.length).setHyperlink(options);
    await context.sync();
    return `"${range.text.substr(index, text.length)}"`;
  });
}

// ---------------------------------------------------------------------------
// Editing existing tables (1-based rows/columns)
// ---------------------------------------------------------------------------

export interface TableEdit {
  insert_rows?: { at: number; count?: number };
  insert_columns?: { at: number; count?: number };
  delete_rows?: number[];
  delete_columns?: number[];
  merge?: { row: number; column: number; row_count: number; column_count: number }[];
  cells?: {
    row: number;
    column: number;
    text?: string;
    fill_color?: string;
    font_color?: string;
    font_size?: number;
    bold?: boolean;
    align?: "Left" | "Center" | "Right" | "Justify";
  }[];
  column_widths?: { column: number; width: number }[];
  style?: string;
  header_row?: boolean;
  banded_rows?: boolean;
  first_column?: boolean;
}

const isNone = (v?: string) => /^(none|transparent|no ?fill)$/i.test(v?.trim() ?? "");

/** Accepts "#RRGGBB"/"RRGGBB"; returns undefined for "none"; throws for anything else. */
function hexColor(value: string | undefined, what: string): string | undefined {
  if (value == null || value === "" || isNone(value)) return undefined;
  const v = value.trim();
  if (/^#?[0-9a-f]{6}$/i.test(v)) return v.startsWith("#") ? v : `#${v}`;
  throw new Error(`${what}: "${value}" is not a colour. Use "#RRGGBB" or "none".`);
}

export async function updateTable(slideNumber: number, shapeId: string, edit: TableEdit): Promise<string[]> {
  requireApi("1.9", "Editing tables");
  return PowerPoint.run(async (context) => {
    const [slide] = await getSlides(context, [slideNumber]);
    const table = slide.shapes.getItem(shapeId).getTable();
    table.load("rowCount,columnCount");
    await context.sync();
    const done: string[] = [];

    /** Syncs and labels Office's terse errors (e.g. "InvalidArgument") with the step that failed. */
    const step = async (label: string) => {
      try {
        await context.sync();
      } catch (err) {
        throw new Error(`${label} failed: ${err instanceof Error ? err.message : String(err)}`);
      }
      table.load("rowCount,columnCount");
      await context.sync();
    };
    const checkRow = (r: number, label: string) => {
      if (!Number.isInteger(r) || r < 1 || r > table.rowCount) {
        throw new Error(`${label}: row ${r} is out of range (the table has ${table.rowCount} rows).`);
      }
    };
    const checkColumn = (c: number, label: string) => {
      if (!Number.isInteger(c) || c < 1 || c > table.columnCount) {
        throw new Error(`${label}: column ${c} is out of range (the table has ${table.columnCount} columns).`);
      }
    };

    // Structural changes first, each synced so later indexes refer to the updated table.
    // An insert position past the end (e.g. rowCount + 1) appends.
    if (edit.insert_rows) {
      const count = edit.insert_rows.count ?? 1;
      const at = edit.insert_rows.at;
      if (at >= 1 && at <= table.rowCount) table.rows.add(at - 1, count);
      else table.rows.add(undefined, count);
      await step("Inserting rows");
      done.push(`inserted ${count} row(s)`);
    }
    if (edit.insert_columns) {
      const count = edit.insert_columns.count ?? 1;
      const at = edit.insert_columns.at;
      if (at >= 1 && at <= table.columnCount) table.columns.add(at - 1, count);
      else table.columns.add(undefined, count);
      await step("Inserting columns");
      done.push(`inserted ${count} column(s)`);
    }
    if (edit.delete_rows?.length) {
      edit.delete_rows.forEach((r) => checkRow(r, "Deleting rows"));
      table.rows.deleteRows(edit.delete_rows.map((r) => table.rows.getItemAt(r - 1)));
      await step("Deleting rows");
      done.push(`deleted ${edit.delete_rows.length} row(s)`);
    }
    if (edit.delete_columns?.length) {
      edit.delete_columns.forEach((c) => checkColumn(c, "Deleting columns"));
      table.columns.deleteColumns(edit.delete_columns.map((c) => table.columns.getItemAt(c - 1)));
      await step("Deleting columns");
      done.push(`deleted ${edit.delete_columns.length} column(s)`);
    }
    if (edit.merge?.length) {
      for (const m of edit.merge) {
        checkRow(m.row, "Merging");
        checkColumn(m.column, "Merging");
        table.mergeCells(m.row - 1, m.column - 1, m.row_count, m.column_count);
      }
      await step("Merging cells");
      done.push(`merged ${edit.merge.length} area(s)`);
    }

    if (edit.cells?.length) {
      for (const c of edit.cells) {
        const where = `Cell (row ${c.row}, column ${c.column})`;
        checkRow(c.row, where);
        checkColumn(c.column, where);
        const cell = table.getCellOrNullObject(c.row - 1, c.column - 1);
        if (c.text != null) cell.text = c.text;
        if (c.fill_color) {
          const fill = hexColor(c.fill_color, `${where} fill_color`);
          if (fill) cell.fill.setSolidColor(fill);
          else cell.fill.clear();
        }
        const font = hexColor(c.font_color, `${where} font_color`);
        if (font) cell.font.color = font;
        if (c.font_size) cell.font.size = c.font_size;
        if (c.bold != null) cell.font.bold = c.bold;
        if (c.align) cell.horizontalAlignment = c.align;
      }
      await step("Updating cells");
      done.push(`updated ${edit.cells.length} cell(s)`);
    }

    if (edit.column_widths?.length) {
      for (const w of edit.column_widths) {
        checkColumn(w.column, "Column width");
        table.columns.getItemAt(w.column - 1).width = w.width;
      }
      await step("Setting column widths");
      done.push("set column widths");
    }

    const settings = table.styleSettings;
    if (edit.style) settings.style = edit.style as PowerPoint.TableStyle;
    if (edit.header_row != null) settings.isFirstRowHighlighted = edit.header_row;
    if (edit.banded_rows != null) settings.areRowsBanded = edit.banded_rows;
    if (edit.first_column != null) settings.isFirstColumnHighlighted = edit.first_column;
    if (edit.style || edit.header_row != null || edit.banded_rows != null || edit.first_column != null) {
      await step(edit.style ? `Applying style "${edit.style}"` : "Changing table style options");
      done.push("updated style");
    }

    return done;
  });
}
