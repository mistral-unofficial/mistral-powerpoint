# Mistral for PowerPoint — test playbook

A short sequence of prompts that exercises most of the add-in's features. Run them in order, one message at a time.

## Before you start

1. Open a **copy** of [`mistral-template.pptx`](mistral-template.pptx) in PowerPoint (or any blank presentation).
2. Open the pane via **Home → Mistral** and make sure an API key is set (⚙).
3. Pick a model that can see images — `mistral-medium-latest` or `mistral-large-latest` — so the visual layout check in step 2 works.

While a step runs, the chat shows one activity line per tool call (e.g. `✓ Read design`, `✓ Looked at slide 3`). Use them to see which features were actually used.

Some tools need a recent PowerPoint. If a step fails with *"needs PowerPointApi 1.x"*, your PowerPoint build doesn't support that feature yet.

---

## 1. Read the design and generate slides

**Tests:** `get_design`, `list_layouts`, `add_slides` (with a named layout)

```
Check the slide size, theme colours and available layouts, then create 4 slides about "Introducing our new product": a title slide, an agenda slide, a slide with the "Two Content" layout comparing old vs. new, and a slide with 3 next steps.
```

**Expect:** four new slides at the end; the comparison slide uses the two-column layout.

## 2. Shapes, theme colours, grouping and the visual check

**Tests:** `get_slide_details`, `add_shapes`, `group_shapes`, `view_slide`

```
On slide 3, replace the bullet text with three side-by-side cards (rounded rectangles in theme accent colours, white text, a short heading and one sentence each). Group each card, then look at the slide and fix any overlaps or overflowing text.
```

**Expect:** three evenly spaced cards in the template's orange/red tones, a `Looked at slide 3` line, and possibly a correction round afterwards.

> Adjust the slide number if your deck already had slides before step 1.

## 3. Numbered list and rotation

**Tests:** `update_shapes` with list style and rotation (PowerPointApi 1.10)

```
Turn the next steps on slide 4 into a numbered list and add a small "NEW" badge rotated by 15° in the top-right corner.
```

**Expect:** the steps are numbered 1., 2., 3.; a tilted badge sits top right.

## 4. Table: insert, then edit

**Tests:** `add_table`, `update_table` (PowerPointApi 1.8 / 1.9)

```
Add a new slide with a table comparing Basic, Pro and Enterprise plans (price, users, support, storage). Then highlight the Pro column with a light accent fill, make the header row bold and add a "SSO" row at the bottom.
```

**Expect:** a new slide with a table; the Pro column is tinted, the header is bold, and the last row is "SSO".

## 5. Hyperlink

**Tests:** `add_shapes`, `add_hyperlink` (PowerPointApi 1.10)

```
On the title slide, add a text box "Learn more at mistral.ai" at the bottom and link the words "mistral.ai" to https://mistral.ai.
```

**Expect:** only the words "mistral.ai" are a link; the rest of the text is plain.

## 6. Backgrounds

**Tests:** `set_background` (PowerPointApi 1.10)

```
Give the last slide a dark solid background in the theme's Dark2 colour and make sure the text is still readable. Give all other slides except the title slide a subtle "Percent5" pattern background.
```

**Expect:** the last slide turns dark navy and its text switches to a light colour; a faint dotted pattern on the other content slides; the title slide is unchanged.

## 7. Image with alt text

**Tests:** attachments, `insert_image`

First attach a logo with 📎, or paste or drop an image into the chat, then send:

```
Put the attached logo in the top-right corner of every slide, about 80 points wide, with alt text.
```

**Expect:** the logo on every slide at the same position, with its aspect ratio kept; the alt text is visible in PowerPoint under *View Alt Text*.

## 8. Import slides from a file

**Tests:** attachments, `insert_slides_from_file`

First attach any `.pptx` file, then send:

```
Insert the slides from the attached file after slide 2 and adopt this presentation's design.
```

**Expect:** the file's slides appear after slide 2, restyled to this deck's theme.

## 9. Clean up

**Tests:** `delete_ai_content` and its confirmation dialog

```
Remove everything you created, but keep the slides that were there before.
```

Or use **Actions → Remove AI content**, which does the same thing without asking Mistral.

**Expect:** a confirmation dialog listing how many slides and items will be removed.

- **Cancel** removes nothing.
- **Remove** deletes everything Mistral created in steps 1–7.
- **Not removed:** slides imported in step 8, and changes to existing content such as backgrounds or rewritten text. Use ⌘Z / Ctrl+Z for those.

---

## Reporting a problem

Please include:

- the step number and the exact prompt you sent
- the activity lines shown in the chat, especially any `⚠ Failed: …` line
- your PowerPoint version (**PowerPoint → About PowerPoint**) and the model selected in ⚙
