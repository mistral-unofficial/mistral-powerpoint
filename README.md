# Mistral for PowerPoint

A PowerPoint task-pane add-in that brings the [Mistral AI API](https://docs.mistral.ai) into your slides.

## How it works

The pane is a single chat. Mistral is given tools (via [function calling](https://docs.mistral.ai/capabilities/function_calling/)) and decides on its own when to use them:

| Tool | Effect | Min. PowerPointApi |
| --- | --- | --- |
| `get_presentation_text` | Reads the text of every slide. | 1.4 |
| `get_selection` | Reads the selected text plus selected slide numbers and shape ids. | 1.4 (ids: 1.5) |
| `get_slide_details` | Lists every shape on a slide: id, type, position, size, text, fill, table cells. | 1.4 |
| `replace_selection` | Replaces the selected text. | 1.4 |
| `add_slides` | Appends title/bullet slides using your template's layouts. | 1.4 |
| `add_shapes` | Adds text boxes, lines and geometric shapes (rectangles, ellipses, arrows, chevrons, …) with colors and text formatting. | 1.4 |
| `update_shapes` | Moves, resizes, recolors or re-texts existing shapes. | 1.4 |
| `delete_shapes` | Removes shapes. | 1.4 |
| `add_table` | Inserts a table, optionally styled with a header color. | 1.8 |
| `delete_slides` / `move_slide` | Removes or reorders slides. | 1.4 / 1.8 |

Not possible through Office.js today: inserting images, charts, SmartArt, animations, or changing themes and backgrounds.

So you can just type things like *"Create 5 slides about our Q3 results"*, *"Turn the bullets on slide 3 into three colored cards"*, *"Translate the selected text to German"* or *"What's missing from this deck?"*. Tool activity is shown inline in the chat (✓ Read 12 slides, ✓ Added 5 slides…). Quick-action buttons above the input send common requests; **Insert at cursor** under an answer puts it into the current slide selection. Settings (API key, model, temperature, base URL) are behind the ⚙ button; **＋** starts a new chat.

The system prompt and tool definitions live in [`src/tools.ts`](src/tools.ts). The selected model must support function calling (current Mistral Large/Medium/Small do).

Calls go directly from the add-in to `https://api.mistral.ai/v1` (the API allows browser CORS). The API key is kept in the add-in's `localStorage` on your machine only.

## Requirements

- Node.js 20+
- PowerPoint (Microsoft 365 desktop on Windows/macOS, or PowerPoint on the web) with **PowerPointApi 1.4+**. Placeholder detection uses 1.8 when available; selection fallbacks use 1.5.
- A Mistral API key from <https://console.mistral.ai>

## Getting started

```bash
npm install
```

```bash
npm run certs
```

Installs and trusts a localhost HTTPS certificate (one-time; asks for your password).

```bash
npm run dev
```

Serves the add-in at `https://localhost:3000`. Then, in a second terminal, sideload it into PowerPoint desktop:

```bash
npm start
```

Open PowerPoint → **Home** tab → **Mistral** button. Paste your API key in the settings (⚙), then **Save**.

Stop sideloading with `npm run stop`.

### PowerPoint on the web

With `npm run dev` running, open a presentation in PowerPoint on the web → **Home → Add-ins → More add-ins → My add-ins → Upload my add-in** and choose `manifest.xml`.

## Scripts

| Script | Purpose |
| --- | --- |
| `npm run dev` | Vite dev server with HTTPS on port 3000 |
| `npm run build` | Production build into `dist/` |
| `npm run typecheck` | TypeScript type-check |
| `npm run validate` | Validate `manifest.xml` |
| `npm start` / `npm run stop` | Sideload / unload in PowerPoint desktop |
| `npm run icons` | Regenerate the icons in `assets/` from `app-icon@2.png` (macOS, uses `sips`) |

## Project layout

```
manifest.xml              Office add-in manifest (Home tab button → task pane)
assets/                   Icons (served at the web root)
src/taskpane/             Task pane HTML, CSS, chat UI and agent loop
src/mistral.ts            Mistral client: chat, SSE streaming with tool calls, model listing
src/tools.ts              System prompt, tool definitions and tool execution
src/powerpoint.ts         Office.js helpers: selection, deck text, slide creation
src/settings.ts           Settings persistence
```

## Deploying

1. `npm run build` and host `dist/` on any HTTPS static host.
2. Replace every `https://localhost:3000` in `manifest.xml` with your host URL.
3. Distribute the manifest via the Microsoft 365 admin center (centralized deployment) or a SharePoint app catalog.

> **Note:** Because the key lives in the browser, every user brings their own Mistral key. For an organisation-wide key, put a small backend proxy in front of the API and point **API base URL** at it.
