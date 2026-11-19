import { addAttachment, clearAttachments, describeAttachment, isSupportedFile, type Attachment } from "../attachments";
import { chat, chatStream, listModels, type ChatMessage, type ContentChunk } from "../mistral";
import { DEFAULT_SETTINGS, loadSettings, saveSettings, type Settings } from "../settings";
import { SYSTEM_PROMPT, TOOLS, ToolRunner, describeToolCall, removeAiContent } from "../tools";
import { confirmDialog } from "./confirm";

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

/** Upper bound on model ↔ tool round trips per user message. */
const MAX_STEPS = 15;

let settings: Settings = loadSettings();

function errorMessage(err: unknown): string {
  if (err instanceof DOMException && err.name === "AbortError") return "Stopped.";
  if (err instanceof Error) return err.message;
  return String(err);
}

function updateBadge(): void {
  $("model-badge").textContent = settings.apiKey ? settings.model : "no API key";
}

// ---------------------------------------------------------------------------
// Views
// ---------------------------------------------------------------------------

function showSettings(show: boolean): void {
  $("settings-view").hidden = !show;
  $("chat-view").hidden = show;
  $("toggle-settings").setAttribute("aria-pressed", String(show));
  setStatus("");
}

function setStatus(message: string, isError = false): void {
  const el = $("status");
  el.hidden = !message;
  el.textContent = message;
  el.classList.toggle("error", isError);
}

// ---------------------------------------------------------------------------
// Chat log rendering
// ---------------------------------------------------------------------------

/** True while the user hasn't scrolled up; the log then follows new content. */
let pinnedToBottom = true;

function toBottom(): void {
  const log = $("chat-log");
  log.scrollTop = log.scrollHeight;
  log.lastElementChild?.scrollIntoView({ block: "end" });
}

/** Keeps the newest entry in view, also after the layout settles (fonts, buttons appearing, resizes). */
function scrollLog(): void {
  pinnedToBottom = true;
  toBottom();
  requestAnimationFrame(toBottom);
}

function initAutoScroll(): void {
  const log = $("chat-log");
  log.addEventListener("scroll", () => {
    pinnedToBottom = log.scrollTop + log.clientHeight >= log.scrollHeight - 24;
  });
  // The log's visible height changes when attachments, the Stop button or the pane size change.
  new ResizeObserver(() => {
    if (pinnedToBottom) toBottom();
  }).observe(log);
}

function appendMessage(role: "user" | "assistant", text: string): HTMLDivElement {
  $("chat-empty").hidden = true;
  const div = document.createElement("div");
  div.className = `msg ${role}`;
  div.textContent = text;
  $("chat-log").append(div);
  scrollLog();
  return div;
}

/** Animated "…" shown in a reply bubble until the first streamed text replaces it. */
function typingIndicator(): HTMLSpanElement {
  const dots = document.createElement("span");
  dots.className = "typing";
  dots.setAttribute("role", "img");
  dots.setAttribute("aria-label", "Mistral is thinking");
  dots.append(...[0, 1, 2].map(() => document.createElement("span")));
  return dots;
}

/** Activity line with a spinner while a tool runs; overwrite textContent to finish it. */
function appendRunning(text: string): HTMLDivElement {
  const div = document.createElement("div");
  div.className = "activity running";
  const spinner = document.createElement("span");
  spinner.className = "spinner";
  spinner.setAttribute("aria-hidden", "true");
  div.append(spinner, document.createTextNode(text));
  $("chat-log").append(div);
  scrollLog();
  return div;
}

function appendActivity(text: string): HTMLDivElement {
  const div = document.createElement("div");
  div.className = "activity";
  div.textContent = `⋯ ${text}`;
  $("chat-log").append(div);
  scrollLog();
  return div;
}

// ---------------------------------------------------------------------------
// Attachments
// ---------------------------------------------------------------------------

let pending: Attachment[] = [];

function renderPending(): void {
  const box = $("attachments");
  box.replaceChildren(
    ...pending.map((a) => {
      const chip = document.createElement("span");
      chip.className = "chip";
      if (a.previewDataUrl) {
        const img = document.createElement("img");
        img.src = a.previewDataUrl;
        img.alt = "";
        chip.append(img);
      }
      chip.append(document.createTextNode(a.name));
      const remove = document.createElement("button");
      remove.type = "button";
      remove.textContent = "×";
      remove.title = `Remove ${a.name}`;
      remove.addEventListener("click", () => {
        pending = pending.filter((p) => p !== a);
        renderPending();
      });
      chip.append(remove);
      return chip;
    }),
  );
  box.hidden = pending.length === 0;
}

async function attachFiles(files: Iterable<File>): Promise<void> {
  for (const file of files) {
    try {
      pending.push(await addAttachment(file));
    } catch (err) {
      appendActivity(`⚠ ${errorMessage(err)}`).classList.add("failed");
    }
  }
  renderPending();
}

function userContent(prompt: string, attachments: Attachment[]): ChatMessage {
  if (!attachments.length) return { role: "user", content: prompt };
  const chunks: ContentChunk[] = [
    { type: "text", text: `${prompt}\n\nAttached files:\n${attachments.map((a) => `- ${describeAttachment(a)}`).join("\n")}` },
  ];
  for (const a of attachments) {
    if (a.previewDataUrl) chunks.push({ type: "image_url", image_url: a.previewDataUrl });
  }
  return { role: "user", content: chunks };
}

/** Images from earlier turns are replaced by a placeholder to keep requests small. */
function withoutOldImages(messages: ChatMessage[], keepFrom: number): ChatMessage[] {
  return messages.map((m, i) => {
    if (i >= keepFrom || m.role !== "user" || typeof m.content === "string") return m;
    return {
      role: "user",
      content: m.content.map((c) => (c.type === "image_url" ? { type: "text" as const, text: "[image omitted]" } : c)),
    };
  });
}

// ---------------------------------------------------------------------------
// Predefined actions menu
// ---------------------------------------------------------------------------

function closeActionsMenu(): void {
  $("actions-menu").hidden = true;
  $("actions-toggle").setAttribute("aria-expanded", "false");
}

function initActionsMenu(): void {
  const toggle = $<HTMLButtonElement>("actions-toggle");
  const menu = $("actions-menu");
  const items = () => Array.from(menu.querySelectorAll<HTMLButtonElement>('[role="menuitem"]'));

  toggle.addEventListener("click", (e) => {
    e.stopPropagation();
    const open = menu.hidden;
    menu.hidden = !open;
    toggle.setAttribute("aria-expanded", String(open));
    if (open) items()[0]?.focus();
  });
  items().forEach((item) =>
    item.addEventListener("click", () => {
      closeActionsMenu();
      if (item.dataset.local === "remove-ai-content") void removeAiContentLocally();
      else void send(item.dataset.quick ?? "");
    }),
  );
  menu.addEventListener("keydown", (e) => {
    const list = items();
    const i = list.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      const next = (i + (e.key === "ArrowDown" ? 1 : -1) + list.length) % list.length;
      list[next].focus();
    } else if (e.key === "Escape") {
      closeActionsMenu();
      toggle.focus();
    }
  });
  document.addEventListener("click", (e) => {
    if (!menu.hidden && !menu.contains(e.target as Node)) closeActionsMenu();
  });
}

/** Runs "Remove AI content" in the add-in only: count, confirm, delete. No request goes to Mistral. */
async function removeAiContentLocally(): Promise<void> {
  if (abort) return;
  $("chat-empty").hidden = true;
  const line = appendRunning("Removing AI content…");
  setBusy(true);
  try {
    const outcome = await removeAiContent(confirmDialog);
    line.textContent = `${outcome.ok ? "✓" : "⚠"} ${outcome.summary}`;
    line.classList.toggle("failed", !outcome.ok && outcome.summary !== "Removal cancelled");
  } catch (err) {
    line.textContent = `⚠ Could not remove AI content: ${errorMessage(err)}`;
    line.classList.add("failed");
  } finally {
    line.classList.remove("running");
    setBusy(false);
  }
}

// ---------------------------------------------------------------------------
// Agent loop
// ---------------------------------------------------------------------------

const history: ChatMessage[] = [];
const tools = new ToolRunner(confirmDialog);
let abort: AbortController | null = null;

function setBusy(isBusy: boolean): void {
  $<HTMLButtonElement>("chat-send").disabled = isBusy;
  $("chat-stop").hidden = !isBusy;
  $<HTMLButtonElement>("actions-toggle").disabled = isBusy;
  if (isBusy) closeActionsMenu();
}

async function send(prompt: string): Promise<void> {
  prompt = prompt.trim();
  if ((!prompt && !pending.length) || abort) return;
  if (!prompt) prompt = "See the attached files.";
  $<HTMLTextAreaElement>("chat-input").value = "";
  const attachments = pending;
  pending = [];
  renderPending();

  const bubble = appendMessage("user", prompt);
  for (const a of attachments) {
    const tag = document.createElement("div");
    tag.className = "msg-attachment";
    tag.textContent = `📎 ${a.name}`;
    bubble.append(tag);
  }
  const turnStart = history.length;
  history.push(userContent(prompt, attachments));

  abort = new AbortController();
  setBusy(true);
  let reply: HTMLDivElement | null = null;
  try {
    for (let step = 0; step < MAX_STEPS; step++) {
      reply = appendMessage("assistant", "");
      reply.append(typingIndicator());
      const result = await chatStream(
        settings,
        [{ role: "system", content: SYSTEM_PROMPT }, ...withoutOldImages(history, turnStart)],
        (full) => {
          // Leading whitespace-only deltas would otherwise replace the typing dots with an empty bubble.
          if (!full.trim()) return;
          reply!.textContent = full;
          scrollLog();
        },
        { tools: TOOLS, signal: abort.signal },
      );

      history.push({
        role: "assistant",
        content: result.content,
        ...(result.toolCalls.length ? { tool_calls: result.toolCalls } : {}),
      });

      if (!result.toolCalls.length) {
        reply.textContent = result.content;
        return;
      }

      // Intermediate text (if any) stays visible; otherwise drop the empty bubble.
      if (result.content.trim()) reply.textContent = result.content;
      else reply.remove();
      reply = null;

      const images: ContentChunk[] = [];
      for (const call of result.toolCalls) {
        const line = appendRunning(describeToolCall(call) + "…");
        const outcome = await tools.run(call);
        line.textContent = `${outcome.ok ? "✓" : "⚠"} ${outcome.summary}`;
        line.classList.remove("running");
        line.classList.toggle("failed", !outcome.ok);
        history.push({ role: "tool", tool_call_id: call.id, name: call.function.name, content: outcome.result });
        if (outcome.image) images.push({ type: "image_url", image_url: outcome.image });
      }
      // Tool results can only carry text, so screenshots are handed to the model as a follow-up message.
      if (images.length) {
        history.push({ role: "user", content: [{ type: "text", text: "Rendered slide image(s) from view_slide:" }, ...images] });
      }
    }
    appendActivity("Stopped after too many steps.");
  } catch (err) {
    // Roll back this turn so the conversation stays valid for the API.
    history.length = turnStart;
    const msg = `⚠ ${errorMessage(err)}`;
    if (reply) reply.textContent = msg;
    else appendMessage("assistant", msg);
  } finally {
    abort = null;
    setBusy(false);
  }
}

function initChat(): void {
  initAutoScroll();
  const input = $<HTMLTextAreaElement>("chat-input");
  $<HTMLFormElement>("chat-form").addEventListener("submit", (e) => {
    e.preventDefault();
    void send(input.value);
  });
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
      e.preventDefault();
      void send(input.value);
    }
  });
  $("chat-stop").addEventListener("click", () => abort?.abort());
  initActionsMenu();

  // Attachments: paperclip, paste and drag & drop.
  const fileInput = $<HTMLInputElement>("file-input");
  $("attach").addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", () => {
    void attachFiles(Array.from(fileInput.files ?? []));
    fileInput.value = "";
  });
  input.addEventListener("paste", (e) => {
    const files = Array.from(e.clipboardData?.files ?? []).filter(isSupportedFile);
    if (files.length) {
      e.preventDefault();
      void attachFiles(files);
    }
  });
  const form = $("chat-form");
  form.addEventListener("dragover", (e) => {
    e.preventDefault();
    form.classList.add("dragging");
  });
  form.addEventListener("dragleave", () => form.classList.remove("dragging"));
  form.addEventListener("drop", (e) => {
    e.preventDefault();
    form.classList.remove("dragging");
    void attachFiles(Array.from(e.dataTransfer?.files ?? []).filter(isSupportedFile));
  });

  $("new-chat").addEventListener("click", () => {
    abort?.abort();
    history.length = 0;
    pending = [];
    clearAttachments();
    renderPending();
    $("chat-log")
      .querySelectorAll(".msg, .activity")
      .forEach((el) => el.remove());
    $("chat-empty").hidden = false;
    showSettings(false);
    input.focus();
  });
}

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

function fillModelSelect(models: string[]): void {
  const select = $<HTMLSelectElement>("set-model");
  const all = Array.from(new Set([settings.model, ...models])).filter(Boolean);
  select.replaceChildren(
    ...all.map((m) => {
      const o = document.createElement("option");
      o.value = o.textContent = m;
      return o;
    }),
  );
  select.value = settings.model;
}

function readSettingsForm(): Settings {
  return {
    apiKey: $<HTMLInputElement>("set-key").value.trim(),
    model: $<HTMLSelectElement>("set-model").value || DEFAULT_SETTINGS.model,
    temperature: Number($<HTMLInputElement>("set-temp").value),
    baseUrl: $<HTMLInputElement>("set-base").value.trim() || DEFAULT_SETTINGS.baseUrl,
  };
}

async function withButton(button: HTMLButtonElement, label: string, task: () => Promise<void>): Promise<void> {
  const original = button.textContent;
  button.disabled = true;
  button.textContent = label;
  setStatus("");
  try {
    await task();
  } catch (err) {
    setStatus(errorMessage(err), true);
  } finally {
    button.disabled = false;
    button.textContent = original;
  }
}

function initSettings(): void {
  $<HTMLInputElement>("set-key").value = settings.apiKey;
  $<HTMLInputElement>("set-base").value = settings.baseUrl;
  const temp = $<HTMLInputElement>("set-temp");
  temp.value = String(settings.temperature);
  const showTemp = () => ($("set-temp-val").textContent = `(${Number(temp.value).toFixed(2)})`);
  temp.addEventListener("input", showTemp);
  const infoToggle = $<HTMLButtonElement>("temp-info-toggle");
  infoToggle.addEventListener("click", () => {
    const info = $("temp-info");
    info.hidden = !info.hidden;
    infoToggle.setAttribute("aria-expanded", String(!info.hidden));
  });
  showTemp();
  fillModelSelect(["mistral-large-latest", "mistral-medium-latest", "mistral-small-latest"]);

  $("toggle-settings").addEventListener("click", () => showSettings($("settings-view").hidden !== false));

  const refresh = $<HTMLButtonElement>("set-refresh");
  refresh.addEventListener("click", () =>
    withButton(refresh, "…", async () => {
      const models = await listModels(readSettingsForm());
      settings.model = $<HTMLSelectElement>("set-model").value;
      fillModelSelect(models);
      setStatus(`Loaded ${models.length} models.`);
    }),
  );

  $("set-save").addEventListener("click", () => {
    settings = readSettingsForm();
    saveSettings(settings);
    updateBadge();
    if (settings.apiKey) showSettings(false);
    else setStatus("Saved — but no API key is set yet.", true);
  });

  const test = $<HTMLButtonElement>("set-test");
  test.addEventListener("click", () =>
    withButton(test, "Testing…", async () => {
      const candidate = readSettingsForm();
      // Reasoning models spend tokens on hidden "thinking" first, so leave room for a visible reply.
      const reply = await chat(candidate, [{ role: "user", content: "Reply with the single word: OK" }], {
        maxTokens: 200,
      });
      const text = reply.content.trim();
      // Any successful response proves key and model work; only quote the reply when there is one.
      setStatus(`✓ Connected to ${candidate.model}${text ? ` — replied "${text.slice(0, 40)}"` : ""}.`);
    }),
  );
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------

Office.onReady((info) => {
  initChat();
  initSettings();
  updateBadge();

  if (!settings.apiKey) {
    showSettings(true);
  } else if (info.host !== Office.HostType.PowerPoint) {
    appendActivity("Not running inside PowerPoint — slide tools will fail, plain chat still works.");
  } else if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.4")) {
    appendActivity("This PowerPoint version is too old (PowerPointApi 1.4 required).");
  }
});
