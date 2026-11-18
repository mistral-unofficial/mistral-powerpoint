import { chat, chatStream, listModels, type ChatMessage } from "../mistral";
import { getSelection, replaceSelection } from "../powerpoint";
import { DEFAULT_SETTINGS, loadSettings, saveSettings, type Settings } from "../settings";
import { SYSTEM_PROMPT, TOOLS, ToolRunner, describeToolCall } from "../tools";

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

function scrollLog(): void {
  const log = $("chat-log");
  log.scrollTop = log.scrollHeight;
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

function appendActivity(text: string): HTMLDivElement {
  const div = document.createElement("div");
  div.className = "activity";
  div.textContent = `⋯ ${text}`;
  $("chat-log").append(div);
  scrollLog();
  return div;
}

/** Lets the user drop a text answer (e.g. speaker notes) into the current slide selection. */
function addInsertButton(div: HTMLDivElement, text: string): void {
  const btn = document.createElement("button");
  btn.className = "insert";
  btn.textContent = "Insert at cursor";
  btn.title = "Replace the current text selection on the slide with this answer";
  btn.addEventListener("click", async () => {
    try {
      const sel = (await getSelection()) ?? { kind: "text" as const, text: "" };
      await replaceSelection(sel, text);
      btn.textContent = "Inserted ✓";
    } catch (err) {
      btn.textContent = `Click into a text box first (${errorMessage(err)})`;
    }
  });
  div.append(document.createElement("br"), btn);
}

// ---------------------------------------------------------------------------
// Agent loop
// ---------------------------------------------------------------------------

const history: ChatMessage[] = [];
const tools = new ToolRunner();
let abort: AbortController | null = null;

function setBusy(isBusy: boolean): void {
  $<HTMLButtonElement>("chat-send").disabled = isBusy;
  $("chat-stop").hidden = !isBusy;
  document.querySelectorAll<HTMLButtonElement>("[data-quick]").forEach((b) => (b.disabled = isBusy));
}

async function send(prompt: string): Promise<void> {
  prompt = prompt.trim();
  if (!prompt || abort) return;
  $<HTMLTextAreaElement>("chat-input").value = "";

  appendMessage("user", prompt);
  const turnStart = history.length;
  history.push({ role: "user", content: prompt });

  abort = new AbortController();
  setBusy(true);
  let reply: HTMLDivElement | null = null;
  try {
    for (let step = 0; step < MAX_STEPS; step++) {
      reply = appendMessage("assistant", "…");
      const result = await chatStream(
        settings,
        [{ role: "system", content: SYSTEM_PROMPT }, ...history],
        (full) => {
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
        if (result.content) addInsertButton(reply, result.content);
        return;
      }

      // Intermediate text (if any) stays visible; otherwise drop the empty bubble.
      if (result.content.trim()) reply.textContent = result.content;
      else reply.remove();
      reply = null;

      for (const call of result.toolCalls) {
        const line = appendActivity(describeToolCall(call) + "…");
        const { result: output, summary, ok } = await tools.run(call);
        line.textContent = `${ok ? "✓" : "⚠"} ${summary}`;
        line.classList.toggle("failed", !ok);
        history.push({ role: "tool", tool_call_id: call.id, name: call.function.name, content: output });
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
  document.querySelectorAll<HTMLButtonElement>("[data-quick]").forEach((b) =>
    b.addEventListener("click", () => void send(b.dataset.quick ?? "")),
  );
  $("new-chat").addEventListener("click", () => {
    abort?.abort();
    history.length = 0;
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
      const reply = await chat(candidate, [{ role: "user", content: "Reply with the single word: OK" }], {
        maxTokens: 5,
      });
      setStatus(`Connected to ${candidate.model}: "${reply.content.trim()}"`);
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
