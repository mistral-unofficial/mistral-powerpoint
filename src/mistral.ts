import type { Settings } from "./settings";

/** Minimal browser client for the Mistral chat completions API (streaming + tool calling). */

export interface ToolCall {
  id: string;
  type?: "function";
  function: { name: string; arguments: string };
}

export type ChatMessage =
  | { role: "system" | "user"; content: string }
  | { role: "assistant"; content: string; tool_calls?: ToolCall[] }
  | { role: "tool"; content: string; tool_call_id: string; name: string };

export interface ToolDefinition {
  type: "function";
  function: { name: string; description: string; parameters: Record<string, unknown> };
}

export interface ChatOptions {
  tools?: ToolDefinition[];
  maxTokens?: number;
  signal?: AbortSignal;
}

export interface ChatResult {
  content: string;
  toolCalls: ToolCall[];
}

export class MistralError extends Error {
  constructor(message: string, readonly status?: number) {
    super(message);
    this.name = "MistralError";
  }
}

function headers(settings: Settings): Record<string, string> {
  if (!settings.apiKey) {
    throw new MistralError("No API key set. Add your Mistral API key in Settings (⚙).");
  }
  return {
    Authorization: `Bearer ${settings.apiKey}`,
    "Content-Type": "application/json",
    Accept: "application/json",
  };
}

function url(settings: Settings, path: string): string {
  return settings.baseUrl.replace(/\/+$/, "") + path;
}

async function raiseForStatus(res: Response): Promise<void> {
  if (res.ok) return;
  let detail = res.statusText;
  try {
    const body = await res.json();
    detail = body?.message ?? body?.detail?.[0]?.msg ?? body?.error?.message ?? JSON.stringify(body);
  } catch {
    // body was not JSON
  }
  if (res.status === 401) detail = "Invalid API key (401).";
  if (res.status === 429) detail = "Rate limit reached (429). Please wait a moment and try again.";
  throw new MistralError(`Mistral API error ${res.status}: ${detail}`, res.status);
}

/**
 * Message content is usually a string, but some models (e.g. reasoning models) return an
 * array of chunks like [{type: "thinking", ...}, {type: "text", text: "..."}]. Keep only the text.
 */
function textOf(content: unknown): string {
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content
      .map((c) => (typeof c === "string" ? c : c?.type === "text" && typeof c.text === "string" ? c.text : ""))
      .join("");
  }
  return "";
}

function requestBody(settings: Settings, messages: ChatMessage[], opts: ChatOptions, stream: boolean) {
  return JSON.stringify({
    model: settings.model,
    messages,
    temperature: settings.temperature,
    stream,
    ...(opts.maxTokens ? { max_tokens: opts.maxTokens } : {}),
    ...(opts.tools?.length ? { tools: opts.tools, tool_choice: "auto" } : {}),
  });
}

/** Non-streaming completion. */
export async function chat(settings: Settings, messages: ChatMessage[], opts: ChatOptions = {}): Promise<ChatResult> {
  const res = await fetch(url(settings, "/chat/completions"), {
    method: "POST",
    headers: headers(settings),
    body: requestBody(settings, messages, opts, false),
    signal: opts.signal,
  });
  await raiseForStatus(res);
  const message = (await res.json())?.choices?.[0]?.message ?? {};
  return { content: textOf(message.content), toolCalls: message.tool_calls ?? [] };
}

/**
 * Streaming completion via server-sent events. Text deltas are reported through onText;
 * tool calls are accumulated and returned once the stream ends.
 */
export async function chatStream(
  settings: Settings,
  messages: ChatMessage[],
  onText: (full: string) => void,
  opts: ChatOptions = {},
): Promise<ChatResult> {
  const res = await fetch(url(settings, "/chat/completions"), {
    method: "POST",
    headers: { ...headers(settings), Accept: "text/event-stream" },
    body: requestBody(settings, messages, opts, true),
    signal: opts.signal,
  });
  await raiseForStatus(res);
  if (!res.body) return chat(settings, messages, opts);

  const reader = res.body.pipeThrough(new TextDecoderStream()).getReader();
  const calls = new Map<number, ToolCall>();
  let buffer = "";
  let content = "";

  const handle = (payload: string) => {
    const delta = JSON.parse(payload)?.choices?.[0]?.delta ?? {};
    const text = textOf(delta.content);
    if (text) {
      content += text;
      onText(content);
    }
    for (const [i, tc] of ((delta.tool_calls ?? []) as Array<Partial<ToolCall> & { index?: number }>).entries()) {
      const index = tc.index ?? i;
      const existing = calls.get(index);
      if (!existing) {
        calls.set(index, {
          id: tc.id ?? `call_${index}`,
          type: "function",
          function: { name: tc.function?.name ?? "", arguments: tc.function?.arguments ?? "" },
        });
      } else {
        if (tc.function?.name) existing.function.name += tc.function.name;
        if (tc.function?.arguments) existing.function.arguments += tc.function.arguments;
      }
    }
  };

  for (;;) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += value;
    const lines = buffer.split("\n");
    buffer = lines.pop() ?? "";
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed.startsWith("data:")) continue;
      const payload = trimmed.slice(5).trim();
      if (payload === "[DONE]") continue;
      try {
        handle(payload);
      } catch {
        // ignore keep-alives / malformed lines
      }
    }
  }
  return { content, toolCalls: [...calls.values()] };
}

interface ModelInfo {
  id: string;
  capabilities?: { completion_chat?: boolean; function_calling?: boolean };
}

/** Lists models available to the key that support chat with function calling. */
export async function listModels(settings: Settings): Promise<string[]> {
  const res = await fetch(url(settings, "/models"), { headers: headers(settings) });
  await raiseForStatus(res);
  const data = (await res.json()) as { data?: ModelInfo[] };
  return (data.data ?? [])
    .filter((m) => m.capabilities?.completion_chat !== false && m.capabilities?.function_calling !== false)
    .map((m) => m.id)
    .sort();
}
