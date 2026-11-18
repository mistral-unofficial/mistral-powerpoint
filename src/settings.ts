export interface Settings {
  apiKey: string;
  model: string;
  temperature: number;
  baseUrl: string;
}

const STORAGE_KEY = "mistral-ppt.settings";

export const DEFAULT_SETTINGS: Settings = {
  apiKey: "",
  model: "mistral-large-latest",
  temperature: 0.4,
  baseUrl: "https://api.mistral.ai/v1",
};

export function loadSettings(): Settings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return { ...DEFAULT_SETTINGS, ...(JSON.parse(raw) as Partial<Settings>) };
  } catch {
    // storage unavailable or corrupted: fall back to defaults
  }
  return { ...DEFAULT_SETTINGS };
}

export function saveSettings(settings: Settings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}
