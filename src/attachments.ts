/** Files the user attached to the chat (images and .pptx), kept in memory for the current conversation. */

export interface Attachment {
  id: string;
  name: string;
  kind: "image" | "pptx";
  /** Raw base64 (no data: prefix), as needed by Office.js. Images are always PNG or JPEG. */
  base64: string;
  mime: string;
  width?: number;
  height?: number;
  /** Downscaled data URL sent to the model for vision. */
  previewDataUrl?: string;
}

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_VISION_EDGE = 1024;

const store = new Map<string, Attachment>();
let counter = 0;

export function getAttachment(id: string): Attachment {
  const a = store.get(id);
  if (!a) throw new Error(`Unknown attachment "${id}". Available: ${[...store.keys()].join(", ") || "none"}.`);
  return a;
}

export function clearAttachments(): void {
  store.clear();
}

export function isSupportedFile(file: File): boolean {
  return file.type.startsWith("image/") || file.type === PPTX_MIME || /\.pptx$/i.test(file.name);
}

function readAsDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error("Could not read file."));
    reader.readAsDataURL(blob);
  });
}

function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Could not decode image."));
    img.src = src;
  });
}

function drawToPng(img: HTMLImageElement, maxEdge?: number): string {
  const scale = maxEdge ? Math.min(1, maxEdge / Math.max(img.naturalWidth, img.naturalHeight)) : 1;
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
  canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
  canvas.getContext("2d")!.drawImage(img, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/png");
}

const stripPrefix = (dataUrl: string) => dataUrl.slice(dataUrl.indexOf(",") + 1);

export async function addAttachment(file: File): Promise<Attachment> {
  const id = `att_${++counter}`;
  const name = file.name || (file.type.startsWith("image/") ? `pasted-image-${counter}.png` : `file-${counter}`);

  if (file.type.startsWith("image/")) {
    const original = await readAsDataUrl(file);
    const img = await loadImage(original);
    // PowerPoint reliably accepts PNG and JPEG; convert anything else (SVG, WebP, ...) to PNG.
    const keep = file.type === "image/png" || file.type === "image/jpeg";
    const dataUrl = keep ? original : drawToPng(img);
    const attachment: Attachment = {
      id,
      name,
      kind: "image",
      base64: stripPrefix(dataUrl),
      mime: keep ? file.type : "image/png",
      width: img.naturalWidth,
      height: img.naturalHeight,
      previewDataUrl: drawToPng(img, MAX_VISION_EDGE),
    };
    store.set(id, attachment);
    return attachment;
  }

  if (isSupportedFile(file)) {
    const attachment: Attachment = {
      id,
      name,
      kind: "pptx",
      base64: stripPrefix(await readAsDataUrl(file)),
      mime: PPTX_MIME,
    };
    store.set(id, attachment);
    return attachment;
  }

  throw new Error(`${name}: only images and .pptx files can be attached.`);
}

/** One line per attachment, telling the model how it can be used. */
export function describeAttachment(a: Attachment): string {
  return a.kind === "image"
    ? `${a.id}: ${a.name} (image, ${a.width}x${a.height}px) — can be placed on a slide with insert_image`
    : `${a.id}: ${a.name} (PowerPoint file) — its slides can be added with insert_slides_from_file`;
}
