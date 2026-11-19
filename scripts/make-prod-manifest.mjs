// Generates manifest.prod.xml from the development manifest.xml:
//  - every https://localhost:3000/ URL points at the hosted add-in instead
//  - a separate add-in ID, so the hosted and the local dev version can be installed side by side
//
// Usage: node scripts/make-prod-manifest.mjs [baseUrl] [outFile]
//   baseUrl defaults to ADDIN_URL or the GitHub Pages URL below.
import { readFileSync, writeFileSync } from "node:fs";

const DEV_URL = "https://localhost:3000/";
const PROD_ID = "561fb2b4-9ceb-4751-823b-415cfeab862f";
const DEFAULT_URL = "https://mistral-unofficial.github.io/mistral-powerpoint/";

let base = process.argv[2] || process.env.ADDIN_URL || DEFAULT_URL;
if (!base.endsWith("/")) base += "/";
if (!base.startsWith("https://")) throw new Error(`Add-in URL must use HTTPS: ${base}`);
const out = process.argv[3] || "manifest.prod.xml";

const dev = readFileSync(new URL("../manifest.xml", import.meta.url), "utf8");
const prod = dev
  .replace(/<Id>[^<]+<\/Id>/, `<Id>${PROD_ID}</Id>`)
  .replaceAll(DEV_URL, base);

if (prod.includes("localhost")) throw new Error("manifest.prod.xml still references localhost.");
writeFileSync(out, prod);
console.log(`Wrote ${out} for ${base}`);
