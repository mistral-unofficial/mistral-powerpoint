import { defineConfig } from "vite";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { getHttpsServerOptions } from "office-addin-dev-certs";

const __dirname = fileURLToPath(new URL(".", import.meta.url));

// Office add-ins must be served over HTTPS. office-addin-dev-certs creates and
// trusts a localhost certificate (run `npm run certs` once).
export default defineConfig(async ({ command, isPreview }) => ({
  root: "src",
  // Relative URLs, so the same build works on https://localhost:3000/ and under a
  // subpath such as https://mistral-unofficial.github.io/mistral-powerpoint/.
  base: "./",
  publicDir: resolve(__dirname, "assets"),
  build: {
    outDir: resolve(__dirname, "dist"),
    emptyOutDir: true,
    rollupOptions: {
      input: { taskpane: resolve(__dirname, "src/taskpane/taskpane.html") },
    },
  },
  server: {
    port: 3000,
    strictPort: true,
    https: command === "serve" && !isPreview ? await getHttpsServerOptions() : undefined,
  },
}));
