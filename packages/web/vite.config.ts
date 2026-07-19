import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  base: process.env.BASE_PATH ?? "/",
  plugins: [
    react(),
    {
      name: "privacy-csp",
      apply: "build",
      transformIndexHtml(html) {
        const policy = "default-src 'self'; base-uri 'self'; connect-src 'none'; font-src 'self'; form-action 'none'; img-src 'self' data:; manifest-src 'self'; object-src 'none'; script-src 'self'; style-src 'self'; worker-src 'self'";
        return html.replace("<head>", `<head>\n    <meta http-equiv="Content-Security-Policy" content="${policy}">`);
      }
    }
  ],
  build: {
    target: "es2022",
    sourcemap: true
  }
});
