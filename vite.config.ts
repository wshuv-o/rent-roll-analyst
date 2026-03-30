import { defineConfig } from "vite";
import react from "@vitejs/plugin-react-swc";
import path from "path";
import { componentTagger } from "lovable-tagger";

// https://vitejs.dev/config/
export default defineConfig(({ mode }) => ({
  server: {
    host: "::",
    port: 8080,
    allowedHosts: ["dynamic-rentroll.bulkscraper.cloud"], // for dev
    hmr: {
      overlay: false,
    },
  },

  preview: {
    host: "::",
    port: 8080,
    allowedHosts: ["dynamic-rentroll.bulkscraper.cloud"], // 🔥 THIS is what you need
  },

  plugins: [react(), mode === "development" && componentTagger()].filter(Boolean),

  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src"),
      "exceljs": path.resolve(__dirname, "node_modules/exceljs/dist/es5/exceljs.browser.js"),
    },
  },
}));
