import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "node:path";

// Vite dev server runs on 5173 — the origin the FastAPI CORS config allows.
export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: { "@": path.resolve(__dirname, "./src") },
  },
  server: {
    port: 5173,
    // Single-service topology: the SPA calls the API same-origin (baseURL "").
    // In dev, forward the API routes to the local FastAPI backend on :8077.
    proxy: {
      "/health": "http://localhost:8077",
      "/analyze": "http://localhost:8077",
      "/generate": "http://localhost:8077",
    },
  },
});
