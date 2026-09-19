import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: {
    target: "es2022",
    outDir: "dist",
    sourcemap: false,
    chunkSizeWarningLimit: 1200,
    rollupOptions: {
      output: {
        manualChunks: {
          vendor: ["react", "react-dom"],
          supabase: ["@supabase/supabase-js"],
          parsers: ["xlsx", "pdfjs-dist"],
        },
      },
    },
  },
  optimizeDeps: {
    esbuildOptions: {
      target: "es2022",
    },
  },
  server: {
    port: 5173,
    host: true,
  },
});
