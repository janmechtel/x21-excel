import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";

export default defineConfig(({ mode }) => {
  return {
    plugins: [
      react(),
      // Plugin to inject env vars into HTML for mock-webview.js
      {
        name: "inject-env-vars",
        transformIndexHtml(html) {
          const mockServerPort = process.env.VITE_MOCK_SERVER_PORT || "8085";
          const skipAuth = process.env.VITE_SKIP_AUTH || "false";
          const branchName = process.env.VITE_BRANCH_NAME || "unknown";

          // Inject script before mock-webview.js loads
          return html.replace(
            '<script src="/mock-webview.js"></script>',
            `<script>
              // Expose Vite env vars to non-module scripts
              window.__MOCK_SERVER_PORT__ = ${mockServerPort};
              window.__SKIP_AUTH__ = ${skipAuth === "true" ? "true" : "false"};
              window.__BRANCH_NAME__ = '${branchName}';
            </script>
            <script src="/mock-webview.js"></script>`,
          );
        },
      },
    ],
    resolve: {
      alias: {
        "@": path.resolve(__dirname, "./src"),
        "@shared": path.resolve(__dirname, "../shared"),
      },
    },
    base: "./",
    build: {
      outDir: "../TaskPane/WebAssets",
      assetsDir: "assets",
      rollupOptions: {
        output: {
          entryFileNames: "assets/[name].[hash].js",
          chunkFileNames: "assets/[name].[hash].js",
          assetFileNames: "assets/[name].[hash].[ext]",
        },
      },
      emptyOutDir: true,
      minify: "terser",
      sourcemap: false,
      assetsInlineLimit: 4096,
    },
    server: {
      port: 5173,
      cors: true,
      host: true,
    },
    preview: {
      port: 4173,
      host: true,
    },
    // Expose env vars to client code
    define: {
      "import.meta.env.VITE_MOCK_SERVER_PORT": JSON.stringify(
        process.env.VITE_MOCK_SERVER_PORT || "8085",
      ),
      "import.meta.env.VITE_SKIP_AUTH": JSON.stringify(
        process.env.VITE_SKIP_AUTH || "false",
      ),
      "import.meta.env.VITE_BRANCH_NAME": JSON.stringify(
        process.env.VITE_BRANCH_NAME || "unknown",
      ),
    },
  };
});
