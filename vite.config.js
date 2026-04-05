import { defineConfig } from "vite";

const buildId = new Date().toISOString();

export default defineConfig({
  define: {
    __APP_BUILD_ID__: JSON.stringify(buildId)
  },
  plugins: [
    {
      name: "golden-leaf-build-version",
      generateBundle() {
        this.emitFile({
          type: "asset",
          fileName: "version.json",
          source: JSON.stringify({ buildId }, null, 2)
        });
      }
    }
  ]
});
