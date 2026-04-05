import { defineConfig } from "vite";

export default defineConfig({
  define: {
    __APP_BUILD_ID__: JSON.stringify(new Date().toISOString())
  }
});
