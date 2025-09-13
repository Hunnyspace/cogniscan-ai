import { defineConfig } from 'vite';

// https://vitejs.dev/config/
export default defineConfig({
  define: {
    'process.env.API_KEY': JSON.stringify(process.env.API_KEY)
  },
  build: {
    rollupOptions: {
      external: ['pdf-lib']
    }
  }
});