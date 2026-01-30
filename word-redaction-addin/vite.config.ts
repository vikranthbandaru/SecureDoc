import { defineConfig } from 'vite';
import mkcert from 'vite-plugin-mkcert';

export default defineConfig({
  plugins: [mkcert()],
  server: {
    https: true,
    port: 3000,
    open: false
  },
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        taskpane: './src/taskpane/index.html'
      }
    }
  },
  base: './'
});
