import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Relative asset URLs are required when the built app is loaded from an
// Electron file:// URL rather than from the Vite development server.
export default defineConfig({
  base: './',
  plugins: [react()],
});
