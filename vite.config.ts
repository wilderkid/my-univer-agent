import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  resolve: {
    // Force a single React instance — prevents "useContext of null" when
    // Univer's internal React root and our React root see different copies.
    dedupe: ['react', 'react-dom', 'react-dom/client'],
  },
  optimizeDeps: {
    include: ['react', 'react-dom', '@univerjs/presets', '@univerjs/preset-sheets-core'],
  },
  build: {
    chunkSizeWarningLimit: 3000,
  },
})
