import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
// Replace 'QuestionNinja' with your exact repository name on GitHub
export default defineConfig({
  base: '/QuestionNinja/',
  plugins: [react()],
})