export default defineNuxtConfig({
  compatibilityDate: '2026-06-26',
  devtools: { enabled: true },
  css: ['~/assets/css/main.css'],
  nitro: {
    routeRules: {
      '/api/**': { cors: false },
    },
  },
})
