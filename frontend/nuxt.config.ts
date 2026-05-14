// https://nuxt.com/docs/api/configuration/nuxt-config
export default defineNuxtConfig({
  compatibilityDate: '2025-01-01',
  devtools: { enabled: true },

  // SPA として配信する. 開発時は Nuxt の SSR で動かし、本番は
  // `pnpm generate` で静的書き出し (Node ランタイム不要) する想定.
  // `ssr: false` は Nuxt 3.21 系の dev で vite-node IPC エラーが出るため
  // 設定せず、production の SPA 化は generate 経由で行う.
  // (関連: https://github.com/nuxt/nuxt/issues — vite-node socket path bug)
  nitro: {
    preset: 'static',
  },

  modules: [
    '@nuxt/ui',
    '@nuxtjs/mdc',
    '@pinia/nuxt',
  ],

  css: ['~/assets/css/main.css'],

  runtimeConfig: {
    public: {
      backendUrl: process.env.NUXT_PUBLIC_BACKEND_URL || 'http://localhost:8000',
    },
  },

  // Nuxt UI v3 ではデフォルトで color mode / icon が同梱される.
  ui: {
    // Tailwind v4 のテーマトークンは assets/css/main.css 側で定義する.
  },

  app: {
    head: {
      title: 'Excel 改修支援ツール',
      htmlAttrs: { lang: 'ja' },
      meta: [
        { charset: 'utf-8' },
        { name: 'viewport', content: 'width=device-width, initial-scale=1' },
        {
          name: 'description',
          content: 'VBA / 数式 / 参照関係を含む .xlsm / .xls の統合設計書を生成し改修対話を支援する',
        },
      ],
    },
  },

  typescript: {
    strict: true,
    typeCheck: false, // CI で別途実行
  },
})
