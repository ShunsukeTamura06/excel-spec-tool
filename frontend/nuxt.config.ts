// https://nuxt.com/docs/api/configuration/nuxt-config
const appBaseURL = process.env.NUXT_APP_BASE_URL || '/'

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
      backendUrl: process.env.NUXT_PUBLIC_BACKEND_URL || '/api',
    },
  },

  // 開発サーバの待受ポート. NUXT_PORT 環境変数で上書き可能.
  devServer: {
    port: Number(process.env.NUXT_PORT ?? 3001),
  },

  // Nuxt UI v3 ではデフォルトで color mode / icon / fonts が同梱される.
  // 本番が閉鎖ネットワークのため、外向き通信を伴うものはすべて切る:
  // - fonts: @nuxt/fonts を無効化 (Google/Bunny CDN への fetch を抑止).
  //   フォントは `@fontsource-variable/*` を assets/css/main.css で直 import する.
  ui: {
    fonts: false,
    // Tailwind v4 のテーマトークンは assets/css/main.css 側で定義する.
  },

  // @nuxt/icon: 使用アイコンを .vue から走査してビルドに焼き込む.
  // これにより api.iconify.design への runtime fetch を完全に排除する.
  // `@iconify-json/lucide` がインストール済みであることが前提.
  // 閉鎖網要件として fallbackToApi: false (走査漏れがあっても API は叩かない).
  icon: {
    provider: 'iconify',
    fallbackToApi: false,
    clientBundle: {
      scan: true,
      includeCustomCollections: true,
    },
  },

  app: {
    baseURL: appBaseURL.endsWith('/') ? appBaseURL : `${appBaseURL}/`,
    head: {
      title: 'Excelツール改修支援AI',
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
