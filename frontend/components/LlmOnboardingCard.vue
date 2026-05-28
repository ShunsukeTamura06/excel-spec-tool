<script setup lang="ts">
/**
 * LLM 未設定 (= MockLLMClient フォールバック) のときに表示する onboarding カード.
 *
 * Mock のままだとチャット応答が "[mock:...] received: ..." 固定になり、
 * 注釈も空のため設計書の用途列が "-" のまま. 「壊れている」と誤解されないよう
 * 「未接続」「設定手順」「Ollama 想定」を明示する.
 *
 * - configured=true なら何も出さない (= 普通の体験)
 * - backend が応答しない (status=null) ときも出さない (別の AppBackendStatus が
 *   バックエンド到達性を表示する責務)
 */

const backend = useBackend()

const { data: status } = useAsyncData('llm-status', () => backend.llmStatus(), {
  // ユーザーが頻繁にホーム/チャットを往復するので 5 分キャッシュ
  default: () => null,
})
</script>

<template>
  <UCard
    v-if="status && status.configured === false"
    :ui="{
      root: 'border-amber-400/40 dark:border-amber-700/40 bg-amber-50/40 dark:bg-amber-950/20',
      header: 'border-b border-amber-400/20 dark:border-amber-700/30',
    }"
  >
    <template #header>
      <div class="flex items-center gap-2">
        <UIcon name="i-lucide-plug-zap" class="size-4 text-amber-600 dark:text-amber-400" />
        <span class="font-semibold text-amber-900 dark:text-amber-200">
          LLM 未接続
        </span>
        <UBadge color="warning" variant="subtle" size="xs">mock mode</UBadge>
      </div>
    </template>

    <div class="space-y-3 text-sm">
      <p class="text-(--ui-text)">
        現在 LLM が設定されていないため、<strong>チャット応答と LLM 注釈はモック</strong>のままです。
        設計書生成 / 依存グラフ / 参照逆引きは LLM 無しでも動作します。
      </p>

      <div>
        <p class="text-xs text-(--ui-text-muted) mb-1.5">
          実 LLM に接続するには、Backend 側で以下を環境変数に設定してから再起動してください:
        </p>
        <pre class="font-mono text-xs bg-(--ui-bg-elevated) rounded p-2 overflow-x-auto"><code>LLM_BASE_URL=http://localhost:11434/v1   # 例: Ollama
LLM_API_KEY=ollama                       # ローカルなら任意の文字列で可
LLM_MODEL=llama3.1:8b                    # ご利用のモデル</code></pre>
        <p class="text-[11px] text-(--ui-text-muted) mt-1.5">
          OpenAI 互換 API ならどのプロバイダでも動作します (Ollama / vLLM / 社内 LLM / OpenAI 等)。
          詳細は <code class="font-mono">README.md</code> の「環境変数」セクションを参照。
        </p>
      </div>
    </div>
  </UCard>
</template>
