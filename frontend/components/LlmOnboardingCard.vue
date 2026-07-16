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
          改修相談は現在利用できません
        </span>
      </div>
    </template>

    <div class="space-y-3 text-sm">
      <p class="text-(--ui-text)">
        Excel診断と改修依頼の整理は利用できますが、この内容から変更計画を作る相談機能は
        管理者による接続設定が必要です。設定後にこの画面を開き直してください。
      </p>
    </div>
  </UCard>
</template>
