<script setup lang="ts">
/**
 * アップロード→抽出→分析 の進捗をフェーズ表示する.
 * 親が `phase` を切り替えると UI が遷移する.
 */

export type AnalyzePhase = 'idle' | 'extracting' | 'analyzing' | 'done' | 'error'

const props = defineProps<{
  phase: AnalyzePhase
  filename?: string
  errorMessage?: string
}>()

interface Step {
  key: 'extract' | 'analyze'
  title: string
  hint: string
}

const steps: Step[] = [
  {
    key: 'extract',
    title: 'ファイル送信・抽出',
    hint: 'VBA・数式・参照関係・全セルを抽出 (大きなファイルでは数分かかる場合があります)',
  },
  {
    key: 'analyze',
    title: '設計書生成',
    hint: 'LLM 注釈 + Markdown 統合設計書を作成 (シート/プロシージャ数によっては数分かかります)',
  },
]

function stateOf(key: Step['key']): 'pending' | 'active' | 'done' {
  if (props.phase === 'error') return 'pending'
  if (key === 'extract') {
    if (props.phase === 'extracting') return 'active'
    if (props.phase === 'analyzing' || props.phase === 'done') return 'done'
    return 'pending'
  }
  // analyze
  if (props.phase === 'analyzing') return 'active'
  if (props.phase === 'done') return 'done'
  return 'pending'
}
</script>

<template>
  <UCard class="w-full">
    <template #header>
      <div class="flex items-center gap-2">
        <UIcon
          :name="phase === 'error' ? 'i-lucide-alert-circle' : phase === 'done' ? 'i-lucide-check-circle-2' : 'i-lucide-loader-2'"
          :class="[
            phase === 'error' && 'text-(--ui-error)',
            phase === 'done' && 'text-(--ui-success)',
            (phase === 'extracting' || phase === 'analyzing') && 'animate-spin text-(--ui-primary)',
          ]"
          class="size-5"
        />
        <span class="font-medium">
          {{
            phase === 'error' ? '分析失敗'
            : phase === 'done' ? '分析完了'
            : '分析中…'
          }}
        </span>
        <span v-if="filename" class="text-xs text-(--ui-text-muted) ml-2 truncate">{{ filename }}</span>
      </div>
    </template>

    <ol class="space-y-3">
      <li
        v-for="(s, i) in steps"
        :key="s.key"
        class="flex items-start gap-3 p-3 rounded-lg transition-colors"
        :class="[
          stateOf(s.key) === 'active' && 'bg-(--ui-primary)/5',
          stateOf(s.key) === 'done' && 'opacity-70',
        ]"
      >
        <div
          class="size-7 rounded-full flex items-center justify-center text-xs font-semibold shrink-0"
          :class="[
            stateOf(s.key) === 'pending' && 'bg-(--ui-bg-muted) text-(--ui-text-muted)',
            stateOf(s.key) === 'active' && 'bg-(--ui-primary) text-white',
            stateOf(s.key) === 'done' && 'bg-(--ui-success) text-white',
          ]"
        >
          <UIcon v-if="stateOf(s.key) === 'done'" name="i-lucide-check" class="size-4" />
          <UIcon v-else-if="stateOf(s.key) === 'active'" name="i-lucide-loader-2" class="size-4 animate-spin" />
          <span v-else>{{ i + 1 }}</span>
        </div>
        <div class="flex-1 min-w-0">
          <p class="font-medium text-sm">{{ s.title }}</p>
          <p class="text-xs text-(--ui-text-muted)">{{ s.hint }}</p>
        </div>
      </li>
    </ol>

    <UAlert
      v-if="phase === 'error' && errorMessage"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      :title="errorMessage"
      class="mt-3"
    />
  </UCard>
</template>
