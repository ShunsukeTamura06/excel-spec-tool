<script setup lang="ts">
/**
 * 直近の応答で LLM が呼んだツール一覧. 透明性を上げて
 * 「なぜこの答えになったか」をユーザが確認できるようにする.
 */

import type { ToolTraceItem } from '~/types/api'

const props = defineProps<{ items: ToolTraceItem[] }>()

const open = ref(false)

function formatArgs(args: Record<string, unknown>): string {
  try {
    return JSON.stringify(args, null, 2)
  } catch {
    return String(args)
  }
}
</script>

<template>
  <div v-if="props.items.length > 0" class="rounded-xl border border-(--ui-border) bg-(--ui-bg-elevated)/40">
    <button
      type="button"
      class="w-full flex items-center justify-between gap-2 px-3 py-2 text-xs text-(--ui-text-muted) hover:text-(--ui-text) transition-colors"
      @click="open = !open"
    >
      <span class="flex items-center gap-2">
        <UIcon name="i-lucide-wrench" class="size-3.5" />
        直近の応答で呼ばれたツール ({{ props.items.length }} 件)
      </span>
      <UIcon :name="open ? 'i-lucide-chevron-up' : 'i-lucide-chevron-down'" class="size-3.5" />
    </button>
    <div v-if="open" class="border-t border-(--ui-border) px-3 py-2 space-y-2">
      <div
        v-for="(t, i) in props.items"
        :key="i"
        class="text-xs"
      >
        <div class="flex items-center gap-2 mb-1">
          <UBadge color="info" variant="subtle" size="sm" class="font-mono">{{ t.name }}</UBadge>
        </div>
        <details class="ml-2">
          <summary class="cursor-pointer text-(--ui-text-muted) hover:text-(--ui-text) text-[11px]">
            引数
          </summary>
          <pre class="mt-1 p-2 rounded bg-(--ui-bg-elevated) overflow-x-auto text-[10px] font-mono">{{ formatArgs(t.arguments) }}</pre>
        </details>
        <details class="ml-2 mt-1">
          <summary class="cursor-pointer text-(--ui-text-muted) hover:text-(--ui-text) text-[11px]">
            結果プレビュー
          </summary>
          <pre class="mt-1 p-2 rounded bg-(--ui-bg-elevated) overflow-x-auto text-[10px] font-mono whitespace-pre-wrap">{{ t.result_preview }}</pre>
        </details>
      </div>
    </div>
  </div>
</template>
