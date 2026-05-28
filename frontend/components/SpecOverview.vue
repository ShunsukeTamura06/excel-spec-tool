<script setup lang="ts">
/** 概要タブ.
 *
 * 上半分: 構造から抽出した「最初の 10 秒で知りたいこと」のダッシュボード
 *         (SpecInsightDashboard).
 * 下半分: 詳細設計書 (Markdown) を折りたたみで提供. 監査・印刷したい人向け.
 *
 * Markdown 先頭の `# 設計書: ...` H1 はページヘッダと重複するため描画時のみ除去.
 * ダウンロード用ファイルには残るので問題ない. */
import type { WorkbookData } from '~/types/api'

const props = defineProps<{ markdown: string; workbook: WorkbookData }>()

const displayMarkdown = computed(() => {
  return props.markdown.replace(/^#\s.*\n+/, '')
})

// 詳細 Markdown は折りたたみ。デフォルトは閉じておき、ダッシュボードに視線を集める.
const detailsOpen = ref(false)
</script>

<template>
  <div class="space-y-4">
    <SpecInsightDashboard :workbook="props.workbook" />

    <UCard>
      <button
        type="button"
        class="w-full flex items-center justify-between gap-2 text-left"
        :aria-expanded="detailsOpen"
        @click="detailsOpen = !detailsOpen"
      >
        <span class="flex items-center gap-2 text-sm font-medium">
          <UIcon name="i-lucide-file-text" class="size-4 text-(--ui-text-muted)" />
          詳細設計書 (Markdown)
        </span>
        <span class="flex items-center gap-2 text-xs text-(--ui-text-muted)">
          <span>{{ detailsOpen ? '閉じる' : '開く' }}</span>
          <UIcon
            :name="detailsOpen ? 'i-lucide-chevron-up' : 'i-lucide-chevron-down'"
            class="size-4"
          />
        </span>
      </button>
      <div v-if="detailsOpen" class="mt-4 border-t border-(--ui-border) pt-4">
        <ClientOnly>
          <MDC :value="displayMarkdown" tag="div" class="spec-prose" />
        </ClientOnly>
      </div>
    </UCard>
  </div>
</template>
