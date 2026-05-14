<script setup lang="ts">
/**
 * ダイアグラムタブ全体. backend から DiagramSet を取得し、シート依存と
 * VBA コールグラフを切替可能にする.
 */

import type { DiagramSet } from '~/types/api'

const props = defineProps<{ jobId: string }>()

const backend = useBackend()
const data = ref<DiagramSet | null>(null)
const loading = ref(false)
const errorMsg = ref('')

async function load() {
  loading.value = true
  errorMsg.value = ''
  try {
    data.value = await backend.getDiagrams(props.jobId)
  } catch (e) {
    errorMsg.value = friendlyMessage(e)
  } finally {
    loading.value = false
  }
}

await load()

const view = ref<'sheet_deps' | 'vba_calls'>('sheet_deps')
const direction = ref<'LR' | 'TB'>('LR')

const current = computed(() => {
  if (!data.value) return null
  return view.value === 'sheet_deps' ? data.value.sheet_deps : data.value.vba_calls
})

const viewItems = computed(() => [
  {
    value: 'sheet_deps' as const,
    label: 'シート依存',
    icon: 'i-lucide-layout-grid',
    count: data.value?.sheet_deps.nodes.length ?? 0,
  },
  {
    value: 'vba_calls' as const,
    label: 'VBA コール',
    icon: 'i-lucide-zap',
    count: data.value?.vba_calls.nodes.length ?? 0,
  },
])
</script>

<template>
  <div class="space-y-3">
    <!-- ヘッダーコントロール -->
    <div class="flex flex-wrap items-center gap-2 justify-between">
      <div class="flex flex-wrap gap-1.5">
        <UButton
          v-for="it in viewItems"
          :key="it.value"
          :icon="it.icon"
          :color="view === it.value ? 'primary' : 'neutral'"
          :variant="view === it.value ? 'soft' : 'ghost'"
          size="sm"
          @click="view = it.value"
        >
          {{ it.label }}
          <UBadge
            :color="view === it.value ? 'primary' : 'neutral'"
            variant="subtle"
            size="sm"
            class="ml-1.5 tabular-nums"
          >
            {{ it.count }}
          </UBadge>
        </UButton>
      </div>
      <div class="flex items-center gap-2">
        <span class="text-[10px] uppercase tracking-wide text-(--ui-text-muted)">レイアウト</span>
        <UButtonGroup size="sm">
          <UButton
            icon="i-lucide-move-horizontal"
            :variant="direction === 'LR' ? 'solid' : 'soft'"
            :color="direction === 'LR' ? 'primary' : 'neutral'"
            @click="direction = 'LR'"
          >
            横
          </UButton>
          <UButton
            icon="i-lucide-move-vertical"
            :variant="direction === 'TB' ? 'solid' : 'soft'"
            :color="direction === 'TB' ? 'primary' : 'neutral'"
            @click="direction = 'TB'"
          >
            縦
          </UButton>
        </UButtonGroup>
        <UButton
          icon="i-lucide-refresh-cw"
          variant="ghost"
          color="neutral"
          size="sm"
          :loading="loading"
          @click="load"
        />
      </div>
    </div>

    <UAlert
      v-if="errorMsg"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="ダイアグラム取得失敗"
      :description="errorMsg"
    />

    <!-- 凡例 -->
    <div class="flex flex-wrap items-center gap-3 text-xs text-(--ui-text-muted)">
      <span class="flex items-center gap-1">
        <span class="size-3 rounded border-2 border-emerald-400 bg-emerald-50 dark:bg-emerald-950" />
        シート
      </span>
      <span class="flex items-center gap-1">
        <span class="size-3 rounded border-2 border-sky-400 bg-sky-50 dark:bg-sky-950" />
        VBA プロシージャ
      </span>
      <span class="flex items-center gap-1">
        <svg width="20" height="6"><line x1="0" y1="3" x2="20" y2="3" stroke="currentColor" stroke-width="2"/></svg>
        矢印の太さ = 参照回数 (×N はエッジに併記)
      </span>
    </div>

    <ClientOnly>
      <DiagramView v-if="current" :diagram="current" :direction="direction" />
      <template #fallback>
        <div class="rounded-xl border border-(--ui-border) p-10 text-center text-sm text-(--ui-text-muted)">
          ダイアグラムを読み込んでいます...
        </div>
      </template>
    </ClientOnly>
  </div>
</template>
