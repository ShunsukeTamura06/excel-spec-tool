<script setup lang="ts">
/**
 * Vue Flow のカスタムノード.
 * 種別 (sheet / module / procedure) と meta に応じて表示を切り替える.
 */

import { Handle, Position, type NodeProps } from '@vue-flow/core'

interface NodeData {
  label: string
  kind: 'sheet' | 'module' | 'procedure'
  meta: Record<string, string | number>
}

const props = defineProps<NodeProps<NodeData>>()

const iconFor: Record<NodeData['kind'], string> = {
  sheet: 'i-lucide-layout-grid',
  module: 'i-lucide-folder',
  procedure: 'i-lucide-zap',
}

const subtitleFor = computed(() => {
  const m = props.data.meta
  if (props.data.kind === 'sheet') {
    const parts: string[] = []
    if (typeof m.formulas === 'number') parts.push(`${m.formulas} 数式`)
    if (typeof m.rows === 'number' && m.rows > 0) parts.push(`${m.rows}×${m.cols}`)
    return parts.join(' ・ ')
  }
  if (props.data.kind === 'procedure') {
    const parts: string[] = []
    if (m.module) parts.push(String(m.module))
    if (m.kind) parts.push(String(m.kind))
    return parts.join(' ・ ')
  }
  return ''
})

const purpose = computed(() => {
  const p = props.data.meta?.purpose
  return typeof p === 'string' && p.length > 0 ? p : ''
})
</script>

<template>
  <div
    class="rounded-xl border-2 px-3 py-2 shadow-sm bg-(--ui-bg) hover:shadow-md transition-shadow"
    :class="[
      props.data.kind === 'sheet' && 'border-emerald-400 dark:border-emerald-600',
      props.data.kind === 'procedure' && 'border-sky-400 dark:border-sky-600',
      props.data.kind === 'module' && 'border-indigo-400 dark:border-indigo-600',
    ]"
  >
    <Handle type="target" :position="Position.Left" class="!bg-transparent !border-(--ui-text-muted)" />
    <div class="flex items-center gap-2">
      <div
        class="size-7 rounded-lg flex items-center justify-center shrink-0"
        :class="[
          props.data.kind === 'sheet' && 'bg-emerald-100 dark:bg-emerald-950 text-emerald-700 dark:text-emerald-300',
          props.data.kind === 'procedure' && 'bg-sky-100 dark:bg-sky-950 text-sky-700 dark:text-sky-300',
          props.data.kind === 'module' && 'bg-indigo-100 dark:bg-indigo-950 text-indigo-700 dark:text-indigo-300',
        ]"
      >
        <UIcon :name="iconFor[props.data.kind]" class="size-4" />
      </div>
      <div class="min-w-0 flex-1">
        <p class="text-sm font-semibold text-(--ui-text-highlighted) truncate" :title="props.data.label">
          {{ props.data.label }}
        </p>
        <p v-if="subtitleFor" class="text-[10px] text-(--ui-text-muted) truncate">{{ subtitleFor }}</p>
      </div>
    </div>
    <p
      v-if="purpose"
      class="mt-1 text-[10px] text-(--ui-text-muted) line-clamp-2"
      :title="purpose"
    >
      {{ purpose }}
    </p>
    <Handle type="source" :position="Position.Right" class="!bg-transparent !border-(--ui-text-muted)" />
  </div>
</template>
