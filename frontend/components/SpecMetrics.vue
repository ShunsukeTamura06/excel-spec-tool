<script setup lang="ts">
import type { WorkbookData } from '~/types/api'

const props = defineProps<{ workbook: WorkbookData }>()

const stats = computed(() => {
  const sheets = props.workbook.sheets
  const modules = props.workbook.vba_modules
  return {
    sheets: sheets.length,
    formulas: sheets.reduce((sum, s) => sum + s.formulas.length, 0),
    named: sheets.reduce((sum, s) => sum + s.named_ranges.length, 0),
    cond: sheets.reduce((sum, s) => sum + s.conditional_formats.length, 0),
    modules: modules.length,
    procedures: modules.reduce((sum, m) => sum + m.procedures.length, 0),
    external: props.workbook.external_links.length,
  }
})

const tiles = computed(() => [
  { label: 'シート',         value: stats.value.sheets,     icon: 'i-lucide-layout-grid',    color: 'text-indigo-600 bg-indigo-50 dark:bg-indigo-950 dark:text-indigo-300' },
  { label: '数式セル',       value: stats.value.formulas,   icon: 'i-lucide-function-square', color: 'text-emerald-600 bg-emerald-50 dark:bg-emerald-950 dark:text-emerald-300' },
  { label: '名前付き範囲',   value: stats.value.named,      icon: 'i-lucide-tag',            color: 'text-amber-600 bg-amber-50 dark:bg-amber-950 dark:text-amber-300' },
  { label: '条件付き書式',   value: stats.value.cond,       icon: 'i-lucide-paintbrush',     color: 'text-purple-600 bg-purple-50 dark:bg-purple-950 dark:text-purple-300' },
  { label: 'VBA モジュール', value: stats.value.modules,    icon: 'i-lucide-folder',         color: 'text-sky-600 bg-sky-50 dark:bg-sky-950 dark:text-sky-300' },
  { label: 'プロシージャ',   value: stats.value.procedures, icon: 'i-lucide-zap',            color: 'text-rose-600 bg-rose-50 dark:bg-rose-950 dark:text-rose-300' },
  { label: '外部リンク',     value: stats.value.external,   icon: 'i-lucide-link',           color: 'text-slate-600 bg-slate-100 dark:bg-slate-800 dark:text-slate-300' },
])
</script>

<template>
  <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-7 gap-2">
    <div
      v-for="t in tiles"
      :key="t.label"
      class="rounded-xl px-3 py-3 flex flex-col items-start gap-1 transition-transform hover:scale-[1.02]"
      :class="t.color"
    >
      <UIcon :name="t.icon" class="size-4" />
      <span class="text-2xl font-bold tabular-nums leading-none">{{ t.value }}</span>
      <span class="text-[10px] uppercase tracking-wide opacity-80">{{ t.label }}</span>
    </div>
  </div>
</template>
