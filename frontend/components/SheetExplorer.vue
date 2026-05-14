<script setup lang="ts">
/**
 * シートタブ. 左にシート選択リスト, 右に詳細パネル.
 * 数式 / 名前付き範囲 / 条件付き書式 / プレビュー / Excel テーブル を表示する.
 */

import type { SheetInfo } from '~/types/api'

const props = defineProps<{ sheets: SheetInfo[] }>()
const emit = defineEmits<{ searchReference: [target: string] }>()

const selectedName = ref(props.sheets[0]?.name ?? '')
watch(
  () => props.sheets.map(s => s.name).join('|'),
  () => {
    if (!props.sheets.some(s => s.name === selectedName.value)) {
      selectedName.value = props.sheets[0]?.name ?? ''
    }
  },
)
const selected = computed<SheetInfo | null>(
  () => props.sheets.find(s => s.name === selectedName.value) ?? null,
)
</script>

<template>
  <div v-if="props.sheets.length === 0" class="rounded-2xl border border-dashed border-(--ui-border) p-10 text-center">
    <UIcon name="i-lucide-file-x" class="size-8 mx-auto text-(--ui-text-muted) mb-2" />
    <p class="text-sm text-(--ui-text-muted)">シートがありません (.xls の場合 openpyxl で読めない可能性あり)</p>
  </div>
  <div v-else class="grid grid-cols-1 lg:grid-cols-[220px_1fr] gap-4">
    <!-- 左: シート選択 -->
    <UCard :ui="{ body: 'p-2' }">
      <ul class="space-y-0.5 max-h-[640px] overflow-y-auto">
        <li v-for="sheet in props.sheets" :key="sheet.name">
          <button
            type="button"
            class="w-full text-left px-3 py-2 rounded-md transition-colors flex items-center justify-between gap-2"
            :class="
              selectedName === sheet.name
                ? 'bg-(--ui-primary)/10 text-(--ui-primary)'
                : 'hover:bg-(--ui-bg-muted) text-(--ui-text)'
            "
            @click="selectedName = sheet.name"
          >
            <span class="text-sm font-medium truncate">{{ sheet.name }}</span>
            <UBadge
              v-if="sheet.formulas.length > 0"
              :color="selectedName === sheet.name ? 'primary' : 'neutral'"
              variant="subtle"
              size="sm"
              class="tabular-nums"
            >
              {{ sheet.formulas.length }}
            </UBadge>
          </button>
        </li>
      </ul>
    </UCard>

    <!-- 右: 詳細 -->
    <div v-if="selected" class="space-y-3">
      <UCard>
        <div class="flex items-baseline justify-between gap-2 flex-wrap">
          <div>
            <h3 class="text-lg font-semibold text-(--ui-text-highlighted)">{{ selected.name }}</h3>
            <p class="text-xs text-(--ui-text-muted) mt-0.5">
              {{ selected.rows }} 行 × {{ selected.cols }} 列
              <span v-if="selected.tables.length" class="ml-1">・ テーブル {{ selected.tables.length }} 件</span>
              <span v-if="selected.merged_ranges.length" class="ml-1">・ 結合セル {{ selected.merged_ranges.length }} 件</span>
            </p>
          </div>
          <UBadge v-if="selected.purpose" color="primary" variant="subtle" icon="i-lucide-sparkles">
            LLM 推定
          </UBadge>
        </div>
        <p v-if="selected.purpose" class="mt-2 text-sm text-(--ui-text)">{{ selected.purpose }}</p>
      </UCard>

      <!-- 数式 -->
      <UCard v-if="selected.formulas.length">
        <template #header>
          <div class="flex items-center gap-2">
            <UIcon name="i-lucide-function-square" class="size-4 text-(--ui-primary)" />
            <span class="font-medium">数式 ({{ selected.formulas.length }} 件)</span>
          </div>
        </template>
        <div class="max-h-[420px] overflow-y-auto -mx-4">
          <table class="w-full text-xs">
            <thead class="bg-(--ui-bg-elevated) sticky top-0">
              <tr>
                <th class="px-3 py-2 text-left font-semibold w-20">セル</th>
                <th class="px-3 py-2 text-left font-semibold">数式</th>
                <th class="px-3 py-2 text-left font-semibold w-48">参照先</th>
              </tr>
            </thead>
            <tbody>
              <tr
                v-for="(f, idx) in selected.formulas.slice(0, 200)"
                :key="`${f.coord}-${idx}`"
                class="border-t border-(--ui-border) hover:bg-(--ui-bg-muted)/40"
              >
                <td class="px-3 py-1.5 font-mono">
                  <button
                    class="text-(--ui-primary) hover:underline"
                    @click="emit('searchReference', `${selected!.name}!${f.coord}`)"
                  >
                    {{ f.coord }}
                  </button>
                </td>
                <td class="px-3 py-1.5 font-mono break-all">{{ f.formula }}</td>
                <td class="px-3 py-1.5">
                  <div class="flex flex-wrap gap-1">
                    <UBadge
                      v-for="(r, i) in f.refs.slice(0, 4)"
                      :key="i"
                      size="sm"
                      color="neutral"
                      variant="subtle"
                      class="font-mono text-[10px]"
                    >
                      {{ r }}
                    </UBadge>
                    <span v-if="f.refs.length > 4" class="text-[10px] text-(--ui-text-muted)">+{{ f.refs.length - 4 }}</span>
                  </div>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
        <p v-if="selected.formulas.length > 200" class="mt-2 text-xs text-(--ui-text-muted) text-center">
          (先頭 200 件のみ表示)
        </p>
      </UCard>

      <!-- 名前付き範囲 + 条件付き書式 -->
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <UCard v-if="selected.named_ranges.length">
          <template #header>
            <div class="flex items-center gap-2">
              <UIcon name="i-lucide-tag" class="size-4 text-emerald-600" />
              <span class="font-medium">名前付き範囲</span>
            </div>
          </template>
          <ul class="space-y-1 text-xs">
            <li
              v-for="(n, i) in selected.named_ranges"
              :key="i"
              class="flex items-center justify-between gap-2"
            >
              <span class="font-medium">{{ n.name }}</span>
              <code class="font-mono text-[10px] text-(--ui-text-muted) truncate">{{ n.refers_to }}</code>
            </li>
          </ul>
        </UCard>

        <UCard v-if="selected.conditional_formats.length">
          <template #header>
            <div class="flex items-center gap-2">
              <UIcon name="i-lucide-paintbrush" class="size-4 text-purple-600" />
              <span class="font-medium">条件付き書式</span>
            </div>
          </template>
          <ul class="space-y-1.5 text-xs">
            <li v-for="(c, i) in selected.conditional_formats" :key="i">
              <code class="font-mono text-[10px] mr-2">{{ c.range }}</code>
              <span class="text-(--ui-text-muted)">{{ c.rule }}</span>
            </li>
          </ul>
        </UCard>
      </div>

      <!-- プレビュー -->
      <UCard v-if="selected.preview_rows.length">
        <template #header>
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <UIcon name="i-lucide-eye" class="size-4 text-sky-600" />
              <span class="font-medium">プレビュー</span>
            </div>
            <span class="text-[10px] font-mono text-(--ui-text-muted)">起点: {{ selected.preview_origin }}</span>
          </div>
        </template>
        <div class="overflow-x-auto -mx-4">
          <table class="w-full text-xs">
            <tbody>
              <tr
                v-for="(row, ri) in selected.preview_rows"
                :key="ri"
                class="border-t border-(--ui-border)"
                :class="ri === 0 && 'bg-(--ui-bg-elevated) font-medium'"
              >
                <td
                  v-for="(cell, ci) in row"
                  :key="ci"
                  class="px-2 py-1 whitespace-nowrap"
                  :class="cell == null && 'text-(--ui-text-muted)/40'"
                >
                  {{ cell ?? '·' }}
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </UCard>
    </div>
  </div>
</template>
