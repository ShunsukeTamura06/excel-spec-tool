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
          <UBadge
            v-if="selected.purpose || selected.usage_scenario || selected.inputs.length || selected.outputs.length || selected.main_calculations.length"
            color="primary"
            variant="subtle"
            icon="i-lucide-sparkles"
          >
            LLM 推定
          </UBadge>
        </div>
        <p v-if="selected.purpose" class="mt-2 text-sm text-(--ui-text)">{{ selected.purpose }}</p>
        <p v-if="selected.usage_scenario" class="mt-1 text-xs text-(--ui-text-muted)">
          <UIcon name="i-lucide-user-check" class="size-3.5 inline align-text-bottom" />
          想定利用シーン: {{ selected.usage_scenario }}
        </p>

        <!-- IN / OUT バッジ -->
        <div
          v-if="selected.inputs.length || selected.outputs.length"
          class="mt-3 flex flex-wrap items-start gap-3"
        >
          <div v-if="selected.inputs.length" class="flex items-center gap-1.5 flex-wrap">
            <span class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold">IN</span>
            <UBadge
              v-for="(it, i) in selected.inputs"
              :key="`in-${i}`"
              color="info"
              variant="subtle"
              size="sm"
              icon="i-lucide-log-in"
            >
              {{ it }}
            </UBadge>
          </div>
          <div v-if="selected.outputs.length" class="flex items-center gap-1.5 flex-wrap">
            <span class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold">OUT</span>
            <UBadge
              v-for="(it, i) in selected.outputs"
              :key="`out-${i}`"
              color="success"
              variant="subtle"
              size="sm"
              icon="i-lucide-log-out"
            >
              {{ it }}
            </UBadge>
          </div>
        </div>

        <!-- 主要計算 (LLM 説明) -->
        <div v-if="selected.main_calculations.length" class="mt-3">
          <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1">
            主要計算
          </p>
          <ul class="space-y-0.5 text-xs">
            <li
              v-for="(c, i) in selected.main_calculations"
              :key="`calc-${i}`"
              class="flex items-start gap-1.5"
            >
              <UIcon name="i-lucide-corner-down-right" class="size-3.5 shrink-0 mt-0.5 text-(--ui-text-muted)" />
              <span>{{ c }}</span>
            </li>
          </ul>
        </div>
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

      <!-- フォームコントロール (ボタン → マクロ) -->
      <UCard v-if="selected.form_controls.length">
        <template #header>
          <div class="flex items-center gap-2">
            <UIcon name="i-lucide-mouse-pointer-click" class="size-4 text-rose-600" />
            <span class="font-medium">フォームコントロール ({{ selected.form_controls.length }} 件)</span>
          </div>
        </template>
        <div class="overflow-x-auto -mx-4">
          <table class="w-full text-xs">
            <thead class="bg-(--ui-bg-elevated)">
              <tr>
                <th class="px-3 py-2 text-left font-semibold w-24">種別</th>
                <th class="px-3 py-2 text-left font-semibold">表示テキスト</th>
                <th class="px-3 py-2 text-left font-semibold w-20">配置</th>
                <th class="px-3 py-2 text-left font-semibold">紐づけマクロ</th>
              </tr>
            </thead>
            <tbody>
              <tr
                v-for="(fc, i) in selected.form_controls"
                :key="`fc-${i}`"
                class="border-t border-(--ui-border)"
              >
                <td class="px-3 py-1.5">
                  <UBadge color="neutral" variant="subtle" size="sm">{{ fc.kind }}</UBadge>
                </td>
                <td class="px-3 py-1.5">{{ fc.text || '-' }}</td>
                <td class="px-3 py-1.5 font-mono text-[10px]">{{ fc.anchor || '-' }}</td>
                <td class="px-3 py-1.5 font-mono">
                  <span v-if="fc.macro">{{ fc.macro }}</span>
                  <span v-else class="text-(--ui-text-muted)">-</span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </UCard>

      <!-- データ検証 (入力規則) -->
      <UCard v-if="selected.data_validations.length">
        <template #header>
          <div class="flex items-center gap-2">
            <UIcon name="i-lucide-list-checks" class="size-4 text-sky-600" />
            <span class="font-medium">入力規則 ({{ selected.data_validations.length }} 件)</span>
          </div>
        </template>
        <div class="overflow-x-auto -mx-4">
          <table class="w-full text-xs">
            <thead class="bg-(--ui-bg-elevated)">
              <tr>
                <th class="px-3 py-2 text-left font-semibold w-32">範囲</th>
                <th class="px-3 py-2 text-left font-semibold w-20">種別</th>
                <th class="px-3 py-2 text-left font-semibold">値 / 数式</th>
                <th class="px-3 py-2 text-left font-semibold">プロンプト</th>
              </tr>
            </thead>
            <tbody>
              <tr
                v-for="(dv, i) in selected.data_validations"
                :key="`dv-${i}`"
                class="border-t border-(--ui-border)"
              >
                <td class="px-3 py-1.5 font-mono">{{ dv.range }}</td>
                <td class="px-3 py-1.5">
                  <UBadge color="info" variant="subtle" size="sm">{{ dv.type }}</UBadge>
                </td>
                <td class="px-3 py-1.5 font-mono break-all">
                  {{ dv.formula || '-' }}<span v-if="dv.operator"> ({{ dv.operator }})</span>
                </td>
                <td class="px-3 py-1.5">{{ dv.prompt || '-' }}</td>
              </tr>
            </tbody>
          </table>
        </div>
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
