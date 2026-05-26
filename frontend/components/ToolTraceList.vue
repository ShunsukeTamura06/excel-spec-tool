<script setup lang="ts">
/**
 * 直近応答のツール結果を、LLM の文章ではなく検証可能な証拠カードとして描画する.
 */

import type { ToolTraceItem } from '~/types/api'

type CellValue = {
  value?: string | null
  formula?: string | null
  data_type?: string | null
}

type CellRangeResult = {
  sheet?: string
  range?: string
  origin_row?: number
  origin_col?: number
  rows?: (CellValue | null)[][]
}

type FindCellsResult = {
  matches?: Array<Record<string, unknown>>
  count?: number
}

type ReferencesResult = {
  refs?: Array<Record<string, unknown>>
  count?: number
  analysis_scope?: string
}

type RisksResult = {
  risks?: Array<Record<string, unknown>>
  counts?: Record<string, number>
  total?: number
  returned?: number
  analysis_scope?: string
}

type WorkbookObjectsResult = {
  charts?: Array<Record<string, unknown>>
  pivot_tables?: Array<Record<string, unknown>>
  power_queries?: Array<Record<string, unknown>>
  counts?: Record<string, number>
  analysis_scope?: string
}

type FormulaListResult = {
  sheet?: string
  formulas?: Array<Record<string, unknown>>
  total?: number
  returned?: number
  truncated?: boolean
}

type VbaProcedureResult = {
  module?: string
  name?: string
  kind?: string
  start_line?: number
  end_line?: number
  code?: string
  annotation?: string
}

const props = withDefaults(defineProps<{
  items: ToolTraceItem[]
  title?: string
}>(), {
  title: '根拠カード',
})

const open = ref(true)

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function asRecord(value: unknown): Record<string, unknown> {
  return isRecord(value) ? value : {}
}

function asCellRange(value: unknown): CellRangeResult {
  return asRecord(value) as CellRangeResult
}

function asFindCells(value: unknown): FindCellsResult {
  return asRecord(value) as FindCellsResult
}

function asReferences(value: unknown): ReferencesResult {
  return asRecord(value) as ReferencesResult
}

function asRisks(value: unknown): RisksResult {
  return asRecord(value) as RisksResult
}

function asWorkbookObjects(value: unknown): WorkbookObjectsResult {
  return asRecord(value) as WorkbookObjectsResult
}

function asFormulaList(value: unknown): FormulaListResult {
  return asRecord(value) as FormulaListResult
}

function asVbaProcedure(value: unknown): VbaProcedureResult {
  return asRecord(value) as VbaProcedureResult
}

function formatArgs(args: Record<string, unknown>): string {
  try {
    return JSON.stringify(args, null, 2)
  } catch {
    return String(args)
  }
}

function formatValue(value: unknown): string {
  if (value == null || value === '') return '-'
  if (Array.isArray(value)) return value.join(', ')
  return String(value)
}

function cellText(cell: CellValue | null | undefined): string {
  if (!cell) return ''
  return cell.formula || cell.value || ''
}

function evidenceTitle(item: ToolTraceItem): string {
  switch (item.name) {
    case 'get_cells_range':
      return 'セル範囲'
    case 'find_cells':
      return 'セル検索'
    case 'lookup_references':
      return '参照関係'
    case 'get_vba_procedure':
      return 'VBAコード'
    case 'list_sheet_formulas':
      return '数式一覧'
    case 'list_analysis_risks':
      return '未解析リスク'
    case 'list_workbook_objects':
      return 'Excelオブジェクト'
    case 'list_external_functions_used':
      return '外部関数の使用箇所'
    case 'lookup_external_function':
      return '外部関数定義'
    default:
      return item.name
  }
}

function evidenceIcon(item: ToolTraceItem): string {
  switch (item.name) {
    case 'get_cells_range':
    case 'find_cells':
      return 'i-lucide-grid-3x3'
    case 'lookup_references':
      return 'i-lucide-git-branch'
    case 'get_vba_procedure':
      return 'i-lucide-code-2'
    case 'list_analysis_risks':
      return 'i-lucide-triangle-alert'
    case 'list_workbook_objects':
      return 'i-lucide-chart-column'
    case 'list_sheet_formulas':
      return 'i-lucide-function-square'
    default:
      return 'i-lucide-database'
  }
}

function riskColor(severity: unknown): 'error' | 'warning' | 'neutral' {
  if (severity === 'high') return 'error'
  if (severity === 'medium') return 'warning'
  return 'neutral'
}

function codeLines(code: string | undefined, limit = 80): string {
  if (!code) return ''
  const lines = code.split('\n')
  const shown = lines.slice(0, limit).join('\n')
  return lines.length > limit ? `${shown}\n... (${lines.length - limit} 行省略)` : shown
}
</script>

<template>
  <div v-if="props.items.length > 0" class="rounded-lg border border-(--ui-border) bg-(--ui-bg-elevated)/40 overflow-hidden">
    <button
      type="button"
      class="w-full flex items-center justify-between gap-2 px-3 py-2 text-xs text-(--ui-text-muted) hover:text-(--ui-text) transition-colors"
      @click="open = !open"
    >
      <span class="flex items-center gap-2">
        <UIcon name="i-lucide-clipboard-check" class="size-3.5" />
        {{ props.title }} ({{ props.items.length }} 件)
      </span>
      <UIcon :name="open ? 'i-lucide-chevron-up' : 'i-lucide-chevron-down'" class="size-3.5" />
    </button>

    <div
      v-if="open"
      class="border-t border-(--ui-border) p-3 space-y-3 max-h-[min(32rem,60vh)] overflow-y-auto overscroll-contain"
    >
      <div
        v-for="(t, i) in props.items"
        :key="i"
        class="rounded-md border border-(--ui-border) bg-(--ui-bg) overflow-hidden"
      >
        <div class="px-3 py-2 border-b border-(--ui-border) flex items-center justify-between gap-2">
          <div class="flex items-center gap-2 min-w-0">
            <UIcon :name="evidenceIcon(t)" class="size-4 text-(--ui-primary) shrink-0" />
            <div class="min-w-0">
              <p class="text-sm font-medium text-(--ui-text-highlighted) truncate">
                {{ evidenceTitle(t) }}
              </p>
              <p class="text-[10px] font-mono text-(--ui-text-muted) truncate">{{ t.name }}</p>
            </div>
          </div>
          <details class="shrink-0 text-right">
            <summary class="cursor-pointer text-[11px] text-(--ui-text-muted) hover:text-(--ui-text)">
              引数
            </summary>
            <pre class="mt-1 p-2 rounded bg-(--ui-bg-elevated) overflow-x-auto text-[10px] font-mono text-left max-w-md">{{ formatArgs(t.arguments) }}</pre>
          </details>
        </div>

        <div class="p-3">
          <!-- get_cells_range -->
          <div v-if="t.name === 'get_cells_range'" class="space-y-2">
            <p class="text-xs text-(--ui-text-muted)">
              <code>{{ asCellRange(t.result).sheet }}</code>!<code>{{ asCellRange(t.result).range }}</code>
            </p>
            <div class="overflow-x-auto">
              <table class="text-xs border-collapse min-w-full">
                <tbody>
                  <tr
                    v-for="(row, ri) in asCellRange(t.result).rows ?? []"
                    :key="ri"
                    class="border-t border-(--ui-border)"
                  >
                    <td
                      v-for="(cell, ci) in row"
                      :key="ci"
                      class="px-2 py-1 border-r border-(--ui-border) whitespace-nowrap font-mono max-w-56 truncate"
                      :class="cell?.formula ? 'text-(--ui-primary)' : 'text-(--ui-text)'"
                    >
                      {{ cellText(cell) || '·' }}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <!-- find_cells -->
          <div v-else-if="t.name === 'find_cells'" class="space-y-2">
            <p class="text-xs text-(--ui-text-muted)">検索結果: {{ asFindCells(t.result).count ?? 0 }} 件</p>
            <div class="overflow-x-auto">
              <table class="w-full text-xs">
                <thead class="bg-(--ui-bg-elevated)">
                  <tr>
                    <th class="px-2 py-1 text-left">シート</th>
                    <th class="px-2 py-1 text-left">セル</th>
                    <th class="px-2 py-1 text-left">値</th>
                  </tr>
                </thead>
                <tbody>
                  <tr
                    v-for="(m, mi) in (asFindCells(t.result).matches ?? []).slice(0, 20)"
                    :key="mi"
                    class="border-t border-(--ui-border)"
                  >
                    <td class="px-2 py-1">{{ formatValue(m.sheet) }}</td>
                    <td class="px-2 py-1 font-mono">{{ formatValue(m.coord) }}</td>
                    <td class="px-2 py-1">{{ formatValue(m.value) }}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <!-- lookup_references -->
          <div v-else-if="t.name === 'lookup_references'" class="space-y-2">
            <p class="text-xs text-(--ui-text-muted)">参照元: {{ asReferences(t.result).count ?? 0 }} 件</p>
            <div class="overflow-x-auto">
              <table class="w-full text-xs">
                <thead class="bg-(--ui-bg-elevated)">
                  <tr>
                    <th class="px-2 py-1 text-left">種別</th>
                    <th class="px-2 py-1 text-left">参照元</th>
                    <th class="px-2 py-1 text-left">コード/内容</th>
                  </tr>
                </thead>
                <tbody>
                  <tr
                    v-for="(r, ri) in (asReferences(t.result).refs ?? []).slice(0, 30)"
                    :key="ri"
                    class="border-t border-(--ui-border)"
                  >
                    <td class="px-2 py-1"><UBadge size="sm" variant="subtle">{{ formatValue(r.kind) }}</UBadge></td>
                    <td class="px-2 py-1 font-mono">{{ formatValue(r.from) }}</td>
                    <td class="px-2 py-1 font-mono break-all">{{ formatValue(r.code) }}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <p v-if="asReferences(t.result).analysis_scope" class="text-[11px] text-(--ui-text-muted)">
              {{ asReferences(t.result).analysis_scope }}
            </p>
          </div>

          <!-- get_vba_procedure -->
          <div v-else-if="t.name === 'get_vba_procedure'" class="space-y-2">
            <p class="text-xs text-(--ui-text-muted)">
              <code>{{ asVbaProcedure(t.result).module }}</code>.<code>{{ asVbaProcedure(t.result).name }}</code>
              L{{ asVbaProcedure(t.result).start_line }}-{{ asVbaProcedure(t.result).end_line }}
            </p>
            <pre class="max-h-80 overflow-auto rounded bg-(--ui-bg-elevated) p-3 text-[11px] font-mono leading-relaxed">{{ codeLines(asVbaProcedure(t.result).code) }}</pre>
          </div>

          <!-- list_sheet_formulas -->
          <div v-else-if="t.name === 'list_sheet_formulas'" class="space-y-2">
            <p class="text-xs text-(--ui-text-muted)">
              {{ asFormulaList(t.result).sheet }}: {{ asFormulaList(t.result).returned ?? 0 }} / {{ asFormulaList(t.result).total ?? 0 }} 件
            </p>
            <div class="space-y-1">
              <div
                v-for="(f, fi) in (asFormulaList(t.result).formulas ?? []).slice(0, 20)"
                :key="fi"
                class="rounded bg-(--ui-bg-elevated) px-2 py-1 text-xs"
              >
                <code>{{ formatValue(f.coord) }}</code>
                <span class="ml-2 font-mono break-all">{{ formatValue(f.formula) }}</span>
              </div>
            </div>
          </div>

          <!-- list_analysis_risks -->
          <div v-else-if="t.name === 'list_analysis_risks'" class="space-y-2">
            <div class="flex flex-wrap gap-1">
              <UBadge
                v-for="(count, key) in asRisks(t.result).counts ?? {}"
                :key="key"
                :color="riskColor(key)"
                variant="subtle"
                size="sm"
              >
                {{ key }}: {{ count }}
              </UBadge>
            </div>
            <div class="space-y-2">
              <div
                v-for="(risk, ri) in (asRisks(t.result).risks ?? []).slice(0, 20)"
                :key="ri"
                class="rounded border border-(--ui-border) p-2 text-xs"
              >
                <div class="flex items-center gap-2 mb-1">
                  <UBadge :color="riskColor(risk.severity)" variant="subtle" size="sm">
                    {{ formatValue(risk.severity) }}
                  </UBadge>
                  <span class="font-mono text-(--ui-text-muted)">{{ formatValue(risk.location) }}</span>
                </div>
                <p>{{ formatValue(risk.description) }}</p>
                <p class="mt-1 text-(--ui-text-muted)">確認: {{ formatValue(risk.recommendation) }}</p>
              </div>
            </div>
          </div>

          <!-- list_workbook_objects -->
          <div v-else-if="t.name === 'list_workbook_objects'" class="space-y-3 text-xs">
            <div class="flex flex-wrap gap-1">
              <UBadge
                v-for="(count, key) in asWorkbookObjects(t.result).counts ?? {}"
                :key="key"
                color="neutral"
                variant="subtle"
                size="sm"
              >
                {{ key }}: {{ count }}
              </UBadge>
            </div>
            <div v-if="(asWorkbookObjects(t.result).charts ?? []).length" class="space-y-1">
              <p class="font-medium">グラフ</p>
              <div v-for="(chart, ci) in (asWorkbookObjects(t.result).charts ?? []).slice(0, 10)" :key="ci">
                {{ formatValue(chart.sheet) }} / {{ formatValue(chart.title || chart.name || chart.chart_type) }}
              </div>
            </div>
            <div v-if="(asWorkbookObjects(t.result).pivot_tables ?? []).length" class="space-y-1">
              <p class="font-medium">ピボット</p>
              <div v-for="(pivot, pi) in (asWorkbookObjects(t.result).pivot_tables ?? []).slice(0, 10)" :key="pi">
                {{ formatValue(pivot.sheet) }} / {{ formatValue(pivot.name) }} / 元データ {{ formatValue(pivot.source_sheet) }}!{{ formatValue(pivot.source_ref) }}
              </div>
            </div>
            <div v-if="(asWorkbookObjects(t.result).power_queries ?? []).length" class="space-y-1">
              <p class="font-medium">Power Query / 外部接続</p>
              <div v-for="(query, qi) in (asWorkbookObjects(t.result).power_queries ?? []).slice(0, 10)" :key="qi">
                {{ formatValue(query.name) }} → {{ formatValue(query.target_sheet || query.target_name) }}
              </div>
            </div>
          </div>

          <!-- fallback -->
          <details v-else>
            <summary class="cursor-pointer text-xs text-(--ui-text-muted) hover:text-(--ui-text)">
              結果プレビュー
            </summary>
            <pre class="mt-2 p-2 rounded bg-(--ui-bg-elevated) overflow-x-auto text-[10px] font-mono whitespace-pre-wrap">{{ t.result_preview }}</pre>
          </details>
        </div>
      </div>
    </div>
  </div>
</template>
