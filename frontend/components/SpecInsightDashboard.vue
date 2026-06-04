<script setup lang="ts">
/**
 * 概要タブの「最初の 10 秒で見たいもの」ダッシュボード.
 *
 * Excel に困って到達したユーザーが、ファイル全体を読まずに「このブックは
 * 何で、計算の中心はどこで、危ない箇所はどこか」を即把握できることを目標とする.
 * Markdown spec.md は別途折りたたみで残す.
 *
 * 入力は /workbook の構造データのみ. 派生指標は全てここで計算する.
 */
import type { WorkbookData, AnalysisRiskItem } from '~/types/api'

const props = defineProps<{ workbook: WorkbookData }>()

interface SheetEdge {
  from: string
  to: string
  weight: number
}

// Excel の参照表記 `[Sheet!]A1[:B2]` から「シート修飾」だけ取り出す.
// シート修飾が無い場合は null (= 同一シート内参照とみなす).
function refToSheetName(ref: string): string | null {
  // `'My Sheet'!A1` (クォート付き) と `Sheet1!A1` の両形態に対応
  const m = ref.match(/^(?:'((?:[^']|'')+)'|([^'!]+))!/)
  if (!m) return null
  return (m[1] ?? m[2] ?? '').replace(/''/g, "'").trim() || null
}

// シート間参照を (from, to, weight) に集約する.
// 同一シート内参照は除外 (自己ループはダッシュボードのノイズになるだけ).
const sheetEdges = computed<SheetEdge[]>(() => {
  const sheetNames = new Set(props.workbook.sheets.map((s) => s.name))
  const counter = new Map<string, number>()
  for (const sheet of props.workbook.sheets) {
    for (const f of sheet.formulas) {
      for (const ref of f.refs) {
        const target = refToSheetName(ref)
        if (!target || target === sheet.name || !sheetNames.has(target)) continue
        const key = `${sheet.name}${target}`
        counter.set(key, (counter.get(key) ?? 0) + 1)
      }
    }
  }
  return Array.from(counter.entries()).map(([key, weight]) => {
    const [from, to] = key.split('')
    return { from, to, weight }
  })
})

interface SheetRank {
  name: string
  value: number
  hint: string
}

// 数式数 TOP 3 (= 計算の中心候補)
const formulaRanking = computed<SheetRank[]>(() => {
  return [...props.workbook.sheets]
    .filter((s) => s.formulas.length > 0)
    .sort((a, b) => b.formulas.length - a.formulas.length)
    .slice(0, 3)
    .map((s) => ({
      name: s.name,
      value: s.formulas.length,
      hint: `${s.rows}×${s.cols}`,
    }))
})

// 他シートから「最も読まれている」シート TOP 3 (= ハブシート / 入力源)
const incomingRanking = computed<SheetRank[]>(() => {
  const incoming = new Map<string, number>()
  for (const e of sheetEdges.value) {
    incoming.set(e.to, (incoming.get(e.to) ?? 0) + e.weight)
  }
  return [...incoming.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([name, value]) => ({
      name,
      value,
      hint: `${value} 参照`,
    }))
})

// 他シートを「最も多く読む」シート TOP 3 (= 出力 / 集計側候補)
const outgoingRanking = computed<SheetRank[]>(() => {
  const outgoing = new Map<string, number>()
  for (const e of sheetEdges.value) {
    outgoing.set(e.from, (outgoing.get(e.from) ?? 0) + e.weight)
  }
  return [...outgoing.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3)
    .map(([name, value]) => ({
      name,
      value,
      hint: `${value} 参照`,
    }))
})

// TL;DR — 1〜2 文で「このブックは何か」を要約.
// LLM 注釈に依存せず、構造から取り出す.
const tldr = computed(() => {
  const wb = props.workbook
  const sheetCount = wb.sheets.length
  const formulaTotal = wb.sheets.reduce((sum, s) => sum + s.formulas.length, 0)
  const vbaProcCount = wb.vba_modules.reduce(
    (sum, m) => sum + m.procedures.length,
    0,
  )
  const parts: string[] = []
  parts.push(`${sheetCount} シート構成`)
  if (formulaTotal > 0) parts.push(`数式 ${formulaTotal} 件`)
  if (wb.vba_modules.length > 0) {
    parts.push(`VBA ${wb.vba_modules.length} モジュール (${vbaProcCount} プロシージャ)`)
  } else {
    parts.push('VBA なし')
  }
  if (wb.power_queries.length > 0) {
    parts.push(`Power Query / 外部接続 ${wb.power_queries.length} 件`)
  }
  if (wb.external_links.length > 0) {
    parts.push(`外部リンク ${wb.external_links.length} 件`)
  }
  return parts.join(' ・ ')
})

// 入口・出口の推定. 入ってくる参照だけがあり出ていく参照が無い ⇒ 集計シート (出口),
// 出ていく参照だけがあり入ってくる参照が無い ⇒ 入力シート, の経験則.
interface RoleHint {
  name: string
  reason: string
}

const inputCandidates = computed<RoleHint[]>(() => {
  const ins = new Set<string>()
  const outs = new Set<string>()
  for (const e of sheetEdges.value) {
    ins.add(e.to)
    outs.add(e.from)
  }
  return props.workbook.sheets
    .filter((s) => ins.has(s.name) && !outs.has(s.name))
    .slice(0, 3)
    .map((s) => ({
      name: s.name,
      reason: '他シートから参照されるだけで、自身は他シートを参照しない',
    }))
})

const outputCandidates = computed<RoleHint[]>(() => {
  const ins = new Set<string>()
  const outs = new Set<string>()
  for (const e of sheetEdges.value) {
    ins.add(e.to)
    outs.add(e.from)
  }
  return props.workbook.sheets
    .filter((s) => outs.has(s.name) && !ins.has(s.name))
    .slice(0, 3)
    .map((s) => ({
      name: s.name,
      reason: '他シートを参照するが、他シートからは参照されない',
    }))
})

// 未解析リスク サマリ
const risksBySeverity = computed(() => {
  const out: Record<AnalysisRiskItem['severity'], number> = {
    high: 0,
    medium: 0,
    low: 0,
  }
  for (const r of props.workbook.analysis_risks) out[r.severity] += 1
  return out
})

const topRisks = computed<AnalysisRiskItem[]>(() => {
  const order = { high: 0, medium: 1, low: 2 } as const
  return [...props.workbook.analysis_risks]
    .sort((a, b) => order[a.severity] - order[b.severity])
    .slice(0, 5)
})

const riskCategoryLabel: Record<AnalysisRiskItem['category'], string> = {
  dynamic_vba: '動的 VBA 参照',
  runtime_state: '実行時状態依存',
  dynamic_formula: '動的数式 (INDIRECT/OFFSET 等)',
  external_dependency: '外部依存',
  event_macro: 'イベントマクロ',
  unknown_object_dependency: 'オブジェクト依存不明',
}

// UBadge の color prop が受け付ける union 型。string だと TS2322 になるため明示する。
type BadgeColor =
  | 'error'
  | 'warning'
  | 'primary'
  | 'secondary'
  | 'success'
  | 'info'
  | 'neutral'

const severityColor: Record<AnalysisRiskItem['severity'], BadgeColor> = {
  high: 'error',
  medium: 'warning',
  low: 'neutral',
}
</script>

<template>
  <div class="space-y-4">
    <!-- TL;DR カード -->
    <UCard>
      <template #header>
        <div class="flex items-center gap-2">
          <UIcon name="i-lucide-file-spreadsheet" class="size-4 text-(--ui-primary)" />
          <span class="font-semibold">{{ props.workbook.filename }}</span>
        </div>
      </template>
      <p class="text-sm leading-relaxed text-(--ui-text)">{{ tldr }}</p>

      <!-- 入口 / 出口候補 (構造から推定) -->
      <div
        v-if="inputCandidates.length || outputCandidates.length"
        class="mt-4 grid grid-cols-1 sm:grid-cols-2 gap-3"
      >
        <div v-if="inputCandidates.length">
          <p class="text-xs text-(--ui-text-muted) mb-1.5 flex items-center gap-1">
            <UIcon name="i-lucide-log-in" class="size-3.5" />
            入力候補シート
          </p>
          <div class="flex flex-wrap gap-1.5">
            <UBadge
              v-for="c in inputCandidates"
              :key="c.name"
              variant="soft"
              color="info"
              :title="c.reason"
            >
              {{ c.name }}
            </UBadge>
          </div>
        </div>
        <div v-if="outputCandidates.length">
          <p class="text-xs text-(--ui-text-muted) mb-1.5 flex items-center gap-1">
            <UIcon name="i-lucide-log-out" class="size-3.5" />
            出力候補シート
          </p>
          <div class="flex flex-wrap gap-1.5">
            <UBadge
              v-for="c in outputCandidates"
              :key="c.name"
              variant="soft"
              color="success"
              :title="c.reason"
            >
              {{ c.name }}
            </UBadge>
          </div>
        </div>
      </div>
    </UCard>

    <!-- シートランキング -->
    <div
      v-if="formulaRanking.length || incomingRanking.length || outgoingRanking.length"
      class="grid grid-cols-1 md:grid-cols-3 gap-3"
    >
      <UCard v-if="formulaRanking.length" class="text-sm">
        <template #header>
          <div class="flex items-center gap-2">
            <UIcon name="i-lucide-calculator" class="size-4 text-purple-500" />
            <span class="font-medium">計算が多いシート</span>
          </div>
        </template>
        <ol class="space-y-1.5">
          <li
            v-for="(s, i) in formulaRanking"
            :key="s.name"
            class="flex items-center justify-between gap-2"
          >
            <span class="flex items-center gap-2 min-w-0">
              <span class="text-(--ui-text-muted) text-xs w-4 shrink-0">{{ i + 1 }}</span>
              <span class="font-medium truncate">{{ s.name }}</span>
            </span>
            <span class="text-xs text-(--ui-text-muted) shrink-0">
              {{ s.value }} 数式 ・ {{ s.hint }}
            </span>
          </li>
        </ol>
      </UCard>

      <UCard v-if="incomingRanking.length" class="text-sm">
        <template #header>
          <div class="flex items-center gap-2">
            <UIcon name="i-lucide-target" class="size-4 text-emerald-500" />
            <span class="font-medium">よく参照されるシート</span>
          </div>
        </template>
        <ol class="space-y-1.5">
          <li
            v-for="(s, i) in incomingRanking"
            :key="s.name"
            class="flex items-center justify-between gap-2"
          >
            <span class="flex items-center gap-2 min-w-0">
              <span class="text-(--ui-text-muted) text-xs w-4 shrink-0">{{ i + 1 }}</span>
              <span class="font-medium truncate">{{ s.name }}</span>
            </span>
            <span class="text-xs text-(--ui-text-muted) shrink-0">{{ s.hint }}</span>
          </li>
        </ol>
      </UCard>

      <UCard v-if="outgoingRanking.length" class="text-sm">
        <template #header>
          <div class="flex items-center gap-2">
            <UIcon name="i-lucide-share-2" class="size-4 text-sky-500" />
            <span class="font-medium">最も読み込むシート</span>
          </div>
        </template>
        <ol class="space-y-1.5">
          <li
            v-for="(s, i) in outgoingRanking"
            :key="s.name"
            class="flex items-center justify-between gap-2"
          >
            <span class="flex items-center gap-2 min-w-0">
              <span class="text-(--ui-text-muted) text-xs w-4 shrink-0">{{ i + 1 }}</span>
              <span class="font-medium truncate">{{ s.name }}</span>
            </span>
            <span class="text-xs text-(--ui-text-muted) shrink-0">{{ s.hint }}</span>
          </li>
        </ol>
      </UCard>
    </div>

    <!-- 未解析リスク サマリ -->
    <UCard v-if="props.workbook.analysis_risks.length > 0">
      <template #header>
        <div class="flex items-center gap-2 flex-wrap">
          <UIcon name="i-lucide-shield-alert" class="size-4 text-amber-500" />
          <span class="font-medium">未解析リスク</span>
          <UBadge
            v-if="risksBySeverity.high > 0"
            variant="soft"
            color="error"
          >high {{ risksBySeverity.high }}</UBadge>
          <UBadge
            v-if="risksBySeverity.medium > 0"
            variant="soft"
            color="warning"
          >medium {{ risksBySeverity.medium }}</UBadge>
          <UBadge
            v-if="risksBySeverity.low > 0"
            variant="soft"
            color="neutral"
          >low {{ risksBySeverity.low }}</UBadge>
          <span class="text-xs text-(--ui-text-muted) ml-auto">
            合計 {{ props.workbook.analysis_risks.length }} 件
          </span>
        </div>
      </template>
      <p class="text-xs text-(--ui-text-muted) mb-2">
        静的解析では「影響なし」と断定できない箇所。改修時は手動確認が必要です。
      </p>
      <ul class="space-y-2">
        <li
          v-for="(r, i) in topRisks"
          :key="i"
          class="flex items-start gap-2 text-sm border-l-2 border-(--ui-border) pl-2"
        >
          <UBadge :color="severityColor[r.severity]" variant="subtle" size="xs">
            {{ r.severity }}
          </UBadge>
          <div class="min-w-0 flex-1">
            <p class="font-mono text-xs text-(--ui-text-muted)">
              {{ riskCategoryLabel[r.category] }} ・ {{ r.location }}
            </p>
            <p class="text-(--ui-text) line-clamp-2">{{ r.description }}</p>
          </div>
        </li>
      </ul>
    </UCard>

    <!-- 何もリスクがない場合の肯定的説明 -->
    <UCard v-else>
      <div class="flex items-center gap-2 text-sm">
        <UIcon name="i-lucide-shield-check" class="size-4 text-emerald-500 shrink-0" />
        <span class="text-(--ui-text)">
          静的解析で検出された未解析リスクはありません。
        </span>
        <span class="text-xs text-(--ui-text-muted)">
          (INDIRECT / OFFSET / 動的 VBA 参照などが見つからなかった)
        </span>
      </div>
    </UCard>
  </div>
</template>
