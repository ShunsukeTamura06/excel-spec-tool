<script setup lang="ts">
/**
 * WorkbookDiffData の描画 (波及範囲 / 既存リスク / カテゴリ別構造差分).
 *
 * `pages/diff.vue` (手動2ジョブ比較) と、チャットの名前定義修正カード
 * (S2 増分1の自己検証結果表示) の両方から使う共有コンポーネント。
 */

import type { ChangeType, WorkbookDiffData } from '~/types/api'

const props = defineProps<{
  diff: WorkbookDiffData
}>()

const changeTypeStyle: Record<ChangeType, { color: 'success' | 'error' | 'warning'; label: string }> = {
  added: { color: 'success', label: '追加' },
  removed: { color: 'error', label: '削除' },
  modified: { color: 'warning', label: '変更' },
}

interface Section {
  key: string
  title: string
  icon: string
  count: number
}

const sections = computed<Section[]>(() => [
  { key: 'cells', title: 'セル', icon: 'i-lucide-grid-3x3', count: props.diff.cells.length },
  { key: 'named_ranges', title: '名前付き範囲', icon: 'i-lucide-tag', count: props.diff.named_ranges.length },
  { key: 'conditional_formats', title: '条件付き書式', icon: 'i-lucide-palette', count: props.diff.conditional_formats.length },
  { key: 'data_validations', title: '入力規則', icon: 'i-lucide-list-checks', count: props.diff.data_validations.length },
  { key: 'charts', title: 'グラフ', icon: 'i-lucide-bar-chart-3', count: props.diff.charts.length },
  { key: 'pivot_tables', title: 'ピボットテーブル', icon: 'i-lucide-table', count: props.diff.pivot_tables.length },
  { key: 'vba_modules', title: 'VBAモジュール', icon: 'i-lucide-code-2', count: props.diff.vba_modules.length },
].filter(s => s.count > 0))

const isEmpty = computed(
  () => sections.value.length === 0 && props.diff.blast_radius.length === 0,
)
</script>

<template>
  <div class="space-y-5">
    <UAlert
      v-if="isEmpty"
      color="success"
      variant="subtle"
      icon="i-lucide-check-circle-2"
      title="差分なし"
      description="構造的な変更は検出されませんでした。"
    />

    <!-- 波及範囲 (先に見せる: 一番注意が必要な情報) -->
    <UCard v-if="diff.blast_radius.length > 0" class="border-amber-300 dark:border-amber-800">
      <template #header>
        <div class="flex items-center gap-2">
          <UIcon name="i-lucide-radar" class="size-5 text-amber-600" />
          <h2 class="font-semibold text-(--ui-text-highlighted)">波及範囲</h2>
          <UBadge color="warning" variant="subtle">{{ diff.blast_radius.length }}</UBadge>
        </div>
      </template>
      <div class="space-y-3">
        <div v-for="entry in diff.blast_radius" :key="entry.location" class="text-sm">
          <div class="flex items-center gap-2">
            <UBadge :color="changeTypeStyle[entry.change_type].color" variant="subtle" size="sm">
              {{ changeTypeStyle[entry.change_type].label }}
            </UBadge>
            <span class="font-mono text-(--ui-text-highlighted)">{{ entry.location }}</span>
            <span class="text-(--ui-text-muted)">を参照している箇所:</span>
          </div>
          <ul class="mt-1 ml-6 space-y-0.5">
            <li v-for="(ref, i) in entry.referenced_by" :key="i" class="text-xs text-(--ui-text-muted) font-mono">
              {{ ref.from }} ({{ ref.kind }})
            </li>
          </ul>
        </div>
      </div>
    </UCard>

    <!-- 既存リスク -->
    <UCard v-if="diff.existing_risks.length > 0">
      <template #header>
        <div class="flex items-center gap-2">
          <UIcon name="i-lucide-shield-alert" class="size-5 text-red-600" />
          <h2 class="font-semibold text-(--ui-text-highlighted)">既知のリスク (前バージョン時点)</h2>
          <UBadge color="error" variant="subtle">{{ diff.existing_risks.length }}</UBadge>
        </div>
      </template>
      <div class="space-y-2">
        <div v-for="(risk, i) in diff.existing_risks" :key="i" class="text-sm border-b border-(--ui-border) pb-2 last:border-0 last:pb-0">
          <div class="flex items-center gap-2">
            <UBadge :color="risk.severity === 'high' ? 'error' : risk.severity === 'medium' ? 'warning' : 'neutral'" variant="subtle" size="sm">
              {{ risk.severity }}
            </UBadge>
            <span class="font-mono text-xs text-(--ui-text-muted)">{{ risk.location }}</span>
          </div>
          <p class="text-(--ui-text-highlighted) mt-1">{{ risk.description }}</p>
          <p class="text-xs text-(--ui-text-muted) mt-0.5">{{ risk.recommendation }}</p>
        </div>
      </div>
    </UCard>

    <!-- 構造差分 (カテゴリ別) -->
    <UCard v-for="section in sections" :key="section.key">
      <template #header>
        <div class="flex items-center gap-2">
          <UIcon :name="section.icon" class="size-5 text-(--ui-text-muted)" />
          <h2 class="font-semibold text-(--ui-text-highlighted)">{{ section.title }}</h2>
          <UBadge color="neutral" variant="subtle">{{ section.count }}</UBadge>
        </div>
      </template>

      <!-- セル -->
      <div v-if="section.key === 'cells'" class="space-y-1.5">
        <div v-for="c in diff.cells" :key="`${c.sheet}!${c.coord}`" class="flex items-start gap-2 text-sm">
          <UBadge :color="changeTypeStyle[c.change_type].color" variant="subtle" size="sm">
            {{ changeTypeStyle[c.change_type].label }}
          </UBadge>
          <span class="font-mono text-(--ui-text-highlighted)">{{ c.sheet }}!{{ c.coord }}</span>
          <span class="text-(--ui-text-muted)">
            {{ c.before_formula ?? c.before_value ?? '(空)' }} → {{ c.after_formula ?? c.after_value ?? '(空)' }}
            <template v-if="c.before_number_format !== c.after_number_format">
              (表示形式: {{ c.before_number_format ?? '既定' }} → {{ c.after_number_format ?? '既定' }})
            </template>
          </span>
        </div>
      </div>

      <!-- 名前付き範囲 -->
      <div v-else-if="section.key === 'named_ranges'" class="space-y-1.5">
        <div v-for="nr in diff.named_ranges" :key="nr.name" class="flex items-start gap-2 text-sm">
          <UBadge :color="changeTypeStyle[nr.change_type].color" variant="subtle" size="sm">
            {{ changeTypeStyle[nr.change_type].label }}
          </UBadge>
          <span class="font-mono text-(--ui-text-highlighted)">{{ nr.name }}</span>
          <span class="text-(--ui-text-muted)">{{ nr.before_refers_to ?? '(なし)' }} → {{ nr.after_refers_to ?? '(なし)' }}</span>
        </div>
      </div>

      <!-- 条件付き書式 -->
      <div v-else-if="section.key === 'conditional_formats'" class="space-y-1.5">
        <div v-for="(cf, i) in diff.conditional_formats" :key="i" class="flex items-start gap-2 text-sm">
          <UBadge :color="changeTypeStyle[cf.change_type].color" variant="subtle" size="sm">
            {{ changeTypeStyle[cf.change_type].label }}
          </UBadge>
          <span class="font-mono text-(--ui-text-highlighted)">{{ cf.sheet }}!{{ cf.range }}</span>
          <span class="text-(--ui-text-muted)">{{ cf.before_rule ?? '(なし)' }} → {{ cf.after_rule ?? '(なし)' }}</span>
        </div>
      </div>

      <!-- 入力規則 -->
      <div v-else-if="section.key === 'data_validations'" class="space-y-1.5">
        <div v-for="(dv, i) in diff.data_validations" :key="i" class="flex items-start gap-2 text-sm">
          <UBadge :color="changeTypeStyle[dv.change_type].color" variant="subtle" size="sm">
            {{ changeTypeStyle[dv.change_type].label }}
          </UBadge>
          <span class="font-mono text-(--ui-text-highlighted)">{{ dv.sheet }}!{{ dv.range }}</span>
          <span class="text-(--ui-text-muted)">{{ dv.before?.type ?? '(なし)' }} → {{ dv.after?.type ?? '(なし)' }}</span>
        </div>
      </div>

      <!-- グラフ -->
      <div v-else-if="section.key === 'charts'" class="space-y-1.5">
        <div v-for="(chart, i) in diff.charts" :key="i" class="flex items-start gap-2 text-sm">
          <UBadge :color="changeTypeStyle[chart.change_type].color" variant="subtle" size="sm">
            {{ changeTypeStyle[chart.change_type].label }}
          </UBadge>
          <span class="font-mono text-(--ui-text-highlighted)">{{ chart.sheet }}!{{ chart.key }}</span>
        </div>
      </div>

      <!-- ピボットテーブル -->
      <div v-else-if="section.key === 'pivot_tables'" class="space-y-1.5">
        <div v-for="(pt, i) in diff.pivot_tables" :key="i" class="flex items-start gap-2 text-sm">
          <UBadge :color="changeTypeStyle[pt.change_type].color" variant="subtle" size="sm">
            {{ changeTypeStyle[pt.change_type].label }}
          </UBadge>
          <span class="font-mono text-(--ui-text-highlighted)">{{ pt.sheet }}!{{ pt.name }}</span>
          <span class="text-(--ui-text-muted)">{{ pt.before?.source_ref ?? '' }} → {{ pt.after?.source_ref ?? '' }}</span>
        </div>
      </div>

      <!-- VBAモジュール -->
      <div v-else-if="section.key === 'vba_modules'" class="space-y-1.5">
        <div v-for="vm in diff.vba_modules" :key="vm.name" class="flex items-start gap-2 text-sm">
          <UBadge :color="changeTypeStyle[vm.change_type].color" variant="subtle" size="sm">
            {{ changeTypeStyle[vm.change_type].label }}
          </UBadge>
          <span class="font-mono text-(--ui-text-highlighted)">{{ vm.name }}</span>
        </div>
      </div>
    </UCard>
  </div>
</template>
