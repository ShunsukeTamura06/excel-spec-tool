<script setup lang="ts">
/**
 * 差分比較ページ (P1 安全ゲートの静的パート).
 *
 * before/after の2ジョブを選び、GET /diff で構造差分 + 波及範囲 + 既存リスクを
 * 取得して表示する。Excel COM での再計算・マクロ実行によるテストは含まない
 * (別増分、docs/VISION.ja.md 参照)。
 */

import type { ChangeType, WorkbookDiffData } from '~/types/api'

definePageMeta({ layout: 'default' })
useHead({ title: '差分比較 — xlblueprint' })

const backend = useBackend()
const jobStore = useJobStore()

onMounted(() => {
  void jobStore.refreshJobs()
})

// /diff は extracted.json / references.json が要るため、抽出済み以降のジョブだけを選べるようにする
const jobOptions = computed(() =>
  jobStore.jobs
    .filter(j => j.status === 'extracted' || j.status === 'analyzed')
    .map(j => ({ label: `${j.filename} (${j.created_at.slice(0, 16).replace('T', ' ')})`, value: j.job_id })),
)

const beforeJobId = ref<string | undefined>(undefined)
const afterJobId = ref<string | undefined>(undefined)

const canCompare = computed(
  () => !!beforeJobId.value && !!afterJobId.value && beforeJobId.value !== afterJobId.value,
)

const pending = ref(false)
const errorMsg = ref<string | null>(null)
const diff = ref<WorkbookDiffData | null>(null)

async function runDiff() {
  if (!canCompare.value || !beforeJobId.value || !afterJobId.value) return
  pending.value = true
  errorMsg.value = null
  diff.value = null
  try {
    diff.value = await backend.getDiff(beforeJobId.value, afterJobId.value)
  } catch (e) {
    errorMsg.value = friendlyMessage(e)
  } finally {
    pending.value = false
  }
}

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

const sections = computed<Section[]>(() => {
  if (!diff.value) return []
  return [
    { key: 'cells', title: 'セル', icon: 'i-lucide-grid-3x3', count: diff.value.cells.length },
    { key: 'named_ranges', title: '名前付き範囲', icon: 'i-lucide-tag', count: diff.value.named_ranges.length },
    { key: 'conditional_formats', title: '条件付き書式', icon: 'i-lucide-palette', count: diff.value.conditional_formats.length },
    { key: 'data_validations', title: '入力規則', icon: 'i-lucide-list-checks', count: diff.value.data_validations.length },
    { key: 'charts', title: 'グラフ', icon: 'i-lucide-bar-chart-3', count: diff.value.charts.length },
    { key: 'pivot_tables', title: 'ピボットテーブル', icon: 'i-lucide-table', count: diff.value.pivot_tables.length },
    { key: 'vba_modules', title: 'VBAモジュール', icon: 'i-lucide-code-2', count: diff.value.vba_modules.length },
  ].filter(s => s.count > 0)
})

const isEmpty = computed(
  () => !!diff.value && sections.value.length === 0 && diff.value.blast_radius.length === 0,
)
</script>

<template>
  <div class="space-y-5">
    <!-- パンくず -->
    <div class="flex items-center gap-2 text-sm text-(--ui-text-muted)">
      <NuxtLink to="/" class="hover:text-(--ui-primary) flex items-center gap-1">
        <UIcon name="i-lucide-home" class="size-3.5" /> ホーム
      </NuxtLink>
      <UIcon name="i-lucide-chevron-right" class="size-4" />
      <span class="text-(--ui-text-highlighted)">差分比較</span>
    </div>

    <div>
      <h1 class="text-2xl font-bold tracking-tight text-(--ui-text-highlighted) flex items-center gap-2">
        <UIcon name="i-lucide-git-compare" class="size-6 text-emerald-600" />
        差分比較
      </h1>
      <p class="text-sm text-(--ui-text-muted) mt-1">
        2つのジョブを「前」「後」として選ぶと、構造差分・波及範囲・既知のリスクをまとめて表示します。
        Excel を実際に開いて再計算するテストはまだ含まれません。
      </p>
    </div>

    <!-- ジョブ選択 -->
    <UCard>
      <div class="grid grid-cols-1 md:grid-cols-[1fr_auto_1fr_auto] gap-3 items-end">
        <div class="space-y-1.5">
          <label class="text-xs font-semibold uppercase tracking-wide text-(--ui-text-muted)">前 (before)</label>
          <USelect
            v-model="beforeJobId"
            :items="jobOptions"
            placeholder="ジョブを選択"
            icon="i-lucide-file-spreadsheet"
          />
        </div>
        <UIcon name="i-lucide-arrow-right" class="size-5 text-(--ui-text-muted) mb-2.5 hidden md:block" />
        <div class="space-y-1.5">
          <label class="text-xs font-semibold uppercase tracking-wide text-(--ui-text-muted)">後 (after)</label>
          <USelect
            v-model="afterJobId"
            :items="jobOptions"
            placeholder="ジョブを選択"
            icon="i-lucide-file-spreadsheet"
          />
        </div>
        <UButton
          icon="i-lucide-git-compare"
          color="primary"
          :loading="pending"
          :disabled="!canCompare"
          @click="runDiff"
        >
          比較する
        </UButton>
      </div>

      <p v-if="beforeJobId && afterJobId && beforeJobId === afterJobId" class="text-xs text-amber-600 dark:text-amber-400 mt-2">
        同じジョブは比較できません。
      </p>
      <p v-if="jobOptions.length === 0" class="text-xs text-(--ui-text-muted) mt-2">
        比較できるジョブがありません。まずホームで Excel をアップロードしてください。
      </p>
    </UCard>

    <UAlert
      v-if="errorMsg"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="比較に失敗しました"
      :description="errorMsg"
    />

    <!-- 結果 -->
    <template v-if="diff">
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
    </template>
  </div>
</template>
