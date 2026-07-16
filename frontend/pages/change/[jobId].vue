<script setup lang="ts">
/** 改修依頼の整理から、限定変更の計画・承認・検証・取得までを扱う導線. */
import type {
  ChangeBrief,
  FormulaFixResponse,
  SafeChangePlanData,
} from '~/types/api'

definePageMeta({ layout: 'default' })
useHead({ title: 'このExcelを直したい — xlblueprint' })

const route = useRoute()
const backend = useBackend()
const jobStore = useJobStore()
const jobId = computed(() => String(route.params.jobId))

onMounted(() => {
  if (jobStore.currentJobId !== jobId.value) jobStore.setCurrentJobId(jobId.value)
  void jobStore.refreshJobs()
})

const { data: diagnosis, error, pending } = useAsyncData(
  () => `change-diagnosis-${jobId.value}`,
  () => backend.getDiagnosis(jobId.value),
  { lazy: true },
)
const { data: workbook, error: workbookError } = useAsyncData(
  () => `change-workbook-${jobId.value}`,
  () => backend.getWorkbook(jobId.value),
  { lazy: true },
)

const selectedFeature = ref(typeof route.query.feature === 'string' ? route.query.feature : '')
const requestedOutcome = ref('')
const brief = ref<ChangeBrief | null>(null)
const creating = ref(false)
const createError = ref('')
const selectedRange = ref('')
const newLastRow = ref<number | null>(null)
const safePlan = ref<SafeChangePlanData | null>(null)
const planCreating = ref(false)
const planError = ref('')
const approved = ref(false)
const execution = ref<FormulaFixResponse | null>(null)
const executing = ref(false)
const downloading = ref(false)

interface RangeCandidate {
  ref: string
  currentLastRow: number
  usageCount: number
  related: boolean
}

const selectedFeatureData = computed(() =>
  diagnosis.value?.features.find(feature => feature.id === selectedFeature.value),
)

const rangeCandidates = computed<RangeCandidate[]>(() => {
  if (!workbook.value) return []
  const counts = new Map<string, number>()
  for (const sheet of workbook.value.sheets) {
    for (const formula of sheet.formulas) {
      for (const ref of formula.refs) {
        if (!ref.includes('!') || !ref.includes(':')) continue
        const cellRange = ref.slice(ref.lastIndexOf('!') + 1)
        if (!/^\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$/i.test(cellRange)) continue
        counts.set(ref, (counts.get(ref) ?? 0) + 1)
      }
    }
  }
  const relatedSheets = new Set(selectedFeatureData.value?.related_sheets ?? [])
  return [...counts.entries()]
    .map(([ref, usageCount]) => {
      const match = ref.match(/:\$?[A-Z]{1,3}\$?(\d+)$/i)
      const sheetName = ref.slice(0, ref.lastIndexOf('!')).replace(/^'|'$/g, '')
      return {
        ref,
        currentLastRow: match ? Number(match[1]) : 0,
        usageCount,
        related: relatedSheets.has(sheetName),
      }
    })
    .filter(candidate => candidate.currentLastRow > 0)
    .sort((left, right) =>
      Number(right.related) - Number(left.related)
      || right.usageCount - left.usageCount
      || left.ref.localeCompare(right.ref),
    )
    .slice(0, 30)
})

const selectedCandidate = computed(() =>
  rangeCandidates.value.find(candidate => candidate.ref === selectedRange.value),
)

const proposedNewRef = computed(() => {
  const candidate = selectedCandidate.value
  const row = newLastRow.value
  if (!candidate || row === null || row <= candidate.currentLastRow) return ''
  return candidate.ref.replace(/(:\$?[A-Z]{1,3}\$?)\d+$/i, `$1${row}`)
})

watch(
  [diagnosis, () => route.query.feature],
  ([loadedDiagnosis, feature]) => {
    if (
      loadedDiagnosis
      && typeof feature === 'string'
      && loadedDiagnosis.features.some(item => item.id === feature)
    ) {
      selectedFeature.value = feature
    }
  },
  { immediate: true },
)

async function createBrief() {
  if (!requestedOutcome.value.trim()) return
  creating.value = true
  createError.value = ''
  brief.value = null
  safePlan.value = null
  execution.value = null
  approved.value = false
  try {
    const createdBrief = await backend.createChangeBrief(
      jobId.value,
      requestedOutcome.value,
      selectedFeature.value || undefined,
    )
    brief.value = createdBrief
    if (createdBrief.automation !== 'supported' || rangeCandidates.value.length === 0) {
      await openConsultation(createdBrief)
    }
  } catch (cause) {
    createError.value = friendlyMessage(cause)
  } finally {
    creating.value = false
  }
}

watch(rangeCandidates, (candidates) => {
  const first = candidates[0]
  if (!first || selectedRange.value) return
  selectedRange.value = first.ref
  newLastRow.value = first.currentLastRow + 10
}, { immediate: true })

watch(selectedRange, () => {
  const candidate = selectedCandidate.value
  if (candidate) newLastRow.value = candidate.currentLastRow + 10
  safePlan.value = null
  execution.value = null
  approved.value = false
})

watch(newLastRow, () => {
  safePlan.value = null
  execution.value = null
  approved.value = false
})

async function createSafePlan() {
  if (!selectedCandidate.value || !proposedNewRef.value) return
  planCreating.value = true
  planError.value = ''
  execution.value = null
  approved.value = false
  try {
    safePlan.value = await backend.createSafeChangePlan(
      jobId.value,
      selectedCandidate.value.ref,
      proposedNewRef.value,
    )
  } catch (cause) {
    planError.value = friendlyMessage(cause)
  } finally {
    planCreating.value = false
  }
}

async function executePlan() {
  if (!safePlan.value || !approved.value) return
  executing.value = true
  planError.value = ''
  try {
    execution.value = await backend.executeSafeChangePlan(jobId.value, safePlan.value.plan)
    await jobStore.refreshJobs()
  } catch (cause) {
    planError.value = friendlyMessage(cause)
  } finally {
    executing.value = false
  }
}

async function downloadResult() {
  if (!execution.value) return
  downloading.value = true
  planError.value = ''
  try {
    const blob = await backend.downloadWorkbook(execution.value.new_job_id)
    const sourceName = diagnosis.value?.filename ?? 'workbook.xlsx'
    const dot = sourceName.lastIndexOf('.')
    const stem = dot > 0 ? sourceName.slice(0, dot) : sourceName
    const suffix = dot > 0 ? sourceName.slice(dot) : '.xlsx'
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = `${stem}_xlblueprint${suffix}`
    anchor.click()
    URL.revokeObjectURL(url)
  } catch (cause) {
    planError.value = friendlyMessage(cause)
  } finally {
    downloading.value = false
  }
}

async function openConsultation(changeBrief: ChangeBrief | null = brief.value) {
  if (!changeBrief) return
  const value = changeBrief
  const featureName = diagnosis.value?.features.find(item => item.id === value.feature_id)?.name
  const prompt = featureName
    ? `${featureName}について、${value.requested_outcome}`
    : value.requested_outcome
  const title = value.requested_outcome.length > 28
    ? `${value.requested_outcome.slice(0, 27)}…`
    : value.requested_outcome
  const session = await backend.createChatSession(jobId.value, title)
  await navigateTo({
    path: `/chat/${jobId.value}`,
    query: { session: session.session_id, request: prompt, autosend: '1' },
  })
}

function rewriteRequest() {
  brief.value = null
  safePlan.value = null
  execution.value = null
  approved.value = false
}
</script>

<template>
  <div class="space-y-5 max-w-5xl mx-auto">
    <div class="flex items-center gap-2 text-sm text-(--ui-text-muted)">
      <NuxtLink to="/" class="hover:text-(--ui-primary)">ホーム</NuxtLink>
      <UIcon name="i-lucide-chevron-right" class="size-4" />
      <NuxtLink :to="`/spec/${jobId}`" class="hover:text-(--ui-primary)">Excel診断</NuxtLink>
      <UIcon name="i-lucide-chevron-right" class="size-4" />
      <span class="text-(--ui-text-highlighted)">直したい</span>
    </div>

    <div>
      <h1 class="text-2xl font-bold text-(--ui-text-highlighted)">何を変えたいですか？</h1>
      <p class="mt-2 text-sm text-(--ui-text-muted)">
        普段の言葉で大丈夫です。必要な場所はこちらで調べます。
      </p>
    </div>

    <UAlert
      v-if="error || workbookError || createError || planError"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="読み込みに失敗しました"
      :description="planError || createError || friendlyMessage(error || workbookError)"
    />

    <UCard v-if="diagnosis">
      <div class="space-y-4">
        <div>
          <UTextarea
            id="request"
            v-model="requestedOutcome"
            :rows="3"
            autoresize
            :maxrows="8"
            class="w-full"
            placeholder="例: Outputシートの結果を見やすくしたい"
          />
        </div>
        <details v-if="diagnosis.features.length" class="text-sm">
          <summary class="cursor-pointer text-(--ui-text-muted)">対象の機能を指定する（任意）</summary>
          <select
            id="feature"
            v-model="selectedFeature"
            class="w-full mt-3 rounded-lg border border-(--ui-border) bg-(--ui-bg) px-3 py-2 text-sm"
          >
            <option value="">ファイル全体／分からない</option>
            <option v-for="feature in diagnosis.features" :key="feature.id" :value="feature.id">
              {{ feature.name }} — {{ feature.summary }}
            </option>
          </select>
        </details>
        <UButton
          color="primary"
          icon="i-lucide-sparkles"
          :loading="creating"
          :disabled="!requestedOutcome.trim()"
          @click="createBrief"
        >
          改修案を見る
        </UButton>
      </div>
    </UCard>
    <SpecSkeleton v-else-if="pending" />

    <div v-if="brief" class="space-y-4">
      <UCard>
        <div class="space-y-4">
          <div class="flex items-start gap-3">
            <div class="rounded-full bg-emerald-100 dark:bg-emerald-950 p-2 shrink-0">
              <UIcon name="i-lucide-check" class="size-5 text-emerald-700 dark:text-emerald-300" />
            </div>
            <div>
              <p class="text-xs text-(--ui-text-muted)">実現したいこと</p>
              <p class="font-semibold mt-1">{{ brief.requested_outcome }}</p>
            </div>
          </div>

          <details class="rounded-lg border border-(--ui-border) px-4 py-3">
            <summary class="cursor-pointer text-sm text-(--ui-text-muted)">調査対象と確認基準を見る</summary>
            <div class="mt-3 space-y-3 text-sm">
              <p>{{ brief.current_behavior }}</p>
              <div v-if="brief.affected_areas.length" class="flex flex-wrap gap-2">
                <UBadge v-for="area in brief.affected_areas" :key="area" color="neutral" variant="subtle">
                  {{ area }}
                </UBadge>
              </div>
              <ul class="list-disc pl-5 text-(--ui-text-muted)">
                <li v-for="item in brief.acceptance_criteria" :key="item">{{ item }}</li>
              </ul>
            </div>
          </details>

          <div class="flex gap-2 flex-wrap">
            <UButton
              v-if="brief.automation !== 'supported'"
              color="primary"
              icon="i-lucide-sparkles"
              @click="openConsultation()"
            >
              改修案を作る
            </UButton>
            <UButton color="neutral" variant="ghost" icon="i-lucide-pencil" @click="rewriteRequest">
              書き直す
            </UButton>
          </div>
        </div>
      </UCard>

      <UCard v-if="brief.automation === 'supported' && rangeCandidates.length">
        <template #header>
          <div class="flex items-center justify-between gap-3 flex-wrap">
            <div>
              <h2 class="text-lg font-semibold">安全に自動変更できる内容</h2>
              <p class="text-sm text-(--ui-text-muted) mt-1">
                現在は、数式が参照するデータ範囲を下方向へ広げる変更に対応しています。
              </p>
            </div>
            <UBadge color="success" variant="subtle">原本を保持</UBadge>
          </div>
        </template>

        <div v-if="rangeCandidates.length" class="space-y-5">
          <div class="grid grid-cols-1 md:grid-cols-[minmax(0,1fr)_180px] gap-4">
            <div>
              <label for="target-range" class="block text-sm font-semibold mb-2">広げるデータ範囲</label>
              <select
                id="target-range"
                v-model="selectedRange"
                class="w-full rounded-lg border border-(--ui-border) bg-(--ui-bg) px-3 py-2 text-sm font-mono"
              >
                <option v-for="candidate in rangeCandidates" :key="candidate.ref" :value="candidate.ref">
                  {{ candidate.ref }} — {{ candidate.usageCount }}個の数式で使用
                </option>
              </select>
              <p class="text-xs text-(--ui-text-muted) mt-2">
                Excel内の数式から実在する参照範囲だけを候補にしています。
              </p>
            </div>
            <div>
              <label for="last-row" class="block text-sm font-semibold mb-2">新しい最終行</label>
              <input
                id="last-row"
                v-model.number="newLastRow"
                type="number"
                :min="(selectedCandidate?.currentLastRow ?? 0) + 1"
                step="1"
                class="w-full rounded-lg border border-(--ui-border) bg-(--ui-bg) px-3 py-2 text-sm"
              >
              <p v-if="selectedCandidate" class="text-xs text-(--ui-text-muted) mt-2">
                現在は {{ selectedCandidate.currentLastRow }} 行目まで
              </p>
            </div>
          </div>

          <UAlert
            color="neutral"
            variant="subtle"
            icon="i-lucide-arrow-right"
            title="予定する参照範囲"
            :description="proposedNewRef ? `${selectedRange} → ${proposedNewRef}` : '現在より大きい最終行を入力してください。'"
          />

          <UButton
            color="primary"
            icon="i-lucide-list-checks"
            :loading="planCreating"
            :disabled="!proposedNewRef"
            @click="createSafePlan"
          >
            変更計画を確認する
          </UButton>
        </div>

        <div v-else class="space-y-3">
          <UAlert
            color="neutral"
            variant="subtle"
            icon="i-lucide-circle-slash"
            title="自動変更できる範囲が見つかりませんでした"
            description="数式内に、下方向へ安全に広げられるシート付きセル範囲がありません。"
          />
          <UButton color="primary" variant="soft" icon="i-lucide-messages-square" @click="openConsultation()">
            根拠を添えて相談する
          </UButton>
        </div>
      </UCard>

      <UCard v-if="safePlan">
        <template #header>
          <div class="flex items-center justify-between gap-3 flex-wrap">
            <div>
              <h2 class="text-lg font-semibold">{{ safePlan.title }}</h2>
              <p class="text-sm text-(--ui-text-muted) mt-1">{{ safePlan.summary }}</p>
            </div>
            <UBadge :color="safePlan.automation === 'supported' ? 'success' : 'warning'" variant="subtle">
              {{ safePlan.automation === 'supported' ? '自動適用可能' : '追加確認が必要' }}
            </UBadge>
          </div>
        </template>

        <div class="space-y-5">
          <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
            <div>
              <h3 class="text-sm font-semibold mb-2">適用前の確認</h3>
              <ul class="list-disc pl-5 space-y-2 text-sm">
                <li v-for="item in safePlan.preconditions" :key="item">{{ item }}</li>
              </ul>
            </div>
            <div>
              <h3 class="text-sm font-semibold mb-2">検証の合格条件</h3>
              <ul class="space-y-2 text-sm">
                <li v-for="item in safePlan.acceptance_criteria" :key="item" class="flex gap-2">
                  <UIcon name="i-lucide-square-check-big" class="size-4 mt-0.5 text-emerald-600 shrink-0" />
                  <span>{{ item }}</span>
                </li>
              </ul>
            </div>
          </div>

          <UAlert
            v-if="safePlan.warnings.length"
            color="warning"
            variant="subtle"
            icon="i-lucide-triangle-alert"
            title="追加で確認してください"
            :description="safePlan.warnings.join(' ')"
          />
          <UAlert
            color="info"
            variant="subtle"
            icon="i-lucide-monitor-check"
            title="この環境で検証できる範囲"
            :description="safePlan.verification_scope"
          />

          <div>
            <h3 class="text-sm font-semibold mb-3">
              変更予定の差分（{{ safePlan.expected_change_count }}件）
            </h3>
            <div class="rounded-lg bg-(--ui-bg-elevated) p-4 text-sm">
              <p class="font-semibold mb-2">主な変更場所</p>
              <div class="flex flex-wrap gap-2">
                <UBadge
                  v-for="location in safePlan.affected_locations.slice(0, 8)"
                  :key="location"
                  color="neutral"
                  variant="subtle"
                >
                  {{ location }}
                </UBadge>
                <span v-if="safePlan.affected_locations.length > 8" class="text-(--ui-text-muted)">
                  ほか {{ safePlan.affected_locations.length - 8 }}件
                </span>
              </div>
            </div>
            <details
              class="mt-3 rounded-lg border border-(--ui-border) p-4"
              :open="safePlan.expected_change_count <= 10"
            >
              <summary class="cursor-pointer text-sm font-semibold">
                数式・波及範囲・既存リスクの完全な差分を確認する
              </summary>
              <div class="mt-4">
                <WorkbookDiffView :diff="safePlan.expected_diff" />
              </div>
            </details>
          </div>

          <label class="flex items-start gap-3 rounded-lg border border-(--ui-border) p-4 cursor-pointer">
            <input v-model="approved" type="checkbox" class="mt-1 size-4 accent-(--ui-primary)">
            <span class="text-sm">
              上記の変更対象・確認事項・検証範囲を確認しました。原本を残したまま、別ファイルとして修正版を作ります。
            </span>
          </label>

          <UButton
            color="primary"
            icon="i-lucide-shield-check"
            :loading="executing"
            :disabled="!approved || !safePlan.can_apply"
            @click="executePlan"
          >
            原本を残して修正版を作る
          </UButton>
        </div>
      </UCard>

      <UCard v-if="execution">
        <template #header>
          <h2 class="text-lg font-semibold">修正版の検証結果</h2>
        </template>
        <div class="space-y-5">
          <VerificationGateAlert :report="execution.verification" :new-job-id="execution.new_job_id" />
          <details class="rounded-lg border border-(--ui-border) p-4">
            <summary class="cursor-pointer text-sm font-semibold">検証済みの完全な実差分を確認する</summary>
            <div class="mt-4">
              <WorkbookDiffView :diff="execution.diff" />
            </div>
          </details>
          <div class="flex gap-2 flex-wrap">
            <UButton
              color="primary"
              icon="i-lucide-download"
              :loading="downloading"
              @click="downloadResult"
            >
              検証済みの修正版をダウンロード
            </UButton>
            <UButton
              color="neutral"
              variant="soft"
              icon="i-lucide-git-compare"
              :to="{
                path: '/diff',
                query: {
                  before_job_id: jobId,
                  after_job_id: execution.new_job_id,
                },
              }"
            >
              差分を別画面で確認
            </UButton>
          </div>
        </div>
      </UCard>
    </div>
  </div>
</template>
