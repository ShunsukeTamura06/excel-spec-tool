<script setup lang="ts">
/**
 * 設計書ページ.
 * - ヘッダー (パンくず + ファイル名 + 状態 + ダウンロード + チャットへ)
 * - メトリクスダッシュボード (SpecMetrics)
 * - タブ (概要 / シート / VBA / 参照検索 / ダイアグラム[次コミット])
 */

import { formatJstDateTime } from '~/utils/dateTime'

definePageMeta({ layout: 'default' })
useHead({ title: 'Excel診断 — xlblueprint' })

const route = useRoute()
const backend = useBackend()
const jobStore = useJobStore()
const toast = useToast()

const jobId = computed(() => String(route.params.jobId))

// URL の jobId を store に反映 (ブックマーク経由の到達対応)
onMounted(() => {
  if (jobId.value && jobStore.currentJobId !== jobId.value) {
    jobStore.setCurrentJobId(jobId.value)
  }
  void jobStore.refreshJobs()
})

// spec.md と workbook 構造を並行取得.
// `lazy: true` でナビゲーションをブロックせず、ページは即座に表示される.
// データ取得中は data.value が null のままで、テンプレート側は
// SpecSkeleton でプレースホルダを出す.
const { data: spec, error: specError, pending: specPending } = useAsyncData(
  () => `spec-${jobId.value}`,
  () => backend.getSpec(jobId.value),
  { lazy: true },
)
const { data: workbook, error: wbError, pending: wbPending } = useAsyncData(
  () => `workbook-${jobId.value}`,
  () => backend.getWorkbook(jobId.value),
  { lazy: true },
)
const { data: diagnosis, error: diagnosisError, pending: diagnosisPending } = useAsyncData(
  () => `diagnosis-${jobId.value}`,
  () => backend.getDiagnosis(jobId.value),
  { lazy: true },
)

// 参照検索タブから外部ジャンプを受け取るための共有 state
const refTarget = ref('')
const tab = ref('overview')
const refSearchRef = ref<{ runSearch: (t: string) => Promise<void> } | null>(null)

function jumpToReference(target: string) {
  refTarget.value = target
  tab.value = 'references'
  // 次の描画タイミングでマウントされた ReferenceSearch に実行を依頼
  nextTick(() => {
    refSearchRef.value?.runSearch(target)
  })
}

// 設計書 Markdown のダウンロード
function downloadSpec() {
  if (!spec.value?.spec_md) return
  const meta = spec.value.meta
  const base = meta.filename.replace(/\.[^.]+$/, '') || 'spec'
  const blob = new Blob([spec.value.spec_md], { type: 'text/markdown;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `${base}_spec.md`
  document.body.appendChild(a)
  a.click()
  a.remove()
  URL.revokeObjectURL(url)
  toast.add({
    title: 'ダウンロードしました',
    description: `${base}_spec.md`,
    color: 'success',
    icon: 'i-lucide-download',
  })
}

const tabItems = computed(() => [
  { value: 'overview',   label: '診断',         icon: 'i-lucide-clipboard-check' },
  { value: 'sheets',     label: 'シート詳細',   icon: 'i-lucide-layout-grid' },
  { value: 'vba',        label: 'VBA詳細',      icon: 'i-lucide-code-2' },
  { value: 'external',   label: '外部関数',     icon: 'i-lucide-puzzle' },
  { value: 'references', label: '参照を調べる', icon: 'i-lucide-search' },
  { value: 'diagrams',   label: '構造図',       icon: 'i-lucide-network' },
])

const errorMsg = computed(() => {
  if (specError.value) return friendlyMessage(specError.value)
  if (wbError.value) return friendlyMessage(wbError.value)
  if (diagnosisError.value) return friendlyMessage(diagnosisError.value)
  return null
})
</script>

<template>
  <div class="space-y-5">
    <!-- パンくず -->
    <div class="flex items-center gap-2 text-sm text-(--ui-text-muted)">
      <NuxtLink to="/" class="hover:text-(--ui-primary) flex items-center gap-1">
        <UIcon name="i-lucide-home" class="size-3.5" /> ホーム
      </NuxtLink>
      <UIcon name="i-lucide-chevron-right" class="size-4" />
      <span class="text-(--ui-text-highlighted)">Excel診断</span>
    </div>

    <UAlert
      v-if="errorMsg"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="読み込みエラー"
      :description="errorMsg"
    />

    <!-- ヘッダー -->
    <div v-if="spec && workbook && diagnosis" class="flex items-start justify-between gap-4 flex-wrap">
      <div>
        <h1 class="text-2xl font-bold tracking-tight text-(--ui-text-highlighted) flex items-center gap-2">
          <UIcon name="i-lucide-file-spreadsheet" class="size-6 text-emerald-600" />
          {{ spec.meta.filename }}
        </h1>
        <div class="flex items-center gap-2 mt-1.5">
          <JobStatusBadge :status="spec.meta.status" size="sm" />
          <span class="text-xs text-(--ui-text-muted)">
            作成: {{ formatJstDateTime(spec.meta.created_at) }}
          </span>
          <span class="text-[10px] font-mono text-(--ui-text-muted) opacity-60">{{ spec.meta.job_id }}</span>
        </div>
      </div>
      <div class="flex gap-2">
        <UButton
          icon="i-lucide-download"
          color="neutral"
          variant="soft"
          @click="downloadSpec"
        >
          技術レポートを保存
        </UButton>
        <UButton
          icon="i-lucide-wrench"
          color="primary"
          :to="`/change/${jobId}`"
        >
          このExcelを直したい
        </UButton>
      </div>
    </div>

    <!-- 初回ロード中: スケルトン表示 -->
    <SpecSkeleton
      v-if="(specPending || wbPending || diagnosisPending) && (!spec || !workbook || !diagnosis)"
    />

    <!-- タブ -->
    <UTabs
      v-if="spec && workbook && diagnosis"
      v-model="tab"
      :items="tabItems"
      class="w-full"
      :ui="{ trigger: 'flex-1 sm:flex-none' }"
    >
      <template #content="{ item }">
        <div class="mt-4">
          <SpecOverview
            v-if="item.value === 'overview'"
            :markdown="spec.spec_md"
            :workbook="workbook"
            :diagnosis="diagnosis"
            :job-id="jobId"
          />

          <SheetExplorer
            v-else-if="item.value === 'sheets'"
            :sheets="workbook.sheets"
            @search-reference="jumpToReference"
          />

          <VbaExplorer v-else-if="item.value === 'vba'" :modules="workbook.vba_modules" />

          <ExternalFunctionsExplorer
            v-else-if="item.value === 'external'"
            :job-id="jobId"
          />

          <ReferenceSearch
            v-else-if="item.value === 'references'"
            ref="refSearchRef"
            :job-id="jobId"
            :initial-target="refTarget"
          />

          <DiagramsPanel v-else-if="item.value === 'diagrams'" :job-id="jobId" />
        </div>
      </template>
    </UTabs>
  </div>
</template>
