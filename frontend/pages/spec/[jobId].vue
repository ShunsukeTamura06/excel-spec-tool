<script setup lang="ts">
/**
 * 設計書ページ.
 * - ヘッダー (パンくず + ファイル名 + 状態 + ダウンロード + チャットへ)
 * - メトリクスダッシュボード (SpecMetrics)
 * - タブ (概要 / シート / VBA / 参照検索 / ダイアグラム[次コミット])
 */

definePageMeta({ layout: 'default' })
useHead({ title: '設計書 — Excel 改修支援ツール' })

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

// 並行取得: spec.md と workbook 構造
const { data: spec, error: specError, pending: specPending } = await useAsyncData(
  () => `spec-${jobId.value}`,
  () => backend.getSpec(jobId.value),
)
const { data: workbook, error: wbError, pending: wbPending } = await useAsyncData(
  () => `workbook-${jobId.value}`,
  () => backend.getWorkbook(jobId.value),
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
  { value: 'overview',   label: '概要',       icon: 'i-lucide-file-text' },
  { value: 'sheets',     label: 'シート',     icon: 'i-lucide-layout-grid' },
  { value: 'vba',        label: 'VBA',        icon: 'i-lucide-code-2' },
  { value: 'references', label: '参照検索',   icon: 'i-lucide-search' },
  { value: 'diagrams',   label: 'ダイアグラム', icon: 'i-lucide-network' },
])

const errorMsg = computed(() => {
  if (specError.value) return `設計書の取得に失敗: ${specError.value.message}`
  if (wbError.value) return `Workbook 構造の取得に失敗: ${wbError.value.message}`
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
      <span class="text-(--ui-text-highlighted)">設計書</span>
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
    <div v-if="spec && workbook" class="flex items-start justify-between gap-4 flex-wrap">
      <div>
        <h1 class="text-2xl font-bold tracking-tight text-(--ui-text-highlighted) flex items-center gap-2">
          <UIcon name="i-lucide-file-spreadsheet" class="size-6 text-emerald-600" />
          {{ spec.meta.filename }}
        </h1>
        <div class="flex items-center gap-2 mt-1.5">
          <JobStatusBadge :status="spec.meta.status" size="sm" />
          <span class="text-xs text-(--ui-text-muted)">
            作成: {{ spec.meta.created_at.slice(0, 16).replace('T', ' ') }}
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
          設計書をダウンロード
        </UButton>
        <UButton
          icon="i-lucide-message-circle"
          color="primary"
          :to="`/chat/${jobId}`"
        >
          チャットで改修相談
        </UButton>
      </div>
    </div>

    <!-- メトリクス -->
    <SpecMetrics v-if="workbook" :workbook="workbook" />

    <!-- ローディング (初回) -->
    <UCard v-if="(specPending || wbPending) && !spec && !workbook">
      <div class="py-10 flex items-center justify-center gap-2 text-(--ui-text-muted)">
        <UIcon name="i-lucide-loader-2" class="animate-spin size-5" />
        <span class="text-sm">読み込み中...</span>
      </div>
    </UCard>

    <!-- タブ -->
    <UTabs
      v-if="spec && workbook"
      v-model="tab"
      :items="tabItems"
      class="w-full"
      :ui="{ trigger: 'flex-1 sm:flex-none' }"
    >
      <template #content="{ item }">
        <div class="mt-4">
          <SpecOverview v-if="item.value === 'overview'" :markdown="spec.spec_md" />

          <SheetExplorer
            v-else-if="item.value === 'sheets'"
            :sheets="workbook.sheets"
            @search-reference="jumpToReference"
          />

          <VbaExplorer v-else-if="item.value === 'vba'" :modules="workbook.vba_modules" />

          <ReferenceSearch
            v-else-if="item.value === 'references'"
            ref="refSearchRef"
            :job-id="jobId"
            :initial-target="refTarget"
          />

          <UCard v-else-if="item.value === 'diagrams'">
            <UAlert
              icon="i-lucide-construction"
              color="warning"
              variant="subtle"
              title="ダイアグラム描画は次のコミットで実装します"
              description="シート依存グラフと VBA コールグラフを Vue Flow でインタラクティブ表示する予定."
            />
          </UCard>
        </div>
      </template>
    </UTabs>
  </div>
</template>
