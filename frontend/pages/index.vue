<script setup lang="ts">
/**
 * ホーム — ジョブ一覧 + 新規アップロード.
 *
 * - 上部: ヒーロー + 現在ジョブ (あれば) + アクションボタン
 * - 中央: アップロードゾーン or 分析進捗
 * - 下部: 過去のジョブ一覧
 */

import type { JobMeta } from '~/types/api'
import type { AnalyzePhase } from '~/components/AnalyzeProgress.vue'

definePageMeta({ layout: 'default' })
useHead({ title: 'Excel 改修支援ツール' })

const backend = useBackend()
const jobStore = useJobStore()
const toast = useToast()

await jobStore.refreshJobs()

const phase = ref<AnalyzePhase>('idle')
const analyzingFilename = ref('')
const errorMessage = ref('')

async function onSelect(file: File) {
  phase.value = 'extracting'
  analyzingFilename.value = file.name
  errorMessage.value = ''
  try {
    const jobId = await backend.extract(file)
    phase.value = 'analyzing'
    await backend.analyze(jobId)
    phase.value = 'done'
    jobStore.setCurrentJobId(jobId)
    await jobStore.refreshJobs()
    toast.add({
      title: '分析完了',
      description: `${file.name} の設計書が生成されました`,
      color: 'success',
      icon: 'i-lucide-check-circle-2',
    })
    // 自動で設計書ページへ遷移
    setTimeout(() => navigateTo(`/spec/${jobId}`), 800)
  } catch (e) {
    phase.value = 'error'
    errorMessage.value = e instanceof Error ? e.message : String(e)
    toast.add({
      title: '分析失敗',
      description: errorMessage.value,
      color: 'error',
      icon: 'i-lucide-x-circle',
    })
  }
}

function onOpen(jobId: string) {
  jobStore.setCurrentJobId(jobId)
  navigateTo(`/spec/${jobId}`)
}

// 削除フロー
const dialogOpen = ref(false)
const dialogJob = ref<JobMeta | null>(null)
function askDelete(job: JobMeta) {
  dialogJob.value = job
  dialogOpen.value = true
}
async function confirmDelete(job: JobMeta) {
  try {
    await backend.deleteJob(job.job_id)
    if (jobStore.currentJobId === job.job_id) jobStore.setCurrentJobId(null)
    await jobStore.refreshJobs()
    toast.add({ title: '削除しました', color: 'neutral', icon: 'i-lucide-trash-2' })
  } catch (e) {
    toast.add({
      title: '削除失敗',
      description: e instanceof Error ? e.message : String(e),
      color: 'error',
    })
  }
}

// 現在ジョブのカード上アクション
function openCurrentSpec() {
  if (jobStore.currentJobId) navigateTo(`/spec/${jobStore.currentJobId}`)
}
function openCurrentChat() {
  if (jobStore.currentJobId) navigateTo(`/chat/${jobStore.currentJobId}`)
}

// メトリクス (ジョブ一覧サマリ)
const stats = computed(() => {
  const total = jobStore.jobs.length
  const analyzed = jobStore.jobs.filter(j => j.status === 'analyzed').length
  const failed = jobStore.jobs.filter(j => j.status === 'failed').length
  return { total, analyzed, failed }
})
</script>

<template>
  <div class="space-y-8">
    <!-- Hero -->
    <header class="flex items-start justify-between gap-6 flex-wrap">
      <div>
        <h1 class="text-3xl font-bold tracking-tight text-(--ui-text-highlighted)">
          Excel 改修支援ツール
        </h1>
        <p class="mt-2 text-(--ui-text-muted) max-w-2xl">
          VBA / 数式 / 参照関係を含む <code class="text-xs px-1 py-0.5 rounded bg-(--ui-bg-elevated)">.xlsm</code> /
          <code class="text-xs px-1 py-0.5 rounded bg-(--ui-bg-elevated)">.xls</code>
          ツールの統合設計書を生成し、LLM と対話しながら安全に改修するためのワークベンチ。
        </p>
      </div>
      <div class="flex gap-3">
        <div class="flex flex-col items-center px-4 py-2 rounded-xl bg-(--ui-bg-elevated)">
          <span class="text-2xl font-bold text-(--ui-text-highlighted) tabular-nums">{{ stats.total }}</span>
          <span class="text-[10px] uppercase tracking-wide text-(--ui-text-muted)">総ジョブ</span>
        </div>
        <div class="flex flex-col items-center px-4 py-2 rounded-xl bg-emerald-50 dark:bg-emerald-950">
          <span class="text-2xl font-bold text-emerald-700 dark:text-emerald-300 tabular-nums">{{ stats.analyzed }}</span>
          <span class="text-[10px] uppercase tracking-wide text-emerald-700/70 dark:text-emerald-300/70">分析済</span>
        </div>
        <div v-if="stats.failed > 0" class="flex flex-col items-center px-4 py-2 rounded-xl bg-red-50 dark:bg-red-950">
          <span class="text-2xl font-bold text-red-700 dark:text-red-300 tabular-nums">{{ stats.failed }}</span>
          <span class="text-[10px] uppercase tracking-wide text-red-700/70 dark:text-red-300/70">失敗</span>
        </div>
      </div>
    </header>

    <!-- 現在のジョブ -->
    <section v-if="jobStore.currentJob" class="space-y-3">
      <h2 class="text-sm font-semibold uppercase tracking-wide text-(--ui-text-muted)">現在のジョブ</h2>
      <UCard>
        <div class="flex items-center justify-between flex-wrap gap-3">
          <div class="flex items-center gap-3">
            <div class="size-11 rounded-xl bg-gradient-to-br from-indigo-500 to-purple-600 text-white flex items-center justify-center">
              <UIcon name="i-lucide-file-spreadsheet" class="size-5" />
            </div>
            <div>
              <p class="font-semibold text-(--ui-text-highlighted)">{{ jobStore.currentJob.filename }}</p>
              <div class="flex items-center gap-2 mt-1">
                <JobStatusBadge :status="jobStore.currentJob.status" size="sm" />
                <span class="text-xs text-(--ui-text-muted)">{{ jobStore.currentJob.created_at.slice(0, 16).replace('T', ' ') }}</span>
              </div>
            </div>
          </div>
          <div class="flex gap-2">
            <UButton icon="i-lucide-file-text" color="primary" variant="solid" @click="openCurrentSpec">
              設計書を見る
            </UButton>
            <UButton icon="i-lucide-message-circle" color="primary" variant="soft" @click="openCurrentChat">
              チャット
            </UButton>
          </div>
        </div>
      </UCard>
    </section>

    <!-- アップロード or 進捗 -->
    <section class="space-y-3">
      <h2 class="text-sm font-semibold uppercase tracking-wide text-(--ui-text-muted)">
        {{ jobStore.currentJob ? '別の Excel を分析' : 'Excel を分析' }}
      </h2>
      <AnalyzeProgress
        v-if="phase !== 'idle'"
        :phase="phase"
        :filename="analyzingFilename"
        :error-message="errorMessage"
      />
      <FileDropzone v-else @select="onSelect" />
      <div v-if="phase === 'error'" class="flex justify-end">
        <UButton variant="ghost" color="neutral" icon="i-lucide-rotate-ccw" @click="phase = 'idle'">
          もう一度
        </UButton>
      </div>
    </section>

    <!-- ジョブ一覧 -->
    <section class="space-y-3">
      <div class="flex items-center justify-between">
        <h2 class="text-sm font-semibold uppercase tracking-wide text-(--ui-text-muted)">過去のジョブ</h2>
        <UButton
          icon="i-lucide-refresh-cw"
          color="neutral"
          variant="ghost"
          size="xs"
          :loading="jobStore.loading"
          @click="jobStore.refreshJobs()"
        >
          再読込
        </UButton>
      </div>

      <UAlert
        v-if="jobStore.error"
        color="error"
        variant="subtle"
        icon="i-lucide-alert-triangle"
        title="ジョブ一覧の取得に失敗"
        :description="jobStore.error"
      />

      <div v-if="jobStore.jobs.length === 0 && !jobStore.loading" class="rounded-2xl border border-dashed border-(--ui-border) p-10 text-center">
        <div class="size-12 mx-auto rounded-2xl bg-(--ui-bg-muted) flex items-center justify-center text-(--ui-text-muted) mb-3">
          <UIcon name="i-lucide-inbox" class="size-6" />
        </div>
        <p class="text-sm text-(--ui-text-muted)">まだジョブがありません。上のフォームから Excel をアップロードしてください。</p>
      </div>

      <div v-else class="grid grid-cols-1 lg:grid-cols-2 gap-3">
        <JobCard
          v-for="job in jobStore.jobs"
          :key="job.job_id"
          :job="job"
          :is-current="jobStore.currentJobId === job.job_id"
          @open="onOpen"
          @delete="askDelete"
        />
      </div>
    </section>

    <DeleteJobDialog v-model:open="dialogOpen" :job="dialogJob" @confirm="confirmDelete" />
  </div>
</template>
