<script setup lang="ts">
/** 空セルへの固定テキスト追加を確認し、OfficeCLIで修正版を生成・取得するカード. */

import type {
  CellTextBatchOperationData,
  FormulaFixResponse,
  SafeChangePlanData,
  ToolTraceItem,
} from '~/types/api'

const props = defineProps<{
  item: ToolTraceItem
  jobId?: string
}>()

interface ProposalResult {
  safe_plan?: SafeChangePlanData
  note?: string
}

const backend = useBackend()
const jobStore = useJobStore()
const result = computed(() => props.item.result as ProposalResult | undefined)
const safePlan = computed(() => result.value?.safe_plan)
const operation = computed<CellTextBatchOperationData | null>(() => {
  const candidate = safePlan.value?.plan.operation
  return candidate?.kind === 'cell_text_batch' ? candidate : null
})
const edits = computed(() => operation.value?.edits ?? [])

const pending = ref(false)
const downloading = ref(false)
const errorMsg = ref<string | null>(null)
const execution = ref<FormulaFixResponse | null>(null)

function downloadName(): string {
  const sourceName = jobStore.jobs.find(job => job.job_id === props.jobId)?.filename
    ?? jobStore.currentJob?.filename
    ?? 'workbook.xlsx'
  const dot = sourceName.lastIndexOf('.')
  const stem = dot > 0 ? sourceName.slice(0, dot) : sourceName
  const suffix = dot > 0 ? sourceName.slice(dot) : '.xlsx'
  return `${stem}_xlblueprint${suffix}`
}

async function createRevisedWorkbook() {
  if (!props.jobId || !safePlan.value) return
  pending.value = true
  errorMsg.value = null
  try {
    execution.value = await backend.executeSafeChangePlan(props.jobId, safePlan.value.plan.plan_id)
    await jobStore.refreshJobs()
  } catch (cause) {
    errorMsg.value = friendlyMessage(cause)
  } finally {
    pending.value = false
  }
}

async function downloadRevisedWorkbook() {
  if (!execution.value) return
  downloading.value = true
  errorMsg.value = null
  try {
    const blob = await backend.downloadWorkbook(execution.value.new_job_id)
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    anchor.href = url
    anchor.download = downloadName()
    anchor.click()
    URL.revokeObjectURL(url)
  } catch (cause) {
    errorMsg.value = friendlyMessage(cause)
  } finally {
    downloading.value = false
  }
}
</script>

<template>
  <div v-if="safePlan && operation" class="space-y-3">
    <div v-if="!execution" class="space-y-3">
      <div>
        <p class="text-sm font-medium text-(--ui-text-highlighted)">
          {{ safePlan.summary }}
        </p>
        <div class="mt-2 space-y-1">
          <div
            v-for="edit in edits.slice(0, 12)"
            :key="`${edit.sheet}!${edit.coord}`"
            class="flex gap-2 text-xs"
          >
            <code class="shrink-0 text-(--ui-primary)">{{ edit.sheet }}!{{ edit.coord }}</code>
            <span class="text-(--ui-text) break-words">{{ edit.value }}</span>
          </div>
          <p v-if="edits.length > 12" class="text-xs text-(--ui-text-muted)">
            他 {{ edits.length - 12 }} 件
          </p>
        </div>
      </div>

      <p class="text-[11px] text-(--ui-text-muted)">
        原本は変更しません。修正版を作った後、予定外の構造差分がないか自動検証します。
      </p>

      <UButton
        color="primary"
        icon="i-lucide-file-check-2"
        size="sm"
        :loading="pending"
        :disabled="!props.jobId || !safePlan.can_apply"
        @click="createRevisedWorkbook"
      >
        修正版を作る
      </UButton>
    </div>

    <UAlert
      v-if="errorMsg"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="修正版を作れませんでした"
      :description="errorMsg"
    />

    <div v-if="execution" class="space-y-3">
      <VerificationGateAlert
        :report="execution.verification"
        :new-job-id="execution.new_job_id"
      />
      <div class="flex flex-wrap gap-2">
        <UButton
          color="primary"
          icon="i-lucide-download"
          size="sm"
          :loading="downloading"
          @click="downloadRevisedWorkbook"
        >
          修正版をダウンロード
        </UButton>
        <UButton
          color="neutral"
          variant="soft"
          icon="i-lucide-git-compare"
          size="sm"
          :to="{
            path: '/diff',
            query: {
              before_job_id: props.jobId,
              after_job_id: execution.new_job_id,
            },
          }"
        >
          差分を別画面で確認
        </UButton>
      </div>
      <details class="rounded-md border border-(--ui-border) p-2">
        <summary class="cursor-pointer text-xs text-(--ui-text-muted)">
          実際の差分を見る
        </summary>
        <div class="mt-2">
          <WorkbookDiffView :diff="execution.diff" />
        </div>
      </details>
      <p class="text-[10px] text-(--ui-text-muted)">
        編集: {{ execution.provider.provider }}
        <span v-if="execution.provider.provider_version">
          {{ execution.provider.provider_version }}
        </span>
      </p>
    </div>
  </div>
</template>
