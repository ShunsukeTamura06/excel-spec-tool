<script setup lang="ts">
/** VBA変更ZIPの取得と、Windows適用後.xlsmの静的検証を扱うカード. */

import type {
  FormulaFixResponse,
  MutationPlanData,
  ToolTraceItem,
  VbaProcedureReplaceOperationData,
} from '~/types/api'

const props = defineProps<{
  item: ToolTraceItem
  jobId?: string
}>()

interface CompactSafePlan {
  plan: MutationPlanData
  title: string
  summary: string
  can_apply: boolean
  preconditions: string[]
  warnings: string[]
  verification_scope: string
}

interface ProposalResult {
  safe_plan?: CompactSafePlan
  note?: string
}

const backend = useBackend()
const jobStore = useJobStore()
const result = computed(() => props.item.result as ProposalResult | undefined)
const safePlan = computed(() => result.value?.safe_plan)
const operation = computed<VbaProcedureReplaceOperationData | null>(() => {
  const candidate = safePlan.value?.plan.operation
  return candidate?.kind === 'vba_procedure_replace' ? candidate : null
})

const packagePending = ref(false)
const verifyPending = ref(false)
const downloadPending = ref(false)
const selectedFile = ref<File | null>(null)
const errorMsg = ref<string | null>(null)
const verificationResult = ref<FormulaFixResponse | null>(null)

function sourceStem(): string {
  const filename = jobStore.jobs.find(job => job.job_id === props.jobId)?.filename
    ?? jobStore.currentJob?.filename
    ?? 'workbook.xlsm'
  const dot = filename.lastIndexOf('.')
  return dot > 0 ? filename.slice(0, dot) : filename
}

function saveBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  anchor.click()
  URL.revokeObjectURL(url)
}

async function downloadPackage() {
  if (!props.jobId || !safePlan.value) return
  packagePending.value = true
  errorMsg.value = null
  try {
    const blob = await backend.downloadVbaChangePackage(props.jobId, safePlan.value.plan.plan_id)
    saveBlob(blob, `${sourceStem()}_vba_change.zip`)
  } catch (cause) {
    errorMsg.value = friendlyMessage(cause)
  } finally {
    packagePending.value = false
  }
}

function selectRevisedFile(event: Event) {
  const input = event.target as HTMLInputElement
  selectedFile.value = input.files?.[0] ?? null
  verificationResult.value = null
  errorMsg.value = null
}

async function verifyRevisedFile() {
  if (!props.jobId || !safePlan.value || !selectedFile.value) return
  verifyPending.value = true
  errorMsg.value = null
  try {
    verificationResult.value = await backend.verifyVbaChangedWorkbook(
      props.jobId,
      safePlan.value.plan.plan_id,
      selectedFile.value,
    )
    await jobStore.refreshJobs()
  } catch (cause) {
    errorMsg.value = friendlyMessage(cause)
  } finally {
    verifyPending.value = false
  }
}

async function downloadVerifiedWorkbook() {
  if (!verificationResult.value) return
  downloadPending.value = true
  errorMsg.value = null
  try {
    const blob = await backend.downloadWorkbook(verificationResult.value.new_job_id)
    saveBlob(blob, `${sourceStem()}_vba_verified.xlsm`)
  } catch (cause) {
    errorMsg.value = friendlyMessage(cause)
  } finally {
    downloadPending.value = false
  }
}
</script>

<template>
  <div v-if="safePlan && operation" class="space-y-3">
    <div class="space-y-2">
      <p class="text-sm font-medium text-(--ui-text-highlighted)">
        {{ operation.module_name }}.{{ operation.procedure_name }} を置き換えます
      </p>
      <p class="text-xs text-(--ui-text-muted)">
        Macでは変更しません。Windows Excelで原本のコピーへ適用します。
      </p>
      <details class="rounded-md border border-(--ui-border) p-2">
        <summary class="cursor-pointer text-xs text-(--ui-text-muted)">
          置換後のVBAコードを見る
        </summary>
        <pre class="mt-2 max-h-72 overflow-auto rounded bg-(--ui-bg-elevated) p-3 text-[11px] font-mono whitespace-pre-wrap">{{ operation.new_code }}</pre>
      </details>
      <UAlert
        color="warning"
        variant="subtle"
        icon="i-lucide-monitor-cog"
        title="Windows版Excelが必要です"
        description="ZIP内のREADMEを確認し、原本ではなく同梱されたコピーへ適用してください。"
      />
      <UButton
        color="primary"
        icon="i-lucide-package-down"
        size="sm"
        :loading="packagePending"
        :disabled="!props.jobId || !safePlan.can_apply"
        @click="downloadPackage"
      >
        Windows用修正パッケージをダウンロード
      </UButton>
    </div>

    <div class="rounded-md border border-(--ui-border) p-3 space-y-2">
      <p class="text-sm font-medium">Windowsで適用した後</p>
      <input
        type="file"
        accept=".xlsm"
        class="block w-full text-xs text-(--ui-text-muted) file:mr-3 file:rounded-md file:border-0 file:bg-(--ui-bg-elevated) file:px-3 file:py-2 file:text-xs file:text-(--ui-text)"
        @change="selectRevisedFile"
      >
      <UButton
        color="primary"
        variant="soft"
        icon="i-lucide-shield-check"
        size="sm"
        :loading="verifyPending"
        :disabled="!selectedFile"
        @click="verifyRevisedFile"
      >
        revised.xlsmを検証する
      </UButton>
    </div>

    <UAlert
      v-if="errorMsg"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="VBA変更を確認できませんでした"
      :description="errorMsg"
    />

    <div v-if="verificationResult" class="space-y-3">
      <VerificationGateAlert
        :report="verificationResult.verification"
        :new-job-id="verificationResult.new_job_id"
      />
      <details class="rounded-md border border-(--ui-border) p-2">
        <summary class="cursor-pointer text-xs text-(--ui-text-muted)">
          実際のVBA差分を見る
        </summary>
        <div class="mt-2">
          <WorkbookDiffView :diff="verificationResult.diff" />
        </div>
      </details>
      <UButton
        color="primary"
        icon="i-lucide-download"
        size="sm"
        :loading="downloadPending"
        @click="downloadVerifiedWorkbook"
      >
        検証済み.xlsmをダウンロード
      </UButton>
    </div>
  </div>
</template>
