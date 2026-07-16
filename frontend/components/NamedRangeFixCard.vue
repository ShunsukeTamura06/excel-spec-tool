<script setup lang="ts">
/**
 * propose_named_range_fix の結果を表示し、適用するボタンを提供する.
 *
 * 提案 (試算のみ、ファイル未変更) を見せた上で、ユーザーが明示的にボタンを押した
 * ときだけ POST /jobs/{jobId}/named-range-fix を呼んで実際にファイルへ適用する
 * (docs/VISION.ja.md §4.2「黙って変更しない」— LLM 自身はここを呼べない)。
 */

import type {
  NamedRangeDiffItem,
  ToolTraceItem,
  VerificationReportData,
  WorkbookDiffData,
} from '~/types/api'

const props = defineProps<{
  item: ToolTraceItem
  jobId?: string
}>()

const backend = useBackend()
const jobStore = useJobStore()

interface ProposalResult {
  proposal?: WorkbookDiffData
  note?: string
}

const args = computed(() => props.item.arguments as { name?: string; new_refers_to?: string })
const result = computed(() => props.item.result as ProposalResult | undefined)
const proposal = computed(() => result.value?.proposal)
const namedRangeChange = computed<NamedRangeDiffItem | undefined>(
  () => proposal.value?.named_ranges[0],
)

const pending = ref(false)
const errorMsg = ref<string | null>(null)
const applyResult = ref<{
  newJobId: string
  diff: WorkbookDiffData
  verification: VerificationReportData
} | null>(null)

async function apply() {
  const name = args.value.name
  const newRefersTo = args.value.new_refers_to
  if (!props.jobId || !name || !newRefersTo) return

  pending.value = true
  errorMsg.value = null
  try {
    const res = await backend.applyNamedRangeFix(props.jobId, {
      name,
      new_refers_to: newRefersTo,
    })
    applyResult.value = {
      newJobId: res.new_job_id,
      diff: res.diff,
      verification: res.verification,
    }
    await jobStore.refreshJobs()
  } catch (e) {
    errorMsg.value = friendlyMessage(e)
  } finally {
    pending.value = false
  }
}
</script>

<template>
  <div class="space-y-3">
    <div v-if="namedRangeChange" class="rounded-md border border-(--ui-border) bg-(--ui-bg) p-3 space-y-2">
      <div class="flex items-center gap-2 text-sm">
        <UIcon name="i-lucide-tag" class="size-4 text-(--ui-primary)" />
        <span class="font-mono font-medium text-(--ui-text-highlighted)">{{ namedRangeChange.name }}</span>
      </div>
      <p class="text-xs text-(--ui-text-muted) font-mono">
        {{ namedRangeChange.before_refers_to ?? '(なし)' }} → {{ namedRangeChange.after_refers_to ?? '(なし)' }}
      </p>
      <p v-if="proposal && proposal.blast_radius.length > 0" class="text-xs text-amber-600 dark:text-amber-400">
        波及範囲: {{ proposal.blast_radius.length }} 箇所がこの参照先を利用しています
      </p>
      <p v-else class="text-xs text-(--ui-text-muted)">波及範囲: 検出されませんでした</p>

      <p class="text-[11px] text-(--ui-text-muted)">
        これは試算結果です。ファイルはまだ変更されていません。
      </p>

      <UButton
        v-if="!applyResult"
        icon="i-lucide-check"
        color="primary"
        size="sm"
        :loading="pending"
        :disabled="!props.jobId"
        @click="apply"
      >
        適用する
      </UButton>
    </div>

    <UAlert
      v-if="errorMsg"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="適用に失敗しました"
      :description="errorMsg"
    />

    <div v-if="applyResult" class="space-y-2">
      <VerificationGateAlert
        :report="applyResult.verification"
        :new-job-id="applyResult.newJobId"
      />
      <WorkbookDiffView :diff="applyResult.diff" />
    </div>
  </div>
</template>
