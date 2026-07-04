<script setup lang="ts">
/**
 * propose_fixed_ref_replace / propose_range_expansion の結果を表示し、
 * 適用するボタンを提供する.
 *
 * 提案 (試算のみ、ファイル未変更) を見せた上で、ユーザーが明示的にボタンを押した
 * ときだけ POST /jobs/{jobId}/formula-fix を呼んで実際にファイルへ適用する
 * (docs/VISION.ja.md §4.2「黙って変更しない」— LLM 自身はここを呼べない)。
 */

import type { FormulaFixKind, ToolTraceItem, WorkbookDiffData } from '~/types/api'

const props = defineProps<{
  item: ToolTraceItem
  jobId?: string
}>()

const backend = useBackend()
const jobStore = useJobStore()

interface ProposalResult {
  proposal?: WorkbookDiffData
  changed_formula_count?: number
  note?: string
}

const kind = computed<FormulaFixKind>(() =>
  props.item.name === 'propose_range_expansion' ? 'range_expansion' : 'fixed_ref_replace',
)
const args = computed(() => {
  const a = props.item.arguments as {
    old_ref?: string
    new_ref?: string
    old_range?: string
    new_range?: string
  }
  return {
    oldRef: a.old_ref ?? a.old_range,
    newRef: a.new_ref ?? a.new_range,
  }
})
const result = computed(() => props.item.result as ProposalResult | undefined)
const proposal = computed(() => result.value?.proposal)
const changedCells = computed(() => proposal.value?.cells ?? [])

const kindLabel = computed(() =>
  kind.value === 'range_expansion' ? '数式範囲拡張' : '固定参照置換',
)

const MAX_SHOWN_CELLS = 10

const pending = ref(false)
const errorMsg = ref<string | null>(null)
const applyResult = ref<{ newJobId: string; diff: WorkbookDiffData } | null>(null)

async function apply() {
  const oldRef = args.value.oldRef
  const newRef = args.value.newRef
  if (!props.jobId || !oldRef || !newRef) return

  pending.value = true
  errorMsg.value = null
  try {
    const res = await backend.applyFormulaFix(props.jobId, {
      kind: kind.value,
      old_ref: oldRef,
      new_ref: newRef,
    })
    applyResult.value = { newJobId: res.new_job_id, diff: res.diff }
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
    <div v-if="proposal" class="rounded-md border border-(--ui-border) bg-(--ui-bg) p-3 space-y-2">
      <div class="flex items-center gap-2 text-sm">
        <UIcon name="i-lucide-replace" class="size-4 text-(--ui-primary)" />
        <span class="font-medium text-(--ui-text-highlighted)">{{ kindLabel }}</span>
      </div>
      <p class="text-xs text-(--ui-text-muted) font-mono">
        {{ args.oldRef }} → {{ args.newRef }}
      </p>

      <div class="text-xs text-(--ui-text-muted) space-y-1">
        <p class="font-medium text-(--ui-text)">
          変更される数式: {{ changedCells.length }} 箇所
        </p>
        <div
          v-for="cell in changedCells.slice(0, MAX_SHOWN_CELLS)"
          :key="`${cell.sheet}!${cell.coord}`"
          class="font-mono"
        >
          <span class="text-(--ui-text)">{{ cell.sheet }}!{{ cell.coord }}</span>:
          {{ cell.before_formula }} → {{ cell.after_formula }}
        </div>
        <p v-if="changedCells.length > MAX_SHOWN_CELLS">
          ... 他 {{ changedCells.length - MAX_SHOWN_CELLS }} 箇所
        </p>
      </div>

      <p v-if="proposal.blast_radius.length > 0" class="text-xs text-amber-600 dark:text-amber-400">
        波及範囲: {{ proposal.blast_radius.length }} 箇所が変更対象のセルを参照しています
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
      <UAlert
        color="success"
        variant="subtle"
        icon="i-lucide-check-circle-2"
        :description="`新しいジョブ (${applyResult.newJobId}) を作成しました。以下が自己検証の結果です。`"
        title="適用完了"
      />
      <WorkbookDiffView :diff="applyResult.diff" />
    </div>
  </div>
</template>
