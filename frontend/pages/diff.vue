<script setup lang="ts">
/**
 * 差分比較ページ (P1 安全ゲートの静的パート).
 *
 * before/after の2ジョブを選び、GET /diff で構造差分 + 波及範囲 + 既存リスクを
 * 取得して表示する。Excel COM での再計算・マクロ実行によるテストは含まない
 * (別増分、docs/VISION.ja.md 参照)。
 */

import type { WorkbookDiffData } from '~/types/api'

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
    <WorkbookDiffView v-if="diff" :diff="diff" />
  </div>
</template>
