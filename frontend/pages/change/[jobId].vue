<script setup lang="ts">
/** 診断済みの機能と自然文要望から改修依頼書を作る一般ユーザー導線. */
import type { ChangeBrief } from '~/types/api'

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

const selectedFeature = ref(typeof route.query.feature === 'string' ? route.query.feature : '')
const requestedOutcome = ref('')
const brief = ref<ChangeBrief | null>(null)
const creating = ref(false)
const createError = ref('')

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
  try {
    brief.value = await backend.createChangeBrief(
      jobId.value,
      requestedOutcome.value,
      selectedFeature.value || undefined,
    )
  } catch (cause) {
    createError.value = friendlyMessage(cause)
  } finally {
    creating.value = false
  }
}

function openConsultation() {
  if (!brief.value) return
  const value = brief.value
  const prompt = [
    `次の改修依頼について、根拠を確認しながら変更計画を作ってください。`,
    `対象: ${value.title}`,
    `現状: ${value.current_behavior}`,
    `実現したいこと: ${value.requested_outcome}`,
    value.affected_areas.length ? `影響候補: ${value.affected_areas.join('、')}` : '',
    value.evidence_ids.length ? `確認すべき根拠: ${value.evidence_ids.join('、')}` : '',
    `受入条件:\n- ${value.acceptance_criteria.join('\n- ')}`,
    `不明点は断定せず、確認事項として示してください。`,
  ].filter(Boolean).join('\n')
  void navigateTo({ path: `/chat/${jobId.value}`, query: { request: prompt } })
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
      <h1 class="text-2xl font-bold text-(--ui-text-highlighted)">このExcelをどう直したいですか？</h1>
      <p class="mt-2 text-sm text-(--ui-text-muted)">
        技術的な変更方法ではなく、困っていることや変更後に実現したい結果を書いてください。
      </p>
    </div>

    <UAlert
      v-if="error || createError"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="読み込みに失敗しました"
      :description="createError || friendlyMessage(error)"
    />

    <UCard v-if="diagnosis">
      <div class="space-y-5">
        <div>
          <label for="feature" class="block text-sm font-semibold mb-2">対象の機能</label>
          <select
            id="feature"
            v-model="selectedFeature"
            class="w-full rounded-lg border border-(--ui-border) bg-(--ui-bg) px-3 py-2 text-sm"
          >
            <option value="">ファイル全体／どの機能か分からない</option>
            <option v-for="feature in diagnosis.features" :key="feature.id" :value="feature.id">
              {{ feature.name }} — {{ feature.summary }}
            </option>
          </select>
        </div>
        <div>
          <label for="request" class="block text-sm font-semibold mb-2">実現したいこと</label>
          <UTextarea
            id="request"
            v-model="requestedOutcome"
            :rows="5"
            autoresize
            :maxrows="10"
            class="w-full"
            placeholder="例: 集計結果に担当部署の列を追加し、部署別に絞り込めるようにしたい"
          />
          <p class="text-xs text-(--ui-text-muted) mt-2">
            例外条件、残したい現在の動作、確認に使えるデータがあれば一緒に書くと安全性が上がります。
          </p>
        </div>
        <UButton
          color="primary"
          icon="i-lucide-clipboard-list"
          :loading="creating"
          :disabled="!requestedOutcome.trim()"
          @click="createBrief"
        >
          改修依頼を整理する
        </UButton>
      </div>
    </UCard>
    <SpecSkeleton v-else-if="pending" />

    <div v-if="brief" class="space-y-4">
      <UCard>
        <template #header>
          <div class="flex items-center justify-between gap-3 flex-wrap">
            <h2 class="text-lg font-semibold">{{ brief.title }}</h2>
            <UBadge color="warning" variant="subtle">適用前の確認が必要</UBadge>
          </div>
        </template>
        <div class="space-y-5">
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div class="rounded-lg bg-(--ui-bg-elevated) p-4">
              <p class="text-xs font-semibold text-(--ui-text-muted) mb-1">現在分かっている動作</p>
              <p class="text-sm">{{ brief.current_behavior }}</p>
            </div>
            <div class="rounded-lg bg-emerald-50 dark:bg-emerald-950 p-4">
              <p class="text-xs font-semibold text-emerald-700 dark:text-emerald-300 mb-1">変更後に実現したいこと</p>
              <p class="text-sm">{{ brief.requested_outcome }}</p>
            </div>
          </div>

          <div>
            <h3 class="text-sm font-semibold mb-2">影響を確認する場所</h3>
            <div v-if="brief.affected_areas.length" class="flex flex-wrap gap-2">
              <UBadge v-for="area in brief.affected_areas" :key="area" color="neutral" variant="subtle">
                {{ area }}
              </UBadge>
            </div>
            <p v-else class="text-sm text-(--ui-text-muted)">診断だけでは場所を絞れませんでした。</p>
            <p v-if="brief.evidence_ids.length" class="text-xs font-mono text-(--ui-text-muted) mt-2">
              根拠 {{ brief.evidence_ids.join(', ') }}
            </p>
          </div>

          <div class="grid grid-cols-1 lg:grid-cols-2 gap-5">
            <div>
              <h3 class="text-sm font-semibold mb-2">先に確認したいこと</h3>
              <ul class="list-disc pl-5 space-y-2 text-sm">
                <li v-for="item in brief.clarification_questions" :key="item">{{ item }}</li>
              </ul>
            </div>
            <div>
              <h3 class="text-sm font-semibold mb-2">完了の判定条件</h3>
              <ul class="space-y-2 text-sm">
                <li v-for="item in brief.acceptance_criteria" :key="item" class="flex gap-2">
                  <UIcon name="i-lucide-square-check-big" class="size-4 mt-0.5 text-emerald-600 shrink-0" />
                  <span>{{ item }}</span>
                </li>
              </ul>
            </div>
          </div>

          <UAlert
            color="warning"
            variant="subtle"
            icon="i-lucide-shield-alert"
            title="まだ自動変更は開始しません"
            :description="brief.automation_reason"
          />

          <div class="flex gap-2 flex-wrap">
            <UButton color="primary" icon="i-lucide-messages-square" @click="openConsultation">
              この内容で変更計画を作る
            </UButton>
            <UButton color="neutral" variant="soft" icon="i-lucide-pencil" @click="brief = null">
              要望を書き直す
            </UButton>
          </div>
        </div>
      </UCard>
    </div>
  </div>
</template>
