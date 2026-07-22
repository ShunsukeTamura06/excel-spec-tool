<script setup lang="ts">
/** 一般ユーザー向けの根拠付き Excel 診断表示. */
import type {
  DiagnosisConfidence,
  GroundedClaim,
  WorkbookDiagnosis,
} from '~/types/api'

const props = defineProps<{ diagnosis: WorkbookDiagnosis; jobId: string }>()

const confidenceStyle: Record<DiagnosisConfidence, {
  label: string
  color: 'success' | 'warning' | 'neutral'
}> = {
  explicit: { label: 'ファイル内で確認', color: 'success' },
  inferred: { label: '構造から推定', color: 'warning' },
  unknown: { label: '不明', color: 'neutral' },
}

function evidenceLabel(claim: GroundedClaim): string {
  if (claim.evidence_ids.length === 0) return '根拠なし'
  return `根拠 ${claim.evidence_ids.join(', ')}`
}
</script>

<template>
  <div class="space-y-5">
    <UCard class="overflow-hidden">
      <div class="space-y-4">
        <div class="flex items-start justify-between gap-3 flex-wrap">
          <div class="space-y-2 max-w-3xl">
            <div class="flex items-center gap-2">
              <UIcon name="i-lucide-clipboard-check" class="size-5 text-emerald-600" />
              <span class="text-sm font-semibold text-emerald-700 dark:text-emerald-300">Excel診断</span>
            </div>
            <h2 class="text-xl font-bold text-(--ui-text-highlighted)">
              {{ props.diagnosis.headline.text }}
            </h2>
            <p class="text-sm leading-6 text-(--ui-text-muted)">
              {{ props.diagnosis.overview.text }}
            </p>
          </div>
          <UBadge
            :color="confidenceStyle[props.diagnosis.overview.confidence].color"
            variant="subtle"
          >
            {{ confidenceStyle[props.diagnosis.overview.confidence].label }}
          </UBadge>
        </div>

        <div class="grid grid-cols-2 sm:grid-cols-5 gap-2">
          <div class="rounded-lg bg-(--ui-bg-elevated) p-3 text-center">
            <p class="text-xl font-bold tabular-nums">{{ props.diagnosis.coverage.sheets }}</p>
            <p class="text-xs text-(--ui-text-muted)">シート</p>
          </div>
          <div class="rounded-lg bg-(--ui-bg-elevated) p-3 text-center">
            <p class="text-xl font-bold tabular-nums">{{ props.diagnosis.coverage.formulas }}</p>
            <p class="text-xs text-(--ui-text-muted)">計算式</p>
          </div>
          <div class="rounded-lg bg-(--ui-bg-elevated) p-3 text-center">
            <p class="text-xl font-bold tabular-nums">{{ props.diagnosis.features.length }}</p>
            <p class="text-xs text-(--ui-text-muted)">機能候補</p>
          </div>
          <div class="rounded-lg bg-(--ui-bg-elevated) p-3 text-center">
            <p class="text-xl font-bold tabular-nums">{{ props.diagnosis.coverage.controls }}</p>
            <p class="text-xs text-(--ui-text-muted)">操作ボタン等</p>
          </div>
          <div class="rounded-lg bg-(--ui-bg-elevated) p-3 text-center col-span-2 sm:col-span-1">
            <p class="text-xl font-bold tabular-nums">{{ props.diagnosis.coverage.external_dependencies }}</p>
            <p class="text-xs text-(--ui-text-muted)">外部依存</p>
          </div>
        </div>

        <div class="flex gap-2 flex-wrap">
          <UButton color="primary" icon="i-lucide-wrench" :to="`/change/${props.jobId}`">
            このExcelを直したい
          </UButton>
          <UButton
            color="neutral"
            variant="soft"
            icon="i-lucide-message-circle-question"
            :to="{ path: `/chat/${props.jobId}`, query: { request: 'このExcelの用途と使い方を、根拠を示しながら説明してください。' } }"
          >
            このExcelについて質問
          </UButton>
        </div>
      </div>
    </UCard>

    <section class="space-y-3">
      <div>
        <h3 class="text-lg font-semibold text-(--ui-text-highlighted)">このツールでできること</h3>
        <p class="text-xs text-(--ui-text-muted)">ボタン、処理、シート構造から確認できた機能候補です。</p>
      </div>
      <div v-if="props.diagnosis.features.length" class="grid grid-cols-1 lg:grid-cols-2 gap-3">
        <UCard v-for="feature in props.diagnosis.features" :key="feature.id">
          <div class="space-y-3">
            <div class="flex items-start justify-between gap-2">
              <div>
                <p class="font-semibold text-(--ui-text-highlighted)">{{ feature.name }}</p>
                <p class="text-sm text-(--ui-text-muted) mt-1">{{ feature.summary }}</p>
              </div>
              <UBadge :color="confidenceStyle[feature.confidence].color" variant="subtle" size="sm">
                {{ confidenceStyle[feature.confidence].label }}
              </UBadge>
            </div>
            <div v-if="feature.entry_points.length || feature.related_sheets.length" class="text-xs space-y-1">
              <p v-if="feature.entry_points.length">
                <span class="text-(--ui-text-muted)">始め方:</span> {{ feature.entry_points.join('、') }}
              </p>
              <p v-if="feature.related_sheets.length">
                <span class="text-(--ui-text-muted)">関係する場所:</span> {{ feature.related_sheets.join('、') }}
              </p>
            </div>
            <div class="flex items-center justify-between gap-2">
              <span class="text-[11px] font-mono text-(--ui-text-muted)">
                根拠 {{ feature.evidence_ids.join(', ') || '未特定' }}
              </span>
              <UButton size="xs" variant="soft" :to="{ path: `/change/${props.jobId}`, query: { feature: feature.id } }">
                この機能を直す
              </UButton>
            </div>
          </div>
        </UCard>
      </div>
      <UAlert
        v-else
        color="neutral"
        variant="subtle"
        icon="i-lucide-circle-help"
        title="機能の入口を特定できませんでした"
        description="シート構造は確認済みです。実際の使い方を改修相談で補足すると、対象を絞り込めます。"
      />
    </section>

    <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
      <UCard>
        <template #header><h3 class="font-semibold">入力として使われるもの</h3></template>
        <ul v-if="props.diagnosis.inputs.length" class="space-y-3">
          <li v-for="claim in props.diagnosis.inputs" :key="claim.text" class="text-sm">
            <p>{{ claim.text }}</p>
            <p class="text-[11px] text-(--ui-text-muted) mt-1">{{ evidenceLabel(claim) }}</p>
          </li>
        </ul>
        <p v-else class="text-sm text-(--ui-text-muted)">入力元を特定できませんでした。</p>
      </UCard>
      <UCard>
        <template #header><h3 class="font-semibold">結果として出るもの</h3></template>
        <ul v-if="props.diagnosis.outputs.length" class="space-y-3">
          <li v-for="claim in props.diagnosis.outputs" :key="claim.text" class="text-sm">
            <p>{{ claim.text }}</p>
            <p class="text-[11px] text-(--ui-text-muted) mt-1">{{ evidenceLabel(claim) }}</p>
          </li>
        </ul>
        <p v-else class="text-sm text-(--ui-text-muted)">出力先を特定できませんでした。</p>
      </UCard>
    </div>

    <UAlert
      v-for="warning in props.diagnosis.warnings"
      :key="warning.text"
      color="warning"
      variant="subtle"
      icon="i-lucide-triangle-alert"
      title="確認が必要です"
      :description="`${warning.text}（${evidenceLabel(warning)}）`"
    />

    <UCard v-if="props.diagnosis.external_dependencies.length">
      <template #header><h3 class="font-semibold">外部ファイル・外部データ</h3></template>
      <ul class="space-y-2 text-sm">
        <li v-for="claim in props.diagnosis.external_dependencies" :key="claim.text">
          {{ claim.text }} <span class="text-xs text-(--ui-text-muted)">— {{ evidenceLabel(claim) }}</span>
        </li>
      </ul>
    </UCard>

    <details class="rounded-xl border border-(--ui-border) bg-(--ui-bg) p-4">
      <summary class="cursor-pointer font-medium">根拠と解析できない範囲を見る</summary>
      <div class="mt-4 grid grid-cols-1 lg:grid-cols-2 gap-4">
        <div>
          <h4 class="text-sm font-semibold mb-2">根拠一覧</h4>
          <ul class="space-y-2 text-xs max-h-80 overflow-y-auto">
            <li v-for="item in props.diagnosis.evidence" :key="item.id">
              <span class="font-mono font-semibold">{{ item.id }}</span>
              <span class="text-(--ui-text-muted)"> {{ item.location }}</span> — {{ item.detail }}
            </li>
          </ul>
        </div>
        <div>
          <h4 class="text-sm font-semibold mb-2">未保証の範囲</h4>
          <ul class="list-disc pl-5 space-y-2 text-xs text-(--ui-text-muted)">
            <li v-for="item in props.diagnosis.limitations" :key="item">{{ item }}</li>
          </ul>
        </div>
      </div>
    </details>
  </div>
</template>
