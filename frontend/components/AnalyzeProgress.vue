<script setup lang="ts">
/**
 * アップロード→抽出→分析 の進捗をフェーズ表示する.
 * 親が `phase` を切り替えると UI が遷移する.
 *
 * 長時間ジョブ対策として、active なフェーズの経過秒数を表示し、
 * 5 分を超えたら「長くかかっています」バナーを出す (バックエンドの
 * 10 分タイムアウトに先立ってユーザーに状況を知らせる).
 */

export type AnalyzePhase = 'idle' | 'extracting' | 'analyzing' | 'done' | 'error'

const props = defineProps<{
  phase: AnalyzePhase
  filename?: string
  errorMessage?: string
}>()

interface Step {
  key: 'extract' | 'analyze'
  title: string
  hint: string
}

const steps: Step[] = [
  {
    key: 'extract',
    title: 'ファイルの中身を確認',
    hint: 'シート、計算、操作、外部とのつながりを確認しています',
  },
  {
    key: 'analyze',
    title: 'Excel診断を作成',
    hint: '用途、機能、入力、出力、注意点を根拠と結び付けています',
  },
]

function stateOf(key: Step['key']): 'pending' | 'active' | 'done' {
  if (props.phase === 'error') return 'pending'
  if (key === 'extract') {
    if (props.phase === 'extracting') return 'active'
    if (props.phase === 'analyzing' || props.phase === 'done') return 'done'
    return 'pending'
  }
  // analyze
  if (props.phase === 'analyzing') return 'active'
  if (props.phase === 'done') return 'done'
  return 'pending'
}

// --- 経過時間トラッキング ---
const SLOW_THRESHOLD_SEC = 5 * 60 // 5 分

const elapsedSec = ref(0)
let phaseStartedAt: number | null = null
let timer: ReturnType<typeof setInterval> | null = null

function isActivePhase(p: AnalyzePhase): boolean {
  return p === 'extracting' || p === 'analyzing'
}

function stopTimer() {
  if (timer !== null) {
    clearInterval(timer)
    timer = null
  }
}

watch(
  () => props.phase,
  (p, prev) => {
    if (isActivePhase(p)) {
      // フェーズ遷移 (extracting → analyzing) では計測継続. idle から入った時だけリセット.
      if (!isActivePhase(prev ?? 'idle')) {
        phaseStartedAt = Date.now()
        elapsedSec.value = 0
        stopTimer()
        timer = setInterval(() => {
          if (phaseStartedAt !== null) {
            elapsedSec.value = Math.floor((Date.now() - phaseStartedAt) / 1000)
          }
        }, 1000)
      }
    } else {
      stopTimer()
      phaseStartedAt = null
    }
  },
  { immediate: true },
)

onBeforeUnmount(stopTimer)

const elapsedLabel = computed(() => {
  const s = elapsedSec.value
  if (s < 60) return `${s}秒`
  const m = Math.floor(s / 60)
  const rem = s % 60
  return `${m}分${rem.toString().padStart(2, '0')}秒`
})

const showSlowBanner = computed(
  () => isActivePhase(props.phase) && elapsedSec.value >= SLOW_THRESHOLD_SEC,
)
</script>

<template>
  <UCard class="w-full">
    <template #header>
      <div class="flex items-center gap-2">
        <UIcon
          :name="phase === 'error' ? 'i-lucide-alert-circle' : phase === 'done' ? 'i-lucide-check-circle-2' : 'i-lucide-loader-2'"
          :class="[
            phase === 'error' && 'text-(--ui-error)',
            phase === 'done' && 'text-(--ui-success)',
            (phase === 'extracting' || phase === 'analyzing') && 'animate-spin text-(--ui-primary)',
          ]"
          class="size-5"
        />
        <span class="font-medium">
          {{
            phase === 'error' ? '分析失敗'
            : phase === 'done' ? '分析完了'
            : '分析中…'
          }}
        </span>
        <span v-if="filename" class="text-xs text-(--ui-text-muted) ml-2 truncate">{{ filename }}</span>
        <span
          v-if="isActivePhase(phase)"
          class="text-xs text-(--ui-text-muted) ml-auto tabular-nums shrink-0"
        >
          {{ elapsedLabel }}経過
        </span>
      </div>
    </template>

    <ol class="space-y-3">
      <li
        v-for="(s, i) in steps"
        :key="s.key"
        class="flex items-start gap-3 p-3 rounded-lg transition-colors"
        :class="[
          stateOf(s.key) === 'active' && 'bg-(--ui-primary)/5',
          stateOf(s.key) === 'done' && 'opacity-70',
        ]"
      >
        <div
          class="size-7 rounded-full flex items-center justify-center text-xs font-semibold shrink-0"
          :class="[
            stateOf(s.key) === 'pending' && 'bg-(--ui-bg-muted) text-(--ui-text-muted)',
            stateOf(s.key) === 'active' && 'bg-(--ui-primary) text-white',
            stateOf(s.key) === 'done' && 'bg-(--ui-success) text-white',
          ]"
        >
          <UIcon v-if="stateOf(s.key) === 'done'" name="i-lucide-check" class="size-4" />
          <UIcon v-else-if="stateOf(s.key) === 'active'" name="i-lucide-loader-2" class="size-4 animate-spin" />
          <span v-else>{{ i + 1 }}</span>
        </div>
        <div class="flex-1 min-w-0">
          <p class="font-medium text-sm">{{ s.title }}</p>
          <p class="text-xs text-(--ui-text-muted)">{{ s.hint }}</p>
        </div>
      </li>
    </ol>

    <UAlert
      v-if="showSlowBanner"
      color="warning"
      variant="subtle"
      icon="i-lucide-clock"
      title="長くかかっています"
      :description="`${elapsedLabel}経過しました。大きなファイルや LLM 応答待ちで時間がかかることがあります。最大 10 分でタイムアウトし、その時点でエラーになります。`"
      class="mt-3"
    />

    <UAlert
      v-if="phase === 'error' && errorMessage"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      :title="errorMessage"
      class="mt-3"
    />
  </UCard>
</template>
