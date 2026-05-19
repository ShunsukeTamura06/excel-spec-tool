<script setup lang="ts">
/**
 * チャットページ — 改修対話.
 *
 * - 上部: ヘッダー (パンくず + ファイル名 + spec/diagram へのリンク)
 * - 中央: メッセージ履歴 (スクロール)
 * - 下部: 直近ツール呼び出し + メッセージ入力欄
 */

import type { ChatMessage, ToolTraceItem } from '~/types/api'

definePageMeta({ layout: 'default' })
useHead({ title: 'チャット — Excelツール改修支援AI' })

const route = useRoute()
const backend = useBackend()
const jobStore = useJobStore()
const toast = useToast()

const jobId = computed(() => String(route.params.jobId))

// 履歴ロード
const history = ref<ChatMessage[]>([])
const loading = ref(false)
const loadError = ref('')

async function loadHistory() {
  loading.value = true
  loadError.value = ''
  try {
    history.value = await backend.getChatHistory(jobId.value)
  } catch (e) {
    loadError.value = friendlyMessage(e)
  } finally {
    loading.value = false
  }
}

// 履歴取得とジョブストア同期はナビゲーションをブロックしないよう onMounted で実行.
// 取得中は空状態 (もしくは loading 表示) を見せ、終わったら履歴を埋める.
onMounted(() => {
  if (jobStore.currentJobId !== jobId.value) jobStore.setCurrentJobId(jobId.value)
  void jobStore.refreshJobs()
  void loadHistory()
})

// 送信
const input = ref('')
const sending = ref(false)
const sendError = ref('')
const lastToolTrace = ref<ToolTraceItem[]>([])
const scrollContainer = ref<HTMLElement | null>(null)

function scrollToBottom() {
  nextTick(() => {
    const el = scrollContainer.value
    if (el) el.scrollTop = el.scrollHeight
  })
}

watch(history, () => scrollToBottom(), { flush: 'post' })
onMounted(() => scrollToBottom())

async function send() {
  const msg = input.value.trim()
  if (!msg || sending.value) return
  sending.value = true
  sendError.value = ''

  const now = new Date().toISOString()
  // optimistic: ユーザー発話を先に表示
  history.value = [...history.value, { role: 'user', content: msg, timestamp: now }]
  input.value = ''
  scrollToBottom()

  try {
    const r = await backend.chat(jobId.value, msg)
    history.value = r.history
    lastToolTrace.value = r.tool_trace ?? []
    scrollToBottom()
  } catch (e) {
    sendError.value = friendlyMessage(e)
    toast.add({
      title: '応答に失敗しました',
      description: sendError.value,
      color: 'error',
      icon: 'i-lucide-alert-triangle',
    })
    // 失敗した発話を取り消す
    history.value = history.value.slice(0, -1)
  } finally {
    sending.value = false
  }
}

function onKeydown(e: KeyboardEvent) {
  // Cmd/Ctrl + Enter で送信
  if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') {
    e.preventDefault()
    void send()
  }
}

const examples = [
  'このブックの目的を要約してください',
  'Calc シートの H2 セルを変更したら、どこに波及しますか？',
  '商品マスタの単価を変えた時の影響範囲を教えてください',
]
</script>

<template>
  <div class="flex flex-col gap-4" style="height: calc(100vh - 3rem)">
    <!-- パンくず + ヘッダー -->
    <div class="flex items-center justify-between gap-3 flex-wrap">
      <div class="flex items-center gap-2 text-sm text-(--ui-text-muted)">
        <NuxtLink to="/" class="hover:text-(--ui-primary) flex items-center gap-1">
          <UIcon name="i-lucide-home" class="size-3.5" /> ホーム
        </NuxtLink>
        <UIcon name="i-lucide-chevron-right" class="size-4" />
        <NuxtLink :to="`/spec/${jobId}`" class="hover:text-(--ui-primary)">設計書</NuxtLink>
        <UIcon name="i-lucide-chevron-right" class="size-4" />
        <span class="text-(--ui-text-highlighted)">チャット</span>
      </div>
      <UButton
        :to="`/spec/${jobId}`"
        icon="i-lucide-file-text"
        color="neutral"
        variant="soft"
        size="sm"
      >
        設計書に戻る
      </UButton>
    </div>

    <UAlert
      v-if="loadError"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="履歴取得失敗"
      :description="loadError"
    />

    <!-- 履歴エリア -->
    <UCard
      :ui="{ body: 'p-0' }"
      class="flex-1 min-h-0 overflow-hidden"
    >
      <div
        ref="scrollContainer"
        class="h-full overflow-y-auto px-4 py-6 space-y-5"
      >
        <!-- 履歴ロード中インジケータ (空状態とは別; ロード完了まで空状態を出さない) -->
        <div
          v-if="loading && history.length === 0"
          class="h-full flex items-center justify-center gap-2 text-(--ui-text-muted)"
        >
          <UIcon name="i-lucide-loader-2" class="animate-spin size-4" />
          <span class="text-sm">履歴を読み込み中...</span>
        </div>

        <!-- 空状態 (履歴ロード完了 & メッセージなし) -->
        <div
          v-else-if="!loading && history.length === 0"
          class="h-full flex flex-col items-center justify-center text-center gap-3 py-10"
        >
          <div class="size-14 rounded-2xl bg-gradient-to-br from-indigo-500 to-purple-600 text-white flex items-center justify-center">
            <UIcon name="i-lucide-sparkles" class="size-7" />
          </div>
          <div class="space-y-1">
            <p class="font-semibold text-(--ui-text-highlighted)">改修について質問しましょう</p>
            <p class="text-xs text-(--ui-text-muted) max-w-md">
              設計書とツールを根拠に、改修手順と波及範囲を提案します。
              質問例から始めるか、自由に入力してください.
            </p>
          </div>
          <div class="flex flex-wrap gap-2 justify-center max-w-xl">
            <UButton
              v-for="ex in examples"
              :key="ex"
              size="sm"
              variant="soft"
              color="neutral"
              icon="i-lucide-corner-down-left"
              @click="input = ex"
            >
              {{ ex }}
            </UButton>
          </div>
        </div>

        <!-- メッセージ -->
        <ChatMessageBubble
          v-for="(m, i) in history"
          :key="i"
          :message="m"
        />

        <!-- アシスタント応答中インジケータ -->
        <div v-if="sending" class="flex gap-3">
          <div class="size-8 rounded-full bg-gradient-to-br from-indigo-500 to-purple-600 text-white flex items-center justify-center shrink-0">
            <UIcon name="i-lucide-sparkles" class="size-4" />
          </div>
          <div class="rounded-2xl rounded-tl-sm px-4 py-3 bg-(--ui-bg-elevated) flex items-center gap-2">
            <span class="size-2 rounded-full bg-(--ui-text-muted) animate-pulse" style="animation-delay: 0ms" />
            <span class="size-2 rounded-full bg-(--ui-text-muted) animate-pulse" style="animation-delay: 150ms" />
            <span class="size-2 rounded-full bg-(--ui-text-muted) animate-pulse" style="animation-delay: 300ms" />
            <span class="text-xs text-(--ui-text-muted) ml-1">考え中…</span>
          </div>
        </div>
      </div>
    </UCard>

    <!-- 直近のツール呼び出し -->
    <ToolTraceList :items="lastToolTrace" />

    <!-- 入力欄 -->
    <UCard :ui="{ body: 'p-3' }">
      <div class="space-y-2">
        <UTextarea
          v-model="input"
          placeholder="改修したい内容や質問を入力 (Cmd/Ctrl + Enter で送信)"
          :rows="3"
          :disabled="sending"
          autoresize
          :maxrows="8"
          class="w-full"
          @keydown="onKeydown"
        />
        <div class="flex items-center justify-between gap-2">
          <p class="text-[10px] text-(--ui-text-muted) flex items-center gap-1">
            <UIcon name="i-lucide-info" class="size-3" />
            応答には数十秒かかる場合があります (LLM 呼び出し + ツール反復)。
          </p>
          <UButton
            :loading="sending"
            :disabled="!input.trim()"
            color="primary"
            icon="i-lucide-send"
            @click="send"
          >
            送信
          </UButton>
        </div>
      </div>
    </UCard>
  </div>
</template>
