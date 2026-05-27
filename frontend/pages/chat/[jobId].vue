<script setup lang="ts">
/**
 * チャットページ — 改修対話.
 *
 * - 左: 同じ Excel ジョブ配下の相談セッション一覧
 * - 中央: 選択中セッションのメッセージ履歴
 * - 下部: メッセージ入力欄
 */

import type { ChatMessage, ChatProgressEvent, ChatSessionMeta, WorkbookData } from '~/types/api'

definePageMeta({ layout: 'default' })
useHead({ title: 'チャット — xlblueprint' })

const route = useRoute()
const backend = useBackend()
const jobStore = useJobStore()
const toast = useToast()

const jobId = computed(() => String(route.params.jobId))

const sessions = ref<ChatSessionMeta[]>([])
const activeSessionId = ref('default')
const showArchived = ref(false)
const history = ref<ChatMessage[]>([])
const loading = ref(false)
const loadError = ref('')
const workbook = ref<WorkbookData | null>(null)

const activeSession = computed(() => {
  return sessions.value.find(s => s.session_id === activeSessionId.value) ?? null
})

const visibleSessions = computed(() => {
  return sessions.value.filter(s => showArchived.value || !s.archived)
})

function sessionTime(session: ChatSessionMeta) {
  return session.updated_at ? session.updated_at.slice(0, 16).replace('T', ' ') : ''
}

function sessionPreview(session: ChatSessionMeta) {
  if (session.last_message_preview) return session.last_message_preview
  return session.message_count > 0 ? `${session.message_count}件のメッセージ` : 'まだメッセージはありません'
}

async function updateSessionQuery(sessionId: string) {
  await navigateTo(
    { path: route.path, query: { ...route.query, session: sessionId } },
    { replace: true },
  )
}

async function refreshSessionsOnly() {
  sessions.value = await backend.listChatSessions(jobId.value, true)
}

async function loadHistory() {
  history.value = await backend.getChatHistory(jobId.value, activeSessionId.value)
}

async function loadWorkbook() {
  try {
    workbook.value = await backend.getWorkbook(jobId.value)
  } catch {
    workbook.value = null
  }
}

async function loadChatState() {
  loading.value = true
  loadError.value = ''
  try {
    await refreshSessionsOnly()
    const querySession = typeof route.query.session === 'string' ? route.query.session : ''
    const activeCandidates = sessions.value.filter(s => !s.archived)
    const selected = sessions.value.find(s => s.session_id === querySession)
      ?? activeCandidates[0]
      ?? sessions.value[0]
    activeSessionId.value = selected?.session_id ?? 'default'
    if (route.query.session !== activeSessionId.value) {
      await updateSessionQuery(activeSessionId.value)
    }
    await Promise.all([
      loadHistory(),
      loadWorkbook(),
    ])
  } catch (e) {
    loadError.value = friendlyMessage(e)
  } finally {
    loading.value = false
  }
}

onMounted(() => {
  if (jobStore.currentJobId !== jobId.value) jobStore.setCurrentJobId(jobId.value)
  void jobStore.refreshJobs()
  void loadChatState()
})

async function selectSession(session: ChatSessionMeta) {
  activeSessionId.value = session.session_id
  await updateSessionQuery(session.session_id)
  loading.value = true
  loadError.value = ''
  try {
    await loadHistory()
  } catch (e) {
    loadError.value = friendlyMessage(e)
  } finally {
    loading.value = false
  }
}

async function createSession() {
  try {
    const session = await backend.createChatSession(jobId.value)
    await refreshSessionsOnly()
    await selectSession(session)
  } catch (e) {
    toast.add({
      title: '相談の作成に失敗しました',
      description: friendlyMessage(e),
      color: 'error',
      icon: 'i-lucide-alert-triangle',
    })
  }
}

async function archiveCurrentSession() {
  const session = activeSession.value
  if (!session) return
  try {
    await backend.updateChatSession(jobId.value, session.session_id, { archived: true })
    await refreshSessionsOnly()
    const next = sessions.value.find(s => !s.archived)
    if (next) {
      await selectSession(next)
    } else {
      await createSession()
    }
    toast.add({ title: '相談をアーカイブしました', color: 'neutral', icon: 'i-lucide-archive' })
  } catch (e) {
    toast.add({
      title: 'アーカイブに失敗しました',
      description: friendlyMessage(e),
      color: 'error',
      icon: 'i-lucide-alert-triangle',
    })
  }
}

async function restoreSession(session: ChatSessionMeta) {
  try {
    const restored = await backend.updateChatSession(jobId.value, session.session_id, {
      archived: false,
    })
    await refreshSessionsOnly()
    await selectSession(restored)
  } catch (e) {
    toast.add({
      title: '復元に失敗しました',
      description: friendlyMessage(e),
      color: 'error',
      icon: 'i-lucide-alert-triangle',
    })
  }
}

const editingTitle = ref(false)
const titleDraft = ref('')

function beginTitleEdit() {
  titleDraft.value = activeSession.value?.title ?? ''
  editingTitle.value = true
}

async function saveTitle() {
  const session = activeSession.value
  const title = titleDraft.value.trim()
  if (!session || !title) {
    editingTitle.value = false
    return
  }
  try {
    await backend.updateChatSession(jobId.value, session.session_id, { title })
    await refreshSessionsOnly()
  } catch (e) {
    toast.add({
      title: 'タイトル変更に失敗しました',
      description: friendlyMessage(e),
      color: 'error',
      icon: 'i-lucide-alert-triangle',
    })
  } finally {
    editingTitle.value = false
  }
}

// 送信
const input = ref('')
const sending = ref(false)
const sendError = ref('')
const scrollContainer = ref<HTMLElement | null>(null)
type ProgressItem = {
  id: string
  message: string
  kind: 'status' | 'tool_start' | 'tool_result'
}
const progressItems = ref<ProgressItem[]>([])
const currentProgress = computed(() => {
  return progressItems.value.at(-1)?.message ?? '回答準備中'
})

function addProgress(kind: ProgressItem['kind'], message: string) {
  if (!message) return
  const previous = progressItems.value.at(-1)
  if (previous?.kind === kind && previous.message === message) return
  progressItems.value = [
    ...progressItems.value.slice(-5),
    { id: `${Date.now()}-${progressItems.value.length}`, kind, message },
  ]
  scrollToBottom()
}

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
  progressItems.value = [{ id: `${Date.now()}-0`, kind: 'status', message: '回答準備中' }]

  const now = new Date().toISOString()
  // optimistic: ユーザー発話を先に表示
  history.value = [...history.value, { role: 'user', content: msg, timestamp: now }]
  input.value = ''
  scrollToBottom()

  try {
    const r = await backend.chatStream(jobId.value, msg, activeSessionId.value, {
      onEvent(event, data) {
        if (event === 'status' || event === 'tool_start' || event === 'tool_result') {
          addProgress(event, (data as ChatProgressEvent).message)
        }
      },
    })
    history.value = r.history
    await refreshSessionsOnly()
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
    progressItems.value = []
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
  <div class="h-[calc(100vh-3rem)] min-h-0 overflow-hidden flex flex-col gap-3">
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

    <div class="flex-1 min-h-0 flex flex-col lg:flex-row gap-3">
      <UCard
        :ui="{ body: 'p-2 sm:p-3 h-full min-h-0' }"
        class="h-32 sm:h-40 lg:h-full lg:w-72 shrink-0 min-h-0 overflow-hidden"
      >
        <div class="h-full flex flex-col gap-3">
          <div class="flex items-center justify-between gap-2">
            <h2 class="text-sm font-semibold text-(--ui-text-highlighted)">相談</h2>
            <UButton
              icon="i-lucide-plus"
              color="primary"
              variant="soft"
              size="xs"
              @click="createSession"
            >
              新しい相談
            </UButton>
          </div>
          <UButton
            :icon="showArchived ? 'i-lucide-eye-off' : 'i-lucide-archive'"
            color="neutral"
            variant="ghost"
            size="xs"
            block
            @click="showArchived = !showArchived"
          >
            {{ showArchived ? 'アーカイブを隠す' : 'アーカイブ済みを表示' }}
          </UButton>

          <div class="flex-1 min-h-0 overflow-y-auto space-y-2 pr-1">
            <div
              v-for="session in visibleSessions"
              :key="session.session_id"
              class="rounded-lg border p-2 transition-colors"
              :class="[
                session.session_id === activeSessionId
                  ? 'border-(--ui-primary) bg-(--ui-primary)/5'
                  : 'border-(--ui-border) hover:bg-(--ui-bg-muted)/50',
                session.archived && 'opacity-70',
              ]"
            >
              <button class="w-full text-left min-w-0" @click="selectSession(session)">
                <div class="flex items-center gap-2 min-w-0">
                  <UIcon
                    :name="session.archived ? 'i-lucide-archive' : 'i-lucide-message-square'"
                    class="size-3.5 text-(--ui-text-muted) shrink-0"
                  />
                  <p class="text-sm font-medium text-(--ui-text-highlighted) truncate">
                    {{ session.title }}
                  </p>
                </div>
                <p class="text-xs text-(--ui-text-muted) mt-1 line-clamp-2">
                  {{ sessionPreview(session) }}
                </p>
                <div class="flex items-center justify-between gap-2 mt-2">
                  <span class="text-[10px] text-(--ui-text-muted)">{{ sessionTime(session) }}</span>
                  <span class="text-[10px] text-(--ui-text-muted)">{{ session.message_count }}件</span>
                </div>
              </button>
              <div v-if="session.archived" class="mt-2">
                <UButton
                  icon="i-lucide-rotate-ccw"
                  color="neutral"
                  variant="soft"
                  size="xs"
                  block
                  @click="restoreSession(session)"
                >
                  復元
                </UButton>
              </div>
            </div>
          </div>
        </div>
      </UCard>

      <div class="flex-1 h-full min-h-0 flex flex-col gap-3">
        <UCard :ui="{ body: 'p-2 sm:p-3' }">
          <div class="flex items-center justify-between gap-3 flex-wrap">
            <div class="min-w-0 flex-1">
              <div v-if="editingTitle" class="flex items-center gap-2">
                <UInput
                  v-model="titleDraft"
                  size="sm"
                  class="max-w-md"
                  autofocus
                  @keydown.enter.prevent="saveTitle"
                  @keydown.esc.prevent="editingTitle = false"
                />
                <UButton icon="i-lucide-check" color="primary" size="xs" @click="saveTitle" />
              </div>
              <div v-else class="min-w-0">
                <div class="flex items-center gap-2 min-w-0">
                  <p class="font-semibold text-(--ui-text-highlighted) truncate">
                    {{ activeSession?.title ?? '新しい相談' }}
                  </p>
                  <UBadge v-if="activeSession?.archived" color="neutral" variant="soft" size="sm">
                    アーカイブ済み
                  </UBadge>
                </div>
                <p class="hidden sm:block text-xs text-(--ui-text-muted) mt-1">
                  この相談は現在のExcel解析結果に紐づきます。
                </p>
              </div>
            </div>
            <div class="flex items-center gap-2 shrink-0">
              <UButton
                icon="i-lucide-pencil"
                color="neutral"
                variant="ghost"
                size="xs"
                :disabled="!activeSession"
                @click="beginTitleEdit"
              >
                名前を変更
              </UButton>
              <UButton
                icon="i-lucide-archive"
                color="neutral"
                variant="ghost"
                size="xs"
                :disabled="!activeSession || activeSession.archived"
                @click="archiveCurrentSession"
              >
                アーカイブ
              </UButton>
            </div>
          </div>
        </UCard>

        <!-- 履歴エリア -->
        <UCard
          :ui="{ body: 'p-0 h-full min-h-0 overflow-hidden' }"
          class="flex-1 min-h-0 overflow-hidden"
        >
          <div
            ref="scrollContainer"
            class="h-full min-h-0 overflow-y-auto px-4 py-6 space-y-5"
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
              <!-- LLM 未設定の場合、ここで明示的にアナウンスする.
                   配置していないと「チャットが壊れている」と誤解されやすい. -->
              <LlmOnboardingCard class="mt-4 max-w-xl text-left" />
            </div>

            <!-- メッセージ -->
            <ChatMessageBubble
              v-for="(m, i) in history"
              :key="i"
              :message="m"
              :job-id="jobId"
              :workbook="workbook"
            />

            <!-- アシスタント応答中インジケータ -->
            <div v-if="sending" class="flex gap-3">
              <div class="size-8 rounded-full bg-gradient-to-br from-indigo-500 to-purple-600 text-white flex items-center justify-center shrink-0">
                <UIcon name="i-lucide-sparkles" class="size-4" />
              </div>
              <div class="rounded-2xl rounded-tl-sm px-4 py-3 bg-(--ui-bg-elevated) min-w-0 flex-1 max-w-2xl">
                <div class="flex items-center gap-2 text-sm text-(--ui-text)">
                  <UIcon name="i-lucide-loader-circle" class="size-4 animate-spin shrink-0" />
                  <span class="truncate">{{ currentProgress }}</span>
                </div>
                <div v-if="progressItems.length > 1" class="mt-2 space-y-1">
                  <div
                    v-for="item in progressItems.slice(-3)"
                    :key="item.id"
                    class="flex items-center gap-2 text-xs text-(--ui-text-muted)"
                  >
                    <UIcon
                      :name="item.kind === 'tool_result' ? 'i-lucide-check' : 'i-lucide-circle-dot'"
                      class="size-3 shrink-0"
                    />
                    <span class="truncate">{{ item.message }}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </UCard>

        <!-- 入力欄 -->
        <UCard :ui="{ body: 'p-2 sm:p-3' }">
          <div class="space-y-2">
            <UTextarea
              v-model="input"
              placeholder="改修したい内容や質問を入力 (Cmd/Ctrl + Enter で送信)"
              :rows="3"
              :disabled="sending || activeSession?.archived"
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
                :disabled="!input.trim() || activeSession?.archived"
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
    </div>
  </div>
</template>
