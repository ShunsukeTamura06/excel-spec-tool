<script setup lang="ts">
/**
 * 1 件のチャットメッセージ. user / assistant / system で出し分け.
 * assistant の content は Markdown としてレンダリングする.
 *
 * アシスタント応答には 👍 / 👎 ボタンを併置. 1 クリックでフィードバック送信.
 * 心理的負担を最小化するため、コメント記入は要求しない (詳細は FAB から).
 */

import type { ChatMessage, FeedbackKind, ToolTraceItem, WorkbookData } from '~/types/api'
import { formatJstTime } from '~/utils/dateTime'

const props = defineProps<{
  message: ChatMessage
  /** チャットページから現在の job_id を受け取る (フィードバック紐付け用) */
  jobId?: string
  /** 設計書と同じ抽出済み Workbook 構造。回答内識別子の根拠表示に使う。 */
  workbook?: WorkbookData | null
}>()

const isUser = computed(() => props.message.role === 'user')
const isAssistant = computed(() => props.message.role === 'assistant')
const evidenceItems = computed<ToolTraceItem[]>(() => {
  return isAssistant.value ? props.message.tool_trace ?? [] : []
})
const hasActionableProposal = computed(() => evidenceItems.value.some(item =>
  [
    'propose_named_range_fix',
    'propose_fixed_ref_replace',
    'propose_range_expansion',
    'propose_cell_text_edits',
    'propose_vba_procedure_replace',
  ].includes(item.name),
))
const answerExpanded = ref(false)
const isLongAnswer = computed(() => {
  if (!isAssistant.value) return false
  const content = props.message.content ?? ''
  return content.length > 700 || content.split('\n').length > 14
})

const timeLabel = computed(() => {
  return formatJstTime(props.message.timestamp)
})

// ---- フィードバック (assistant 応答のみ) ----
const backend = useBackend()
const route = useRoute()
const toast = useToast()
const voted = ref<FeedbackKind | null>(null)
const sending = ref(false)

async function sendVote(kind: 'thumbs_up' | 'thumbs_down') {
  if (sending.value || voted.value) return
  sending.value = true
  try {
    await backend.submitFeedback({
      kind,
      page: route.fullPath,
      job_id: props.jobId ?? null,
      target_id: props.message.timestamp || undefined,
      target_excerpt: (props.message.content || '').slice(0, 200),
    })
    voted.value = kind
    toast.add({
      title: kind === 'thumbs_up' ? '評価ありがとうございます' : '改善のヒントになります',
      description: kind === 'thumbs_down'
        ? '詳細を書きたい場合は右下のフィードバックボタンからどうぞ.'
        : undefined,
      color: kind === 'thumbs_up' ? 'success' : 'info',
      icon: kind === 'thumbs_up' ? 'i-lucide-thumbs-up' : 'i-lucide-thumbs-down',
    })
  } catch (e) {
    toast.add({
      title: '送信に失敗しました',
      description: friendlyMessage(e),
      color: 'error',
    })
  } finally {
    sending.value = false
  }
}
</script>

<template>
  <div class="flex gap-3" :class="isUser ? 'flex-row-reverse' : 'flex-row'">
    <!-- アバター -->
    <div
      class="size-8 rounded-full flex items-center justify-center shrink-0"
      :class="[
        isUser && 'bg-(--ui-primary)/10 text-(--ui-primary)',
        isAssistant && 'bg-gradient-to-br from-indigo-500 to-purple-600 text-white',
        !isUser && !isAssistant && 'bg-(--ui-bg-muted) text-(--ui-text-muted)',
      ]"
    >
      <UIcon
        :name="
          isUser ? 'i-lucide-user'
          : isAssistant ? 'i-lucide-sparkles'
          : 'i-lucide-info'
        "
        class="size-4"
      />
    </div>

    <!-- バブル -->
    <div
      class="min-w-0"
      :class="isUser ? 'max-w-[80%] items-end' : 'flex-1 max-w-none items-start'"
    >
      <div
        class="rounded-2xl px-4 py-2.5 shadow-sm"
        :class="[
          isUser && 'bg-(--ui-primary) text-white rounded-tr-sm',
          isAssistant && 'w-full bg-(--ui-bg-elevated) text-(--ui-text) rounded-tl-sm',
          !isUser && !isAssistant && 'bg-amber-50 dark:bg-amber-950 text-amber-900 dark:text-amber-200',
        ]"
      >
        <template v-if="isAssistant">
          <div class="relative">
            <div
              :class="isLongAnswer && !answerExpanded ? 'max-h-48 overflow-hidden' : ''"
            >
              <ChatMarkdown
                :markdown="message.content"
                :evidence-items="evidenceItems"
                :workbook="workbook ?? null"
              />
            </div>
            <div
              v-if="isLongAnswer && !answerExpanded"
              class="pointer-events-none absolute inset-x-0 bottom-0 h-24 bg-gradient-to-t from-(--ui-bg-elevated) to-transparent"
            />
          </div>
          <UButton
            v-if="isLongAnswer"
            color="neutral"
            variant="soft"
            size="xs"
            class="mt-3"
            :icon="answerExpanded ? 'i-lucide-chevron-up' : 'i-lucide-chevron-down'"
            @click="answerExpanded = !answerExpanded"
          >
            {{ answerExpanded ? '折りたたむ' : '全文を表示' }}
          </UButton>
        </template>
        <p v-else class="whitespace-pre-wrap text-sm leading-relaxed">{{ message.content }}</p>
      </div>
      <div
        class="flex items-center gap-2 mt-1 px-1"
        :class="isUser ? 'justify-end' : 'justify-start'"
      >
        <p v-if="timeLabel" class="text-[10px] text-(--ui-text-muted) tabular-nums">
          {{ timeLabel }}
        </p>
        <!-- アシスタント応答の評価 (1 クリックで送信) -->
        <div v-if="isAssistant" class="flex items-center gap-1">
          <button
            type="button"
            class="p-1 rounded transition-colors"
            :class="
              voted === 'thumbs_up'
                ? 'text-emerald-600 bg-emerald-50 dark:bg-emerald-950'
                : 'text-(--ui-text-muted) hover:text-emerald-600 hover:bg-(--ui-bg-elevated)'
            "
            :disabled="sending || voted !== null"
            :aria-label="voted === 'thumbs_up' ? '良い評価済み' : '良い評価を送る'"
            :title="voted ? '送信済み' : '良かったらクリック'"
            @click="sendVote('thumbs_up')"
          >
            <UIcon
              :name="voted === 'thumbs_up' ? 'i-lucide-thumbs-up' : 'i-lucide-thumbs-up'"
              class="size-3.5"
              :class="voted === 'thumbs_up' && 'fill-current'"
            />
          </button>
          <button
            type="button"
            class="p-1 rounded transition-colors"
            :class="
              voted === 'thumbs_down'
                ? 'text-rose-600 bg-rose-50 dark:bg-rose-950'
                : 'text-(--ui-text-muted) hover:text-rose-600 hover:bg-(--ui-bg-elevated)'
            "
            :disabled="sending || voted !== null"
            :aria-label="voted === 'thumbs_down' ? 'いまいち評価済み' : 'いまいち評価を送る'"
            :title="voted ? '送信済み' : '改善が必要ならクリック'"
            @click="sendVote('thumbs_down')"
          >
            <UIcon
              name="i-lucide-thumbs-down"
              class="size-3.5"
              :class="voted === 'thumbs_down' && 'fill-current'"
            />
          </button>
        </div>
      </div>
      <ToolTraceList
        v-if="evidenceItems.length > 0"
        :items="evidenceItems"
        :title="hasActionableProposal ? '変更内容を確認' : 'この回答の根拠カード'"
        :job-id="props.jobId"
        class="mt-2"
      />
    </div>
  </div>
</template>
