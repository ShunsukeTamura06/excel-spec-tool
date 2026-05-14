<script setup lang="ts">
/**
 * 1 件のチャットメッセージ. user / assistant / system で出し分け.
 * assistant の content は Markdown としてレンダリングする.
 */

import type { ChatMessage } from '~/types/api'

const props = defineProps<{ message: ChatMessage }>()

const isUser = computed(() => props.message.role === 'user')
const isAssistant = computed(() => props.message.role === 'assistant')

const timeLabel = computed(() => {
  const t = props.message.timestamp
  return t ? t.slice(11, 16) : ''
})
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
    <div class="max-w-[80%] min-w-0" :class="isUser ? 'items-end' : 'items-start'">
      <div
        class="rounded-2xl px-4 py-2.5 shadow-sm"
        :class="[
          isUser && 'bg-(--ui-primary) text-white rounded-tr-sm',
          isAssistant && 'bg-(--ui-bg-elevated) text-(--ui-text) rounded-tl-sm',
          !isUser && !isAssistant && 'bg-amber-50 dark:bg-amber-950 text-amber-900 dark:text-amber-200',
        ]"
      >
        <div v-if="isAssistant" class="spec-prose text-sm">
          <ClientOnly>
            <MDC :value="message.content" tag="div" />
            <template #fallback>
              <pre class="whitespace-pre-wrap text-sm">{{ message.content }}</pre>
            </template>
          </ClientOnly>
        </div>
        <p v-else class="whitespace-pre-wrap text-sm leading-relaxed">{{ message.content }}</p>
      </div>
      <p
        v-if="timeLabel"
        class="text-[10px] text-(--ui-text-muted) mt-1 px-1 tabular-nums"
        :class="isUser ? 'text-right' : 'text-left'"
      >
        {{ timeLabel }}
      </p>
    </div>
  </div>
</template>
