<script setup lang="ts">
/**
 * フィードバック送信モーダル.
 *
 * 設計方針: 心理的負担を下げる.
 *  - 種別は icon + 短文の 3 択 (改善要望 / 不具合 / その他). デフォルト「改善要望」.
 *  - コメント以外は任意. 名前も任意 (デフォルト匿名).
 *  - 自動コンテキスト (URL / job_id) を控えめに表示し「これも一緒に送られます」を明示.
 *  - 送信後は短いトーストだけ. 画面遷移しない.
 */

import type { FeedbackKind } from '~/types/api'

const open = defineModel<boolean>('open', { default: false })

const route = useRoute()
const jobStore = useJobStore()
const backend = useBackend()
const toast = useToast()

const kind = ref<FeedbackKind>('improvement')
const comment = ref('')
const userLabel = ref('')
const sending = ref(false)

const kindOptions: { value: FeedbackKind; label: string; icon: string; color: string }[] = [
  { value: 'improvement', label: '改善要望',   icon: 'i-lucide-lightbulb',    color: 'text-amber-600' },
  { value: 'bug',         label: '不具合',     icon: 'i-lucide-bug',          color: 'text-rose-600' },
  { value: 'other',       label: 'その他',     icon: 'i-lucide-message-circle', color: 'text-sky-600' },
]

watch(open, (v) => {
  // 開く時に毎回リセット (前回入力は持ち越さない)
  if (v) {
    kind.value = 'improvement'
    comment.value = ''
    // userLabel は前回値を保持 (毎回打ち直すのは面倒)
  }
})

async function submit() {
  // コメントが空でも種別だけは送れるようにする (👎相当の用途も想定)
  sending.value = true
  try {
    await backend.submitFeedback({
      kind: kind.value,
      comment: comment.value.trim(),
      page: route.fullPath,
      job_id: jobStore.currentJobId,
      user_label: userLabel.value.trim(),
    })
    toast.add({
      title: 'フィードバックをありがとうございます',
      description: 'いただいた内容は改善に活かします。',
      color: 'success',
      icon: 'i-lucide-heart',
    })
    open.value = false
  } catch (e) {
    toast.add({
      title: '送信に失敗しました',
      description: friendlyMessage(e),
      color: 'error',
      icon: 'i-lucide-alert-triangle',
    })
  } finally {
    sending.value = false
  }
}
</script>

<template>
  <UModal v-model:open="open" :ui="{ content: 'max-w-lg' }">
    <template #content>
      <UCard>
        <template #header>
          <div class="flex items-center gap-2">
            <div class="size-9 rounded-xl bg-(--ui-primary)/10 text-(--ui-primary) flex items-center justify-center">
              <UIcon name="i-lucide-message-square-plus" class="size-5" />
            </div>
            <div>
              <p class="font-semibold">フィードバック</p>
              <p class="text-xs text-(--ui-text-muted)">
                気づいたこと・困っていることを 1 行でも構いません。お気軽にどうぞ。
              </p>
            </div>
          </div>
        </template>

        <div class="space-y-4">
          <!-- 種別 -->
          <div>
            <p class="text-xs text-(--ui-text-muted) mb-2">種別</p>
            <div class="grid grid-cols-3 gap-2">
              <button
                v-for="opt in kindOptions"
                :key="opt.value"
                type="button"
                class="flex flex-col items-center gap-1 p-3 rounded-lg border-2 transition-colors"
                :class="
                  kind === opt.value
                    ? 'border-(--ui-primary) bg-(--ui-primary)/5'
                    : 'border-(--ui-border) hover:border-(--ui-primary)/40 hover:bg-(--ui-bg-muted)'
                "
                @click="kind = opt.value"
              >
                <UIcon :name="opt.icon" :class="opt.color" class="size-5" />
                <span class="text-xs font-medium">{{ opt.label }}</span>
              </button>
            </div>
          </div>

          <!-- コメント -->
          <div>
            <p class="text-xs text-(--ui-text-muted) mb-1">
              コメント
              <span class="text-(--ui-text-muted)/60">(任意)</span>
            </p>
            <UTextarea
              v-model="comment"
              :rows="5"
              autoresize
              :maxrows="12"
              placeholder="例: 設計書のシート一覧をキーボードで切り替えられると嬉しい"
              :disabled="sending"
              class="w-full"
            />
          </div>

          <!-- 任意の自己申告 -->
          <div>
            <p class="text-xs text-(--ui-text-muted) mb-1">
              お名前
              <span class="text-(--ui-text-muted)/60">(任意・匿名 OK)</span>
            </p>
            <UInput
              v-model="userLabel"
              placeholder="例: 田中 / 業務部"
              :disabled="sending"
              class="w-full"
            />
          </div>

          <!-- 自動添付情報 -->
          <details class="text-xs text-(--ui-text-muted)">
            <summary class="cursor-pointer hover:text-(--ui-text)">
              <UIcon name="i-lucide-info" class="size-3 inline align-text-bottom" />
              一緒に送られる情報 (匿名)
            </summary>
            <div class="mt-2 pl-4 space-y-0.5">
              <p>ページ: <code class="font-mono">{{ route.fullPath }}</code></p>
              <p v-if="jobStore.currentJobId">
                関連ジョブ: <code class="font-mono">{{ jobStore.currentJobId.slice(0, 8) }}…</code>
              </p>
              <p>個人特定可能な情報は送信されません.</p>
            </div>
          </details>
        </div>

        <template #footer>
          <div class="flex items-center justify-end gap-2">
            <UButton
              variant="ghost"
              color="neutral"
              :disabled="sending"
              @click="open = false"
            >
              キャンセル
            </UButton>
            <UButton
              color="primary"
              icon="i-lucide-send"
              :loading="sending"
              @click="submit"
            >
              送信する
            </UButton>
          </div>
        </template>
      </UCard>
    </template>
  </UModal>
</template>
