<script setup lang="ts">
import type { JobMeta } from '~/types/api'

const open = defineModel<boolean>('open', { default: false })
const props = defineProps<{ job: JobMeta | null }>()
const emit = defineEmits<{ confirm: [job: JobMeta] }>()

const deleting = ref(false)

async function onConfirm() {
  if (!props.job) return
  deleting.value = true
  try {
    emit('confirm', props.job)
  } finally {
    deleting.value = false
    open.value = false
  }
}
</script>

<template>
  <UModal v-model:open="open" :ui="{ content: 'max-w-md' }">
    <template #content>
      <UCard>
        <template #header>
          <div class="flex items-center gap-2">
            <div class="size-9 rounded-xl bg-red-100 dark:bg-red-950 text-red-600 dark:text-red-300 flex items-center justify-center">
              <UIcon name="i-lucide-alert-triangle" class="size-5" />
            </div>
            <div>
              <p class="font-semibold">削除の確認</p>
              <p class="text-xs text-(--ui-text-muted)">この操作は取り消せません</p>
            </div>
          </div>
        </template>

        <div v-if="job" class="space-y-2 text-sm">
          <p>
            <span class="font-medium text-(--ui-text-highlighted)">{{ job.filename }}</span>
            を削除します。
          </p>
          <p class="text-xs font-mono text-(--ui-text-muted) break-all">{{ job.job_id }}</p>
          <p class="text-xs text-(--ui-text-muted)">
            アップロード原本・抽出データ・設計書・チャット履歴がすべて削除されます。
          </p>
        </div>

        <template #footer>
          <div class="flex justify-end gap-2">
            <UButton color="neutral" variant="ghost" :disabled="deleting" @click="open = false">
              キャンセル
            </UButton>
            <UButton color="error" :loading="deleting" icon="i-lucide-trash-2" @click="onConfirm">
              はい、削除する
            </UButton>
          </div>
        </template>
      </UCard>
    </template>
  </UModal>
</template>
