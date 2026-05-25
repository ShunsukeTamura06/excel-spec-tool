<script setup lang="ts">
import type { JobMeta } from '~/types/api'

const props = defineProps<{
  job: JobMeta
  isCurrent: boolean
}>()
const emit = defineEmits<{
  open: [jobId: string]
  delete: [job: JobMeta]
}>()

const created = computed(() => {
  // ISO 文字列 → 「YYYY-MM-DD HH:MM」表記
  const s = props.job.created_at
  return s ? s.slice(0, 16).replace('T', ' ') : ''
})

const shortHash = computed(() => props.job.file_sha256?.slice(0, 10) ?? '')
</script>

<template>
  <UCard
    class="transition-all hover:shadow-md"
    :class="[isCurrent && 'ring-2 ring-(--ui-primary) ring-offset-1 ring-offset-(--ui-bg)']"
  >
    <div class="flex items-start gap-3">
      <div class="size-10 rounded-xl bg-emerald-100 dark:bg-emerald-950 text-emerald-700 dark:text-emerald-300 flex items-center justify-center shrink-0">
        <UIcon name="i-lucide-file-spreadsheet" class="size-5" />
      </div>
      <div class="flex-1 min-w-0">
        <div class="flex items-center gap-2 flex-wrap">
          <p class="font-medium text-(--ui-text-highlighted) truncate" :title="job.filename">
            {{ job.filename }}
          </p>
          <UBadge v-if="isCurrent" color="primary" variant="solid" size="sm">選択中</UBadge>
        </div>
        <div class="flex items-center gap-2 mt-1 flex-wrap">
          <JobStatusBadge :status="job.status" size="sm" />
          <span class="text-xs text-(--ui-text-muted)">{{ created }}</span>
        </div>
        <p class="text-[10px] font-mono text-(--ui-text-muted) mt-1 truncate">
          {{ shortHash ? `sha256:${shortHash}` : job.job_id }}
        </p>
      </div>
      <div class="flex flex-col gap-1.5 shrink-0">
        <UButton
          v-if="!isCurrent"
          icon="i-lucide-folder-open"
          color="primary"
          variant="soft"
          size="xs"
          @click="emit('open', job.job_id)"
        >
          開く
        </UButton>
        <UButton
          v-else
          icon="i-lucide-check"
          color="primary"
          variant="solid"
          size="xs"
          disabled
        >
          選択中
        </UButton>
        <UButton
          icon="i-lucide-trash-2"
          color="neutral"
          variant="ghost"
          size="xs"
          aria-label="削除"
          @click="emit('delete', job)"
        />
      </div>
    </div>
  </UCard>
</template>
