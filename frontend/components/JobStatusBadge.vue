<script setup lang="ts">
import type { JobStatus } from '~/types/api'

const props = defineProps<{
  status: JobStatus
  size?: 'sm' | 'md'
}>()

const styleFor: Record<JobStatus, { color: 'neutral' | 'info' | 'success' | 'error'; label: string; icon: string }> = {
  uploaded:  { color: 'neutral', label: 'アップロード済', icon: 'i-lucide-cloud-upload' },
  extracted: { color: 'info',    label: '抽出済',         icon: 'i-lucide-scan-text' },
  analyzed:  { color: 'success', label: '分析完了',       icon: 'i-lucide-check-circle-2' },
  failed:    { color: 'error',   label: '失敗',           icon: 'i-lucide-alert-circle' },
}

const cfg = computed(() => styleFor[props.status])
</script>

<template>
  <UBadge :color="cfg.color" variant="subtle" :size="props.size ?? 'md'" :icon="cfg.icon">
    {{ cfg.label }}
  </UBadge>
</template>
