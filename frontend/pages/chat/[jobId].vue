<script setup lang="ts">
/** チャットページ (プレースホルダ — 次コミットで本実装). */

definePageMeta({ layout: 'default' })
useHead({ title: 'チャット — Excel 改修支援ツール' })

const route = useRoute()
const jobStore = useJobStore()
const jobId = computed(() => String(route.params.jobId))

onMounted(() => {
  if (jobId.value && jobStore.currentJobId !== jobId.value) {
    jobStore.setCurrentJobId(jobId.value)
  }
})
</script>

<template>
  <div class="space-y-4">
    <div class="flex items-center gap-2 text-sm text-(--ui-text-muted)">
      <NuxtLink to="/" class="hover:text-(--ui-primary)">ホーム</NuxtLink>
      <UIcon name="i-lucide-chevron-right" class="size-4" />
      <span>チャット</span>
    </div>
    <UAlert
      icon="i-lucide-construction"
      color="warning"
      variant="subtle"
      title="チャットページは次の段階で実装します"
      :description="`Job ID: ${jobId}`"
    />
  </div>
</template>
