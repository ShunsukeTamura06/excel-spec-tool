<script setup lang="ts">
/** 概要タブ. 統合設計書 Markdown をそのまま描画する. */
const props = defineProps<{ markdown: string }>()
</script>

<template>
  <UCard>
    <ClientOnly>
      <MDC :value="props.markdown" tag="div" class="spec-prose" />
    </ClientOnly>
  </UCard>
</template>
