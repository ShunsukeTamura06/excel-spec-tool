<script setup lang="ts">
/** Backend (FastAPI) への到達性インジケータ. 30 秒ごとに /health を叩く. */

const ok = ref<boolean | null>(null)
const backend = useBackend()
let timer: ReturnType<typeof setInterval> | null = null

async function check() {
  ok.value = await backend.health()
}

onMounted(() => {
  void check()
  timer = setInterval(check, 30_000)
})

onUnmounted(() => {
  if (timer) clearInterval(timer)
})

const cfg = computed(() => {
  if (ok.value === null) return { color: 'neutral', label: '確認中', dot: 'bg-gray-400' }
  if (ok.value) return { color: 'success', label: '接続中', dot: 'bg-emerald-500' }
  return { color: 'error', label: '切断', dot: 'bg-red-500' }
})
</script>

<template>
  <div class="flex items-center gap-1.5 text-[10px] text-(--ui-text-muted)" :title="`サービス: ${cfg.label}`">
    <span class="inline-block size-2 rounded-full" :class="cfg.dot" />
    <span>{{ cfg.label }}</span>
  </div>
</template>
