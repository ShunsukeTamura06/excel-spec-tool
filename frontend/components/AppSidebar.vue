<script setup lang="ts">
/**
 * 左サイドバー.
 * - ブランド (アプリ名)
 * - グローバルナビ (Home / Spec / Chat). Spec/Chat はジョブ未選択時 disabled.
 * - 現在のジョブ表示
 * - ダークモード切替 + Backend 接続状態
 */

const route = useRoute()
const jobStore = useJobStore()
const colorMode = useColorMode()

// colorMode.value は SSR 時にユーザーの実際の配色を解決できず、クライアント
// 初回描画と食い違ってハイドレーション不一致 (アイコン/aria-label) を起こす。
// マウント完了までは既定 (ライト相当) を描画し、マウント後に確定させることで
// SSR と初回クライアント描画を一致させる。
const mounted = ref(false)
onMounted(() => {
  mounted.value = true
})
const isDark = computed(() => mounted.value && colorMode.value === 'dark')

const navItems = computed(() => {
  const id = jobStore.currentJobId
  return [
    {
      label: 'ホーム',
      icon: 'i-lucide-home',
      to: '/',
      active: route.path === '/',
      disabled: false,
    },
    {
      label: '設計書',
      icon: 'i-lucide-file-text',
      to: id ? `/spec/${id}` : '#',
      active: route.path.startsWith('/spec'),
      disabled: !id,
    },
    {
      label: 'チャット',
      icon: 'i-lucide-message-circle',
      to: id ? `/chat/${id}` : '#',
      active: route.path.startsWith('/chat'),
      disabled: !id,
    },
  ]
})

function toggleDark() {
  colorMode.preference = colorMode.value === 'dark' ? 'light' : 'dark'
}
</script>

<template>
  <aside class="flex flex-col w-64 shrink-0 h-screen sticky top-0 border-r border-(--ui-border) bg-(--ui-bg-elevated)/40 backdrop-blur">
    <!-- Brand -->
    <div class="px-4 py-5 border-b border-(--ui-border)">
      <NuxtLink to="/" class="flex items-center gap-2.5 group">
        <div class="size-9 rounded-xl bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center text-white shadow-sm">
          <UIcon name="i-lucide-table-2" class="size-5" />
        </div>
        <div class="flex flex-col leading-tight">
          <span class="font-semibold text-base text-(--ui-text-highlighted) leading-none">xlblueprint</span>
        </div>
      </NuxtLink>
    </div>

    <!-- Nav -->
    <nav class="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
      <template v-for="item in navItems" :key="item.label">
        <UButton
          :to="item.disabled ? undefined : item.to"
          :icon="item.icon"
          :color="item.active ? 'primary' : 'neutral'"
          :variant="item.active ? 'soft' : 'ghost'"
          :disabled="item.disabled"
          block
          class="justify-start"
        >
          {{ item.label }}
        </UButton>
      </template>

      <USeparator class="my-3" />

      <!-- Current job -->
      <div v-if="jobStore.currentJob" class="px-2 py-2 space-y-1">
        <p class="text-[11px] uppercase tracking-wide text-(--ui-text-muted)">選択中のジョブ</p>
        <p class="text-sm font-medium text-(--ui-text-highlighted) truncate" :title="jobStore.currentJob.filename">
          {{ jobStore.currentJob.filename }}
        </p>
        <JobStatusBadge :status="jobStore.currentJob.status" size="sm" />
        <p class="text-[10px] font-mono text-(--ui-text-muted) truncate">
          {{ jobStore.currentJob.job_id.slice(0, 8) }}…
        </p>
      </div>
      <div v-else class="px-2 py-3 rounded-lg bg-(--ui-bg-muted) text-xs text-(--ui-text-muted) text-center">
        ジョブ未選択
      </div>
    </nav>

    <!-- Footer -->
    <div class="px-3 py-3 border-t border-(--ui-border) flex items-center justify-between">
      <UButton
        :icon="isDark ? 'i-lucide-sun' : 'i-lucide-moon'"
        color="neutral"
        variant="ghost"
        size="sm"
        :aria-label="isDark ? 'ライトモードに切替' : 'ダークモードに切替'"
        @click="toggleDark"
      />
      <AppBackendStatus />
    </div>
  </aside>
</template>
