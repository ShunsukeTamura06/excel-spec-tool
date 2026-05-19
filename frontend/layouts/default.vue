<script setup lang="ts">
/** デフォルトレイアウト. 左サイドバー + 右コンテンツ. */
</script>

<template>
  <div class="min-h-screen flex bg-(--ui-bg)">
    <AppSidebar />
    <main class="flex-1 min-w-0">
      <div class="max-w-7xl mx-auto px-6 py-6">
        <slot />
      </div>
    </main>
    <!-- 全ページ共通: フィードバック FAB (右下固定). -->
    <FeedbackFab />
  </div>
</template>
