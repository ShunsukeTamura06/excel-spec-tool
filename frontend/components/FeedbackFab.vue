<script setup lang="ts">
/**
 * 全画面共通の「フィードバック」FAB. 画面右下に常時表示する.
 * クリックすると FeedbackModal を開く.
 *
 * 表示位置: 画面右下固定 (z-50). スクロールに追従.
 * 心理的負担を下げるため、デフォルトでは控えめなアイコンのみ. ホバー時に
 * ラベルが展開する.
 */

const open = ref(false)
</script>

<template>
  <div>
    <button
      type="button"
      aria-label="フィードバックを送る"
      class="group fixed bottom-4 right-4 z-50 flex items-center gap-2 px-3 py-2.5 rounded-full bg-(--ui-primary) text-white shadow-lg hover:shadow-xl hover:scale-105 active:scale-95 transition-all"
      @click="open = true"
    >
      <UIcon name="i-lucide-message-square-plus" class="size-5" />
      <span class="text-xs font-medium hidden group-hover:inline whitespace-nowrap">
        フィードバック
      </span>
    </button>
    <FeedbackModal v-model:open="open" />
  </div>
</template>
