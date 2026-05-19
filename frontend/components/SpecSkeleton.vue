<script setup lang="ts">
/**
 * 設計書ページがデータをロード中に出すスケルトン.
 * 最終的な構造 (メトリクス + タブ) と同じ形を pulse 表示しておくことで、
 * 「ページに到達した」「これから内容が埋まる」というシグナルになり、
 * 単なる白画面より体感速度がよくなる.
 */
</script>

<template>
  <div class="space-y-5 animate-pulse">
    <!-- ヘッダー領域 -->
    <div class="flex items-start justify-between gap-4 flex-wrap">
      <div class="space-y-2">
        <div class="h-7 w-72 rounded bg-(--ui-bg-elevated)" />
        <div class="h-4 w-56 rounded bg-(--ui-bg-elevated)" />
      </div>
      <div class="flex gap-2">
        <div class="h-9 w-36 rounded-md bg-(--ui-bg-elevated)" />
        <div class="h-9 w-36 rounded-md bg-(--ui-bg-elevated)" />
      </div>
    </div>

    <!-- メトリクス (7 タイル) -->
    <div class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-7 gap-2">
      <div
        v-for="i in 7"
        :key="i"
        class="rounded-xl bg-(--ui-bg-elevated) h-[88px]"
      />
    </div>

    <!-- タブバー -->
    <div class="flex gap-1 border-b border-(--ui-border) pb-1">
      <div v-for="i in 5" :key="i" class="h-8 w-24 rounded-md bg-(--ui-bg-elevated)" />
    </div>

    <!-- タブ中身 (概要タブの想定で大きめのブロック) -->
    <div class="rounded-xl border border-(--ui-border) p-6 space-y-3">
      <div class="h-5 w-1/2 rounded bg-(--ui-bg-elevated)" />
      <div class="h-3 w-full rounded bg-(--ui-bg-elevated)" />
      <div class="h-3 w-11/12 rounded bg-(--ui-bg-elevated)" />
      <div class="h-3 w-8/12 rounded bg-(--ui-bg-elevated)" />
      <div class="h-3 w-10/12 rounded bg-(--ui-bg-elevated)" />
      <div class="h-3 w-7/12 rounded bg-(--ui-bg-elevated)" />
    </div>
  </div>
</template>
