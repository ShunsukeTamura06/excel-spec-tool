<script setup lang="ts">
/** Excelについて尋ねる代表的な質問を、編集可能な下書きとして提供する. */

const emit = defineEmits<{
  select: [prompt: string]
}>()

const templates = [
  {
    id: 'overview',
    label: '全体像',
    icon: 'i-lucide-telescope',
    prompt: 'このExcelは何のためのツールですか？主要な機能と全体の流れを、根拠を示しながら説明してください。',
  },
  {
    id: 'usage',
    label: '使い方',
    icon: 'i-lucide-mouse-pointer-click',
    prompt: 'このExcelを初めて使う人向けに、入力から結果確認までの操作手順を説明してください。不明な手順は推測せず明示してください。',
  },
  {
    id: 'input-output',
    label: '入力と出力',
    icon: 'i-lucide-arrow-left-right',
    prompt: 'このExcelでは、何をどこに入力し、最終的にどこへ何が出力されますか？関係するシートと根拠を示してください。',
  },
  {
    id: 'value-source',
    label: '数字の根拠',
    icon: 'i-lucide-binary',
    prompt: '「確認したいシート名・セル番地」の数字は、どの入力や計算から作られていますか？参照関係を順番に説明してください。',
  },
  {
    id: 'feature',
    label: '機能を調べる',
    icon: 'i-lucide-workflow',
    prompt: '「確認したいボタン・シート・機能」は何をするものですか？起動条件、入力、処理、出力、関連箇所を説明してください。',
  },
  {
    id: 'risks',
    label: '注意点',
    icon: 'i-lucide-shield-alert',
    prompt: 'このExcelを使用・引き継ぎするときに注意すべき点は何ですか？外部依存、動的処理、壊れやすい箇所、不明点を分けて説明してください。',
  },
] as const
</script>

<template>
  <div class="space-y-2">
    <div class="flex items-center gap-2 text-xs text-(--ui-text-muted)">
      <UIcon name="i-lucide-message-circle-question" class="size-3.5" />
      <span>質問テンプレート</span>
      <span class="hidden sm:inline">— 選択後に内容を書き換えられます</span>
    </div>
    <div class="flex flex-wrap gap-1.5">
      <UButton
        v-for="item in templates"
        :key="item.id"
        :icon="item.icon"
        color="neutral"
        variant="soft"
        size="xs"
        @click="emit('select', item.prompt)"
      >
        {{ item.label }}
      </UButton>
    </div>
  </div>
</template>
