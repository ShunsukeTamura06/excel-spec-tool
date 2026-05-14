<script setup lang="ts">
/**
 * .xlsm / .xls / .xlsx ファイル受付用ドロップゾーン.
 * クリックでファイル選択ダイアログ、またはドラッグ&ドロップ.
 */

const props = withDefaults(
  defineProps<{ disabled?: boolean; accept?: string }>(),
  {
    disabled: false,
    accept: '.xlsm,.xls,.xlsx',
  },
)
const emit = defineEmits<{ select: [file: File] }>()

const dragOver = ref(false)
const fileInput = ref<HTMLInputElement | null>(null)

function pickFile() {
  if (props.disabled) return
  fileInput.value?.click()
}

function onChange(e: Event) {
  const t = e.target as HTMLInputElement
  const f = t.files?.[0]
  if (f) emit('select', f)
  // 同じファイルを再選択できるようリセット
  t.value = ''
}

function onDrop(e: DragEvent) {
  dragOver.value = false
  if (props.disabled) return
  const f = e.dataTransfer?.files?.[0]
  if (f) emit('select', f)
}
</script>

<template>
  <div
    role="button"
    tabindex="0"
    class="group relative w-full rounded-2xl border-2 border-dashed transition-all p-10 text-center cursor-pointer"
    :class="[
      dragOver
        ? 'border-(--ui-primary) bg-(--ui-primary)/5'
        : 'border-(--ui-border) hover:border-(--ui-primary)/50 hover:bg-(--ui-bg-muted)/40',
      disabled && 'opacity-50 cursor-not-allowed',
    ]"
    @click="pickFile"
    @keydown.enter.prevent="pickFile"
    @keydown.space.prevent="pickFile"
    @dragover.prevent="!disabled && (dragOver = true)"
    @dragleave.prevent="dragOver = false"
    @drop.prevent="onDrop"
  >
    <div class="flex flex-col items-center gap-3">
      <div
        class="size-14 rounded-2xl flex items-center justify-center transition-transform"
        :class="[
          dragOver
            ? 'bg-(--ui-primary) text-white scale-110'
            : 'bg-(--ui-primary)/10 text-(--ui-primary) group-hover:scale-105',
        ]"
      >
        <UIcon name="i-lucide-file-spreadsheet" class="size-7" />
      </div>
      <div>
        <p class="text-base font-medium text-(--ui-text-highlighted)">
          {{ dragOver ? 'ここにドロップ' : 'クリックまたはドラッグ&ドロップで Excel をアップロード' }}
        </p>
        <p class="text-xs text-(--ui-text-muted) mt-1">
          対応: .xlsm / .xls / .xlsx ・ 上限 50MB 想定
        </p>
      </div>
    </div>
    <input
      ref="fileInput"
      type="file"
      :accept="accept"
      class="hidden"
      :disabled="disabled"
      @change="onChange"
    />
  </div>
</template>
