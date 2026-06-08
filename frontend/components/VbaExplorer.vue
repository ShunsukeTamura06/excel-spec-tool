<script setup lang="ts">
/**
 * VBA タブ. 左にモジュール一覧 (procedure 含む), 右に選択中プロシージャの
 * コードと注釈を表示する.
 */

import type { VbaModule, VbaProcedure } from '~/types/api'

const props = defineProps<{ modules: VbaModule[] }>()

type Selection = { moduleName: string; procName: string } | null
const selected = ref<Selection>(null)

// 初期選択: 最初のモジュールの最初のプロシージャ
watchEffect(() => {
  if (!selected.value && props.modules.length > 0) {
    const m = props.modules[0]
    const proc = m?.procedures[0]
    if (m && proc) {
      selected.value = { moduleName: m.name, procName: proc.name }
    }
  }
})

const selectedProc = computed<VbaProcedure | null>(() => {
  if (!selected.value) return null
  const m = props.modules.find(x => x.name === selected.value!.moduleName)
  return m?.procedures.find(p => p.name === selected.value!.procName) ?? null
})
const selectedModule = computed<VbaModule | null>(() => {
  if (!selected.value) return null
  return props.modules.find(x => x.name === selected.value!.moduleName) ?? null
})

const kindIcon: Record<VbaProcedure['kind'], string> = {
  Sub: 'i-lucide-zap',
  Function: 'i-lucide-function-square',
  Property: 'i-lucide-key',
}
</script>

<template>
  <div v-if="props.modules.length === 0" class="rounded-2xl border border-dashed border-(--ui-border) p-10 text-center">
    <UIcon name="i-lucide-code-2" class="size-8 mx-auto text-(--ui-text-muted) mb-2" />
    <p class="text-sm text-(--ui-text-muted)">VBA モジュールがありません</p>
  </div>
  <div v-else class="grid grid-cols-1 lg:grid-cols-[280px_1fr] gap-4">
    <!-- 左: モジュール / プロシージャツリー -->
    <UCard :ui="{ body: 'p-2' }">
      <div class="space-y-3 max-h-[640px] overflow-y-auto">
        <div v-for="m in props.modules" :key="m.name">
          <div class="px-2 py-1 text-xs uppercase tracking-wide text-(--ui-text-muted) font-semibold flex items-center gap-1">
            <UIcon name="i-lucide-folder" class="size-3.5" />
            {{ m.name }}
            <span class="font-normal normal-case ml-auto">{{ m.type }}</span>
          </div>
          <ul class="space-y-0.5 mt-1">
            <li v-for="p in m.procedures" :key="p.name">
              <button
                type="button"
                class="w-full text-left px-2.5 py-1.5 rounded-md text-xs transition-colors flex items-center gap-2"
                :class="
                  selected?.moduleName === m.name && selected?.procName === p.name
                    ? 'bg-(--ui-primary)/10 text-(--ui-primary)'
                    : 'hover:bg-(--ui-bg-muted) text-(--ui-text)'
                "
                @click="selected = { moduleName: m.name, procName: p.name }"
              >
                <UIcon :name="kindIcon[p.kind]" class="size-3.5 shrink-0" />
                <span class="truncate flex-1">{{ p.name }}</span>
                <span class="text-[10px] text-(--ui-text-muted)">L{{ p.start_line }}</span>
              </button>
            </li>
            <li v-if="m.procedures.length === 0" class="px-2.5 py-1 text-[10px] text-(--ui-text-muted)">
              (プロシージャなし)
            </li>
          </ul>
        </div>
      </div>
    </UCard>

    <!-- 右: 選択中プロシージャ -->
    <UCard v-if="selectedProc && selectedModule">
      <template #header>
        <div class="flex items-baseline justify-between gap-2 flex-wrap">
          <div>
            <h3 class="text-lg font-semibold text-(--ui-text-highlighted) flex items-center gap-2">
              <UIcon :name="kindIcon[selectedProc.kind]" class="size-4 text-(--ui-primary)" />
              <span class="font-mono">{{ selectedModule.name }}.{{ selectedProc.name }}</span>
            </h3>
            <p class="text-xs text-(--ui-text-muted) mt-0.5">
              {{ selectedProc.kind }}
              ・ 行 {{ selectedProc.start_line }}–{{ selectedProc.end_line }}
              ({{ selectedProc.end_line - selectedProc.start_line + 1 }} 行)
            </p>
          </div>
          <UBadge
            v-if="selectedProc.annotation || selectedProc.side_effects.length || selectedProc.triggers.length || selectedProc.calls.length"
            color="primary"
            variant="subtle"
            icon="i-lucide-sparkles"
          >
            LLM 注釈
          </UBadge>
        </div>
        <p v-if="selectedProc.annotation" class="mt-2 text-sm">{{ selectedProc.annotation }}</p>

        <!-- 構造化注釈: 副作用 / 起動契機 / 呼出先 -->
        <div
          v-if="selectedProc.side_effects.length || selectedProc.triggers.length || selectedProc.calls.length"
          class="mt-3 grid grid-cols-1 md:grid-cols-3 gap-3 text-xs"
        >
          <div v-if="selectedProc.side_effects.length">
            <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1 flex items-center gap-1">
              <UIcon name="i-lucide-pencil" class="size-3" />副作用
            </p>
            <ul class="space-y-0.5">
              <li
                v-for="(s, i) in selectedProc.side_effects"
                :key="`se-${i}`"
                class="font-mono text-[11px] text-(--ui-text)"
              >
                {{ s }}
              </li>
            </ul>
          </div>
          <div v-if="selectedProc.triggers.length">
            <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1 flex items-center gap-1">
              <UIcon name="i-lucide-mouse-pointer-click" class="size-3" />起動契機
            </p>
            <ul class="space-y-0.5">
              <li v-for="(t, i) in selectedProc.triggers" :key="`tr-${i}`">{{ t }}</li>
            </ul>
          </div>
          <div v-if="selectedProc.calls.length">
            <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1 flex items-center gap-1">
              <UIcon name="i-lucide-arrow-right-circle" class="size-3" />呼出先
            </p>
            <ul class="space-y-0.5">
              <li
                v-for="(c, i) in selectedProc.calls"
                :key="`cl-${i}`"
                class="font-mono text-[11px]"
              >
                {{ c }}
              </li>
            </ul>
          </div>
        </div>
      </template>

      <pre class="text-xs font-mono leading-relaxed p-3 rounded-md bg-(--ui-bg-elevated) overflow-x-auto"><code>{{ selectedProc.code }}</code></pre>
    </UCard>
  </div>
</template>
