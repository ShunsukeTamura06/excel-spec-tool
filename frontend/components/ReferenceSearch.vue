<script setup lang="ts">
/** 参照逆引き. target を入力すると /references を叩いて結果テーブルを出す. */

import type { ReferenceItem } from '~/types/api'

const props = defineProps<{ jobId: string; initialTarget?: string }>()

const backend = useBackend()
const target = ref(props.initialTarget ?? '')
const loading = ref(false)
const errorMsg = ref('')
const results = ref<ReferenceItem[]>([])
const lastQuery = ref('')

async function search() {
  const q = target.value.trim()
  if (!q) return
  loading.value = true
  errorMsg.value = ''
  try {
    results.value = await backend.getReferences(props.jobId, q)
    lastQuery.value = q
  } catch (e) {
    errorMsg.value = e instanceof Error ? e.message : String(e)
  } finally {
    loading.value = false
  }
}

watch(
  () => props.initialTarget,
  (v) => {
    if (v) {
      target.value = v
      void search()
    }
  },
)

defineExpose({ runSearch: (t: string) => { target.value = t; return search() } })
</script>

<template>
  <UCard>
    <template #header>
      <div class="flex items-center gap-2">
        <UIcon name="i-lucide-search" class="size-4 text-(--ui-primary)" />
        <span class="font-medium">参照逆引き検索</span>
      </div>
      <p class="text-xs text-(--ui-text-muted) mt-1">
        セルや範囲 (例: <code class="px-1 rounded bg-(--ui-bg-elevated) text-[10px]">Calc!H2</code>,
        <code class="px-1 rounded bg-(--ui-bg-elevated) text-[10px]">Input!A:A</code>) を参照している箇所を検索します.
      </p>
    </template>

    <div class="flex gap-2">
      <UInput
        v-model="target"
        placeholder="例: Calc!H2"
        icon="i-lucide-target"
        class="flex-1"
        @keydown.enter="search"
      />
      <UButton color="primary" :loading="loading" icon="i-lucide-search" @click="search">検索</UButton>
    </div>

    <UAlert
      v-if="errorMsg"
      class="mt-3"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      :title="errorMsg"
    />

    <div v-if="lastQuery && !loading" class="mt-4">
      <div v-if="results.length === 0" class="text-sm text-(--ui-text-muted) text-center py-6">
        <UIcon name="i-lucide-inbox" class="size-6 mx-auto mb-1" />
        <code class="font-mono text-xs">{{ lastQuery }}</code> を参照している箇所はありません.
      </div>
      <div v-else>
        <p class="text-xs text-(--ui-text-muted) mb-2">
          <code class="font-mono">{{ lastQuery }}</code> を参照: <strong>{{ results.length }}</strong> 件
        </p>
        <div class="overflow-x-auto -mx-4">
          <table class="w-full text-xs">
            <thead class="bg-(--ui-bg-elevated)">
              <tr>
                <th class="px-3 py-2 text-left font-semibold w-20">種類</th>
                <th class="px-3 py-2 text-left font-semibold w-64">参照元</th>
                <th class="px-3 py-2 text-left font-semibold">コード</th>
              </tr>
            </thead>
            <tbody>
              <tr
                v-for="(r, i) in results"
                :key="i"
                class="border-t border-(--ui-border) hover:bg-(--ui-bg-muted)/40"
              >
                <td class="px-3 py-1.5">
                  <UBadge
                    :color="r.kind === 'vba' ? 'info' : 'neutral'"
                    variant="subtle"
                    size="sm"
                  >
                    {{ r.kind }}
                  </UBadge>
                </td>
                <td class="px-3 py-1.5 font-mono">{{ r.from }}</td>
                <td class="px-3 py-1.5 font-mono break-all">{{ r.code }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </UCard>
</template>
