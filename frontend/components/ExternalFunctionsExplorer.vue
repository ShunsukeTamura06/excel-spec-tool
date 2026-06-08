<script setup lang="ts">
/**
 * 「外部関数」タブ. Bloomberg BDH/BDP/BDS 等の使用箇所と定義を表示する.
 *
 * - 左: このブックで使用されている関数のリスト (件数 + ベンダー)
 * - 右: 選択中関数の定義 (シグネチャ / 引数 / 例 / 注意点) と使用箇所
 * - 「サポート対象一覧」リンクで未使用のものも参照可能
 */

import type {
  ExternalFunction,
  ExternalFunctionRegistry,
  ExternalFunctionUsageItem,
  ExternalFunctionsUsed,
} from '~/types/api'

const props = defineProps<{ jobId: string }>()

const backend = useBackend()

const registry = ref<ExternalFunctionRegistry | null>(null)
const usage = ref<ExternalFunctionsUsed | null>(null)
const loading = ref(false)
const errorMsg = ref('')

const selectedName = ref<string>('')
const showAllRegistry = ref(false)

async function load() {
  loading.value = true
  errorMsg.value = ''
  try {
    const [reg, used] = await Promise.all([
      backend.getExternalFunctions(),
      backend.getExternalFunctionsUsed(props.jobId),
    ])
    registry.value = reg
    usage.value = used
    if (used.items.length > 0) {
      selectedName.value = used.items[0]!.name
    } else if (reg.functions.length > 0) {
      selectedName.value = reg.functions[0]!.name
    }
  } catch (e) {
    errorMsg.value = friendlyMessage(e)
  } finally {
    loading.value = false
  }
}
await load()

// 表示候補: 使用関数 (上) + 未使用だが登録済みの関数 (下、showAllRegistry オン時のみ)
type ListEntry =
  | { kind: 'used'; usage: ExternalFunctionUsageItem }
  | { kind: 'available'; fn: ExternalFunction }

const listEntries = computed<ListEntry[]>(() => {
  const out: ListEntry[] = []
  const usedNames = new Set<string>()
  for (const u of usage.value?.items ?? []) {
    out.push({ kind: 'used', usage: u })
    usedNames.add(u.name)
  }
  if (showAllRegistry.value && registry.value) {
    for (const fn of registry.value.functions) {
      if (!usedNames.has(fn.name)) {
        out.push({ kind: 'available', fn })
      }
    }
  }
  return out
})

const selectedFunction = computed<ExternalFunction | null>(() => {
  if (!registry.value || !selectedName.value) return null
  return registry.value.functions.find(f => f.name === selectedName.value) ?? null
})

const selectedUsage = computed<ExternalFunctionUsageItem | null>(() => {
  if (!usage.value || !selectedName.value) return null
  return usage.value.items.find(u => u.name === selectedName.value) ?? null
})
</script>

<template>
  <div class="space-y-3">
    <!-- サマリーバー -->
    <div v-if="usage" class="flex items-center justify-between gap-3 flex-wrap">
      <div class="flex items-center gap-3 text-sm">
        <UIcon name="i-lucide-puzzle" class="size-4 text-(--ui-primary)" />
        <span class="font-medium">外部 Add-In 関数</span>
        <UBadge color="primary" variant="subtle" size="sm" class="tabular-nums">
          {{ usage.total_kinds }} 種類
        </UBadge>
        <UBadge color="neutral" variant="subtle" size="sm" class="tabular-nums">
          {{ usage.total_uses }} 箇所
        </UBadge>
      </div>
      <div class="flex items-center gap-2">
        <USwitch v-model="showAllRegistry" size="sm" />
        <span class="text-xs text-(--ui-text-muted)">未使用も含めて全関数を表示</span>
      </div>
    </div>

    <UAlert
      v-if="errorMsg"
      color="error"
      variant="subtle"
      icon="i-lucide-alert-triangle"
      title="外部関数情報の取得に失敗"
      :description="errorMsg"
    />

    <div
      v-if="usage && usage.items.length === 0 && !showAllRegistry"
      class="rounded-2xl border border-dashed border-(--ui-border) p-10 text-center"
    >
      <div class="size-12 mx-auto rounded-2xl bg-(--ui-bg-muted) flex items-center justify-center text-(--ui-text-muted) mb-3">
        <UIcon name="i-lucide-puzzle" class="size-6" />
      </div>
      <p class="text-sm text-(--ui-text-muted) mb-2">
        このブックでは Bloomberg 等の外部 Add-In 関数は検出されませんでした。
      </p>
      <p class="text-xs text-(--ui-text-muted)">
        サポート対象関数の定義を見たい場合は、右上のトグルを ON にしてください。
      </p>
    </div>

    <div v-else-if="listEntries.length > 0" class="grid grid-cols-1 lg:grid-cols-[280px_1fr] gap-4">
      <!-- 左: 関数リスト -->
      <UCard :ui="{ body: 'p-2' }">
        <ul class="space-y-0.5 max-h-[640px] overflow-y-auto">
          <li v-for="entry in listEntries" :key="`${entry.kind}-${entry.kind === 'used' ? entry.usage.name : entry.fn.name}`">
            <button
              type="button"
              class="w-full text-left px-3 py-2 rounded-md transition-colors flex items-center justify-between gap-2"
              :class="
                (entry.kind === 'used' ? entry.usage.name : entry.fn.name) === selectedName
                  ? 'bg-(--ui-primary)/10 text-(--ui-primary)'
                  : 'hover:bg-(--ui-bg-muted) text-(--ui-text)'
              "
              @click="selectedName = entry.kind === 'used' ? entry.usage.name : entry.fn.name"
            >
              <span class="flex flex-col">
                <span class="text-sm font-mono font-semibold">
                  {{ entry.kind === 'used' ? entry.usage.name : entry.fn.name }}
                </span>
                <span class="text-[10px] text-(--ui-text-muted)">
                  {{ entry.kind === 'used' ? entry.usage.vendor : entry.fn.vendor }}
                </span>
              </span>
              <UBadge
                v-if="entry.kind === 'used'"
                color="primary"
                variant="subtle"
                size="sm"
                class="tabular-nums"
              >
                {{ entry.usage.count }}
              </UBadge>
              <UBadge v-else color="neutral" variant="subtle" size="sm" class="text-[10px]">
                未使用
              </UBadge>
            </button>
          </li>
        </ul>
      </UCard>

      <!-- 右: 詳細 -->
      <div v-if="selectedFunction" class="space-y-3">
        <UCard>
          <template #header>
            <div class="flex items-baseline justify-between gap-2 flex-wrap">
              <div>
                <h3 class="text-lg font-semibold text-(--ui-text-highlighted) flex items-center gap-2">
                  <code class="font-mono">{{ selectedFunction.name }}</code>
                  <UBadge color="info" variant="subtle" size="sm">{{ selectedFunction.vendor }}</UBadge>
                </h3>
                <p class="text-xs text-(--ui-text-muted) mt-0.5">{{ selectedFunction.short }}</p>
              </div>
              <div v-if="selectedUsage" class="flex items-center gap-1.5">
                <UBadge color="primary" variant="solid" size="sm" class="tabular-nums">
                  {{ selectedUsage.count }} 箇所で使用
                </UBadge>
              </div>
            </div>
          </template>

          <div class="space-y-3">
            <!-- シグネチャ -->
            <div>
              <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1">
                シグネチャ
              </p>
              <code class="block font-mono text-xs p-2 rounded bg-(--ui-bg-elevated) break-all">{{
                selectedFunction.signature
              }}</code>
            </div>

            <!-- 長文説明 -->
            <p class="text-sm whitespace-pre-line leading-relaxed">{{ selectedFunction.long }}</p>

            <!-- 引数 -->
            <div v-if="selectedFunction.params.length">
              <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1">
                引数
              </p>
              <ul class="space-y-1.5 text-xs">
                <li
                  v-for="(p, i) in selectedFunction.params"
                  :key="`p-${i}`"
                  class="flex items-start gap-2"
                >
                  <span class="font-mono font-semibold text-(--ui-text-highlighted) shrink-0">
                    {{ p.name }}
                  </span>
                  <UBadge
                    :color="p.required ? 'primary' : 'neutral'"
                    variant="subtle"
                    size="sm"
                    class="shrink-0"
                  >
                    {{ p.required ? '必須' : '任意' }}
                  </UBadge>
                  <span v-if="p.type" class="text-(--ui-text-muted) shrink-0">({{ p.type }})</span>
                  <span class="text-(--ui-text)">{{ p.description }}</span>
                </li>
              </ul>
            </div>

            <!-- 返り値 -->
            <div v-if="selectedFunction.returns">
              <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1">
                返り値
              </p>
              <p class="text-sm">{{ selectedFunction.returns }}</p>
            </div>

            <!-- 使用例 -->
            <div v-if="selectedFunction.examples.length">
              <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1">
                使用例
              </p>
              <div class="space-y-1">
                <code
                  v-for="(ex, i) in selectedFunction.examples"
                  :key="`ex-${i}`"
                  class="block font-mono text-xs p-2 rounded bg-(--ui-bg-elevated) break-all"
                >{{ ex }}</code>
              </div>
            </div>

            <!-- 注意点 -->
            <div v-if="selectedFunction.notes.length">
              <p class="text-[10px] uppercase tracking-wide text-(--ui-text-muted) font-semibold mb-1 flex items-center gap-1">
                <UIcon name="i-lucide-alert-triangle" class="size-3 text-amber-600" />
                注意点
              </p>
              <ul class="space-y-1 text-xs">
                <li v-for="(n, i) in selectedFunction.notes" :key="`n-${i}`" class="flex items-start gap-1.5">
                  <UIcon name="i-lucide-dot" class="size-3.5 shrink-0 mt-0.5" />
                  <span>{{ n }}</span>
                </li>
              </ul>
            </div>

            <!-- 公式参考 -->
            <p v-if="selectedFunction.doc_url" class="text-xs text-(--ui-text-muted)">
              <UIcon name="i-lucide-external-link" class="size-3 inline align-text-bottom" />
              公式参考: <code class="font-mono">{{ selectedFunction.doc_url }}</code>
            </p>
          </div>
        </UCard>

        <!-- 使用箇所 (使用関数のみ表示) -->
        <UCard v-if="selectedUsage && selectedUsage.locations.length">
          <template #header>
            <div class="flex items-center gap-2">
              <UIcon name="i-lucide-target" class="size-4 text-(--ui-primary)" />
              <span class="font-medium">使用箇所 ({{ selectedUsage.locations.length }} 件)</span>
            </div>
          </template>
          <div class="overflow-x-auto -mx-4 max-h-[400px]">
            <table class="w-full text-xs">
              <thead class="bg-(--ui-bg-elevated) sticky top-0">
                <tr>
                  <th class="px-3 py-2 text-left font-semibold w-32">シート</th>
                  <th class="px-3 py-2 text-left font-semibold w-40">セル</th>
                  <th class="px-3 py-2 text-left font-semibold">数式</th>
                </tr>
              </thead>
              <tbody>
                <tr
                  v-for="(loc, i) in selectedUsage.locations"
                  :key="`loc-${i}`"
                  class="border-t border-(--ui-border) hover:bg-(--ui-bg-muted)/40"
                >
                  <td class="px-3 py-1.5 font-mono">{{ loc.sheet }}</td>
                  <td class="px-3 py-1.5 font-mono">{{ loc.coord }}</td>
                  <td class="px-3 py-1.5 font-mono break-all">{{ loc.formula }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </UCard>
      </div>
    </div>
  </div>
</template>
