<script setup lang="ts">
/**
 * 単一のグラフを Vue Flow で描画する.
 * 親から `diagram` を受け取り、dagre レイアウトを適用して表示する.
 */

import { VueFlow, useVueFlow } from '@vue-flow/core'
import { Background } from '@vue-flow/background'
import { Controls } from '@vue-flow/controls'
import { MiniMap } from '@vue-flow/minimap'
import type { Diagram } from '~/types/api'
import SpecNode from './SpecNode.vue'

const props = withDefaults(
  defineProps<{ diagram: Diagram; direction?: 'LR' | 'TB' }>(),
  { direction: 'LR' },
)

// 同じページ内に複数 VueFlow を出すと内部 ID 衝突するため id を分ける
const flowId = computed(() => `flow-${props.diagram.kind}`)

const { nodes, edges } = layoutDiagram(props.diagram, { direction: props.direction })
const nodeRefs = ref(nodes)
const edgeRefs = ref(edges)

// diagram が差し替わったら再レイアウト
watch(
  () => [props.diagram, props.direction],
  () => {
    const out = layoutDiagram(props.diagram, { direction: props.direction })
    nodeRefs.value = out.nodes
    edgeRefs.value = out.edges
  },
  { deep: true },
)

// VueFlow インスタンスは flowId 毎に分かれる
const { fitView } = useVueFlow(flowId.value)
onMounted(() => {
  nextTick(() => fitView({ padding: 0.15, duration: 0 }))
})

const isEmpty = computed(() => props.diagram.nodes.length === 0)
</script>

<template>
  <div
    class="relative w-full rounded-xl border border-(--ui-border) bg-(--ui-bg-elevated)/30 overflow-hidden"
    style="height: 640px"
  >
    <div
      v-if="isEmpty"
      class="absolute inset-0 flex flex-col items-center justify-center text-center"
    >
      <UIcon name="i-lucide-network-off" class="size-10 text-(--ui-text-muted) mb-2" />
      <p class="text-sm text-(--ui-text-muted)">表示できるノードがありません</p>
    </div>
    <VueFlow
      v-else
      :id="flowId"
      :nodes="nodeRefs"
      :edges="edgeRefs"
      :node-types="{ specNode: markRaw(SpecNode) }"
      :default-edge-options="{ type: 'smoothstep' }"
      :min-zoom="0.2"
      :max-zoom="2.5"
      fit-view-on-init
    >
      <Background pattern-color="var(--ui-border)" :gap="20" :size="1" />
      <Controls position="bottom-left" />
      <MiniMap
        pannable
        zoomable
        position="bottom-right"
        mask-color="color-mix(in srgb, var(--ui-bg) 65%, transparent)"
        :node-color="miniMapNodeColor"
        :node-stroke-color="miniMapNodeColor"
        :node-stroke-width="0"
      />
    </VueFlow>
  </div>
</template>

<script lang="ts">
// MiniMap のノード色を node-{kind} に揃える. テンプレ内で関数定義すると
// 毎レンダ新インスタンスになって MiniMap が無駄に再描画されるので、
// モジュールスコープに置く.
function miniMapNodeColor(node: { class?: string }): string {
  const cls = node.class ?? ''
  if (cls.includes('node-procedure')) return '#0ea5e9'
  if (cls.includes('node-module')) return '#6366f1'
  return '#10b981' // sheet
}
</script>

<style scoped>
:deep(.spec-diagram-edge) {
  transition:
    opacity 120ms ease,
    filter 120ms ease,
    stroke-width 120ms ease;
}

:deep(.spec-diagram-edge:hover),
:deep(.spec-diagram-edge.selected) {
  opacity: 1 !important;
  filter: drop-shadow(0 0 3px color-mix(in srgb, currentColor 35%, transparent));
}

:deep(.edge-kind-call) {
  stroke-dasharray: 8 5;
}

/* MiniMap がデフォルトで真っ白の背景を出してダークテーマに浮くので、
   コンテナ自体を bg-elevated に合わせる. */
:deep(.vue-flow__minimap) {
  background: color-mix(in srgb, var(--ui-bg-elevated) 92%, transparent);
  border: 1px solid var(--ui-border);
  border-radius: 8px;
}
</style>
