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
      <MiniMap pannable zoomable position="bottom-right" />
    </VueFlow>
  </div>
</template>

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
</style>
