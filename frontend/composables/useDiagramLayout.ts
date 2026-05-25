/**
 * Backend の Diagram (DiagramNode/Edge) を Vue Flow の Node/Edge にマッピングする。
 *
 * - dagre で階層レイアウトを計算 (左→右 or 上→下) し、ノードに座標を割り当てる
 * - エッジの太さは weight で段階表示
 * - 種別 (sheet / module / procedure) ごとに class を切り替えてカラーリング
 */

import dagre from 'dagre'
import { MarkerType, Position, type Edge, type Node } from '@vue-flow/core'
import type { Diagram } from '~/types/api'

export interface LayoutOptions {
  /** dagre のランク方向. デフォルト 'LR' (左→右). 階層が浅いときは 'TB' (上→下) も検討. */
  direction?: 'LR' | 'TB'
  /** ノード幅 (px). カスタムノードのサイズに合わせる. */
  nodeWidth?: number
  /** ノード高 (px). */
  nodeHeight?: number
  /** ランク間距離 (px). */
  rankSep?: number
  /** 同一ランク内のノード間距離 (px). */
  nodeSep?: number
}

const DEFAULT_OPTS: Required<LayoutOptions> = {
  direction: 'LR',
  nodeWidth: 200,
  nodeHeight: 64,
  rankSep: 150,
  nodeSep: 70,
}

const EDGE_PALETTE = [
  '#2563eb',
  '#059669',
  '#dc2626',
  '#7c3aed',
  '#ca8a04',
  '#0891b2',
  '#db2777',
  '#4f46e5',
]

function hashString(value: string): number {
  let hash = 0
  for (let i = 0; i < value.length; i += 1) {
    hash = (hash * 31 + value.charCodeAt(i)) >>> 0
  }
  return hash
}

function edgeColor(source: string, index: number): string {
  return EDGE_PALETTE[(hashString(source) + index) % EDGE_PALETTE.length]
}

function edgeOffset(index: number): number {
  return 18 + (index % 5) * 8
}

/** Diagram を Vue Flow 用の {nodes, edges} に変換する. */
export function layoutDiagram(d: Diagram, opts: LayoutOptions = {}) {
  const o = { ...DEFAULT_OPTS, ...opts }
  const g = new dagre.graphlib.Graph()
  g.setGraph({
    rankdir: o.direction,
    ranksep: o.rankSep,
    nodesep: o.nodeSep,
    marginx: 20,
    marginy: 20,
  })
  g.setDefaultEdgeLabel(() => ({}))

  for (const n of d.nodes) {
    g.setNode(n.id, { width: o.nodeWidth, height: o.nodeHeight })
  }
  for (const e of d.edges) {
    g.setEdge(e.src, e.dst)
  }
  dagre.layout(g)

  const maxWeight = d.edges.reduce((m, e) => Math.max(m, e.weight), 1)

  const nodes: Node[] = d.nodes.map((n) => {
    const p = g.node(n.id)
    const isHorizontal = o.direction === 'LR'
    return {
      id: n.id,
      type: 'specNode',
      position: {
        x: (p?.x ?? 0) - o.nodeWidth / 2,
        y: (p?.y ?? 0) - o.nodeHeight / 2,
      },
      sourcePosition: isHorizontal ? Position.Right : Position.Bottom,
      targetPosition: isHorizontal ? Position.Left : Position.Top,
      data: {
        label: n.label,
        kind: n.kind,
        meta: n.meta,
      },
      // 上記 type で <SpecNode> を描画するため、Vue Flow デフォルトのスタイルは
      // 適用されないが、念のためサイズ指定だけしておく.
      style: { width: `${o.nodeWidth}px` },
      class: `node-${n.kind}`,
    }
  })

  const edges: Edge[] = d.edges.map((e, i) => {
    // 太さ: weight に応じて 1.4〜4.2 で段階
    const w = 1.4 + (e.weight / maxWeight) * 2.8
    const color = edgeColor(e.src, i)
    return {
      id: `e${i}-${e.src}->${e.dst}`,
      source: e.src,
      target: e.dst,
      // 重みが 2 以上のときだけラベル表示
      label: e.weight >= 2 ? `×${e.weight}` : undefined,
      labelShowBg: true,
      labelBgPadding: [5, 3],
      labelBgBorderRadius: 5,
      labelBgStyle: {
        fill: 'var(--ui-bg)',
        stroke: color,
        strokeWidth: 1,
      },
      labelStyle: {
        fill: color,
        fontSize: '10px',
        fontWeight: '700',
      },
      type: 'smoothstep',
      pathOptions: {
        offset: edgeOffset(i),
        borderRadius: 10,
      },
      animated: false,
      markerEnd: {
        type: MarkerType.ArrowClosed,
        color,
        width: 16,
        height: 16,
      },
      class: ['spec-diagram-edge', `edge-kind-${e.kind}`],
      interactionWidth: 18,
      zIndex: Math.round(w * 10),
      style: {
        strokeWidth: w,
        stroke: color,
        opacity: 0.78,
      },
    }
  })

  return { nodes, edges }
}
