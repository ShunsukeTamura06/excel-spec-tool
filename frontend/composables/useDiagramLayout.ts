/**
 * Backend の Diagram (DiagramNode/Edge) を Vue Flow の Node/Edge にマッピングする。
 *
 * - dagre で階層レイアウトを計算 (左→右 or 上→下) し、ノードに座標を割り当てる
 * - エッジの太さは weight で段階表示
 * - 種別 (sheet / module / procedure) ごとに class を切り替えてカラーリング
 */

import dagre from 'dagre'
import { MarkerType, type Edge, type Node } from '@vue-flow/core'
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
  rankSep: 90,
  nodeSep: 30,
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
    return {
      id: n.id,
      type: 'specNode',
      position: {
        x: (p?.x ?? 0) - o.nodeWidth / 2,
        y: (p?.y ?? 0) - o.nodeHeight / 2,
      },
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
    // 太さ: weight に応じて 1.2〜4 で段階
    const w = 1.2 + (e.weight / maxWeight) * 2.8
    return {
      id: `e${i}-${e.src}->${e.dst}`,
      source: e.src,
      target: e.dst,
      // 重みが 2 以上のときだけラベル表示
      label: e.weight >= 2 ? `×${e.weight}` : undefined,
      labelBgPadding: [4, 2],
      labelBgBorderRadius: 4,
      labelStyle: { fontSize: '10px', fontWeight: '600' },
      type: 'smoothstep',
      animated: e.kind === 'call',
      markerEnd: MarkerType.ArrowClosed,
      style: {
        strokeWidth: w,
        stroke: e.kind === 'call' ? 'var(--ui-color-info-500)' : 'var(--ui-color-primary-500)',
      },
    }
  })

  return { nodes, edges }
}
