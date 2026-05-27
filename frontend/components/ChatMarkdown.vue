<script lang="ts">
import { defineComponent, h, ref, type PropType, type VNode } from 'vue'
import type { ToolTraceItem } from '~/types/api'

type Block =
  | { kind: 'paragraph'; text: string }
  | { kind: 'heading'; level: 2 | 3; text: string }
  | { kind: 'ul' | 'ol'; items: string[] }
  | { kind: 'code'; code: string }

type EvidenceHint = {
  title: string
  source: string
  lines: string[]
}

type ActiveTooltip = {
  x: number
  y: number
  token: string
  hints: EvidenceHint[]
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function asRecord(value: unknown): Record<string, unknown> {
  return isRecord(value) ? value : {}
}

function valueText(value: unknown): string {
  if (value == null || value === '') return ''
  if (Array.isArray(value)) return value.map(valueText).filter(Boolean).join(', ')
  return String(value)
}

function keyFor(value: unknown): string {
  return valueText(value).trim().toLowerCase()
}

function addHint(
  index: Map<string, EvidenceHint[]>,
  token: unknown,
  hint: EvidenceHint,
) {
  const key = keyFor(token)
  if (!key) return
  const current = index.get(key) ?? []
  if (current.length >= 3) return
  index.set(key, [...current, hint])
}

function buildEvidenceIndex(items: ToolTraceItem[]): Map<string, EvidenceHint[]> {
  const index = new Map<string, EvidenceHint[]>()

  for (const item of items) {
    const result = asRecord(item.result)

    if (item.name === 'find_cells') {
      const matches = Array.isArray(result.matches) ? result.matches : []
      for (const raw of matches) {
        const match = asRecord(raw)
        const sheet = valueText(match.sheet)
        const coord = valueText(match.coord)
        const value = valueText(match.value)
        const location = sheet && coord ? `${sheet}!${coord}` : sheet || coord
        const hint: EvidenceHint = {
          title: location || 'セル検索結果',
          source: 'セル検索',
          lines: [value && `値: ${value}`, location && `位置: ${location}`].filter(Boolean),
        }
        addHint(index, sheet, { ...hint, title: `${sheet} シート` })
        addHint(index, coord, hint)
        addHint(index, location, hint)
        addHint(index, value, hint)
      }
    }

    if (item.name === 'get_cells_range') {
      const sheet = valueText(result.sheet)
      const range = valueText(result.range)
      const rows = Array.isArray(result.rows) ? result.rows : []
      const preview = rows
        .slice(0, 3)
        .map((row) => Array.isArray(row)
          ? row.map(cell => valueText(asRecord(cell).formula || asRecord(cell).value)).join(' | ')
          : '')
        .filter(Boolean)
      const title = sheet && range ? `${sheet}!${range}` : sheet || range || 'セル範囲'
      const hint: EvidenceHint = {
        title,
        source: 'セル範囲',
        lines: [`範囲: ${title}`, ...preview],
      }
      addHint(index, sheet, { ...hint, title: `${sheet} シート` })
      addHint(index, range, hint)
      addHint(index, title, hint)
    }

    if (item.name === 'lookup_references') {
      const refs = Array.isArray(result.refs) ? result.refs : []
      for (const raw of refs) {
        const ref = asRecord(raw)
        const from = valueText(ref.from)
        const to = valueText(ref.to)
        const code = valueText(ref.code)
        const hint: EvidenceHint = {
          title: from || to || '参照関係',
          source: '参照関係',
          lines: [from && `参照元: ${from}`, to && `参照先: ${to}`, code].filter(Boolean),
        }
        addHint(index, from, hint)
        addHint(index, to, hint)
      }
    }

    if (item.name === 'get_vba_procedure') {
      const moduleName = valueText(result.module)
      const name = valueText(result.name)
      const start = valueText(result.start_line)
      const end = valueText(result.end_line)
      const annotation = valueText(result.annotation)
      const code = valueText(result.code).split('\n').slice(0, 4).join('\n')
      const title = moduleName && name ? `${moduleName}.${name}` : moduleName || name || 'VBA'
      const hint: EvidenceHint = {
        title,
        source: 'VBAコード',
        lines: [
          start && end && `行: ${start}-${end}`,
          annotation,
          code,
        ].filter(Boolean),
      }
      addHint(index, moduleName, hint)
      addHint(index, name, hint)
      addHint(index, title, hint)
    }

    if (item.name === 'list_sheet_formulas') {
      const sheet = valueText(result.sheet)
      const formulas = Array.isArray(result.formulas) ? result.formulas : []
      addHint(index, sheet, {
        title: `${sheet} シート`,
        source: '数式一覧',
        lines: [`数式: ${valueText(result.returned)} / ${valueText(result.total)} 件`],
      })
      for (const raw of formulas.slice(0, 30)) {
        const formula = asRecord(raw)
        const coord = valueText(formula.coord)
        const body = valueText(formula.formula)
        const location = sheet && coord ? `${sheet}!${coord}` : coord
        addHint(index, coord, {
          title: location || '数式',
          source: '数式一覧',
          lines: [body].filter(Boolean),
        })
      }
    }

    if (item.name === 'list_workbook_objects') {
      for (const group of ['charts', 'pivot_tables', 'power_queries']) {
        const objects = Array.isArray(result[group]) ? result[group] : []
        for (const raw of objects) {
          const object = asRecord(raw)
          const name = valueText(object.name || object.title)
          const sheet = valueText(object.sheet || object.target_sheet)
          const source = valueText(object.source_ref || object.source_name || object.command)
          const hint: EvidenceHint = {
            title: name || sheet || 'Excelオブジェクト',
            source: 'Excelオブジェクト',
            lines: [sheet && `シート: ${sheet}`, source && `元データ: ${source}`].filter(Boolean),
          }
          addHint(index, name, hint)
          addHint(index, sheet, hint)
        }
      }
    }
  }

  return index
}

function parseBlocks(markdown: string): Block[] {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n')
  const blocks: Block[] = []
  let i = 0

  while (i < lines.length) {
    const line = lines[i] ?? ''
    if (!line.trim()) {
      i += 1
      continue
    }

    if (line.startsWith('```')) {
      const code: string[] = []
      i += 1
      while (i < lines.length && !(lines[i] ?? '').startsWith('```')) {
        code.push(lines[i] ?? '')
        i += 1
      }
      if (i < lines.length) i += 1
      blocks.push({ kind: 'code', code: code.join('\n') })
      continue
    }

    const heading = line.match(/^(#{2,3})\s+(.+)$/)
    if (heading) {
      blocks.push({
        kind: 'heading',
        level: heading[1].length as 2 | 3,
        text: heading[2],
      })
      i += 1
      continue
    }

    const unorderedItems: string[] = []
    while (i < lines.length) {
      const match = (lines[i] ?? '').match(/^\s*[-*]\s+(.+)$/)
      if (!match) break
      unorderedItems.push(match[1])
      i += 1
    }
    if (unorderedItems.length) {
      blocks.push({ kind: 'ul', items: unorderedItems })
      continue
    }

    const orderedItems: string[] = []
    while (i < lines.length) {
      const match = (lines[i] ?? '').match(/^\s*\d+\.\s+(.+)$/)
      if (!match) break
      orderedItems.push(match[1])
      i += 1
    }
    if (orderedItems.length) {
      blocks.push({ kind: 'ol', items: orderedItems })
      continue
    }

    const paragraph: string[] = []
    while (i < lines.length && (lines[i] ?? '').trim()) {
      const current = lines[i] ?? ''
      if (
        current.startsWith('```')
        || /^(#{2,3})\s+/.test(current)
        || /^\s*[-*]\s+/.test(current)
        || /^\s*\d+\.\s+/.test(current)
      ) {
        break
      }
      paragraph.push(current)
      i += 1
    }
    blocks.push({ kind: 'paragraph', text: paragraph.join('\n') })
  }

  return blocks
}

function safeHref(value: string): string | null {
  if (/^https?:\/\//i.test(value)) return value
  if (value.startsWith('/')) return value
  return null
}

function fallbackHint(token: string): EvidenceHint {
  return {
    title: token,
    source: '回答内の識別子',
    lines: [
      'この回答に紐づく根拠カード内では、完全一致する実データを特定できませんでした。',
      '必要に応じて下の根拠カードや設計書で確認してください。',
    ],
  }
}

function uniqueHints(hints: EvidenceHint[]): EvidenceHint[] {
  const seen = new Set<string>()
  return hints.filter((hint) => {
    const key = `${hint.source}:${hint.title}:${hint.lines.join('\n')}`
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}

function resolveEvidenceHints(
  token: string,
  evidenceIndex: Map<string, EvidenceHint[]>,
): EvidenceHint[] {
  const exact = evidenceIndex.get(keyFor(token))
  if (exact?.length) return exact

  const hints: EvidenceHint[] = []
  const parts = token.split('!')
  if (parts.length === 2) {
    hints.push(...(evidenceIndex.get(keyFor(parts[0])) ?? []))
    hints.push(...(evidenceIndex.get(keyFor(parts[1])) ?? []))
  }

  for (const [key, value] of evidenceIndex.entries()) {
    if (key.endsWith(`!${keyFor(token)}`) || key.includes(keyFor(token))) {
      hints.push(...value)
    }
  }

  const resolved = uniqueHints(hints).slice(0, 3)
  return resolved.length ? resolved : [fallbackHint(token)]
}

function tooltipTitle(token: string, hints: EvidenceHint[]): string {
  return [token, ...hints.flatMap(hint => [hint.title, ...hint.lines])].join('\n')
}

function renderEvidenceToken(
  token: string,
  hints: EvidenceHint[],
  showTooltip: (event: MouseEvent | FocusEvent, token: string, hints: EvidenceHint[]) => void,
  moveTooltip: (event: MouseEvent) => void,
  hideTooltip: () => void,
): VNode {
  return h('span', { class: 'chat-evidence-token' }, [
    h('code', {
      tabindex: '0',
      title: tooltipTitle(token, hints),
      onMouseenter: (event: MouseEvent) => showTooltip(event, token, hints),
      onMousemove: (event: MouseEvent) => moveTooltip(event),
      onMouseleave: hideTooltip,
      onFocus: (event: FocusEvent) => showTooltip(event, token, hints),
      onBlur: hideTooltip,
    }, token),
  ])
}

function renderInline(
  text: string,
  evidenceIndex: Map<string, EvidenceHint[]>,
  showTooltip: (event: MouseEvent | FocusEvent, token: string, hints: EvidenceHint[]) => void,
  moveTooltip: (event: MouseEvent) => void,
  hideTooltip: () => void,
): VNode[] {
  const nodes: VNode[] = []
  let i = 0

  const pushText = (value: string) => {
    if (!value) return
    const parts = value.split('\n')
    parts.forEach((part, index) => {
      if (index > 0) nodes.push(h('br'))
      if (part) nodes.push(h('span', part))
    })
  }

  while (i < text.length) {
    if (text.startsWith('**', i)) {
      const end = text.indexOf('**', i + 2)
      if (end > i + 2) {
        nodes.push(h('strong', renderInline(
          text.slice(i + 2, end),
          evidenceIndex,
          showTooltip,
          moveTooltip,
          hideTooltip,
        )))
        i = end + 2
        continue
      }
    }

    if (text[i] === '`') {
      const end = text.indexOf('`', i + 1)
      if (end > i + 1) {
        const token = text.slice(i + 1, end)
        nodes.push(renderEvidenceToken(
          token,
          resolveEvidenceHints(token, evidenceIndex),
          showTooltip,
          moveTooltip,
          hideTooltip,
        ))
        i = end + 1
        continue
      }
    }

    if (text[i] === '[') {
      const labelEnd = text.indexOf(']', i + 1)
      const hrefStart = labelEnd >= 0 ? text.indexOf('(', labelEnd) : -1
      const hrefEnd = hrefStart >= 0 ? text.indexOf(')', hrefStart) : -1
      if (labelEnd > i && hrefStart === labelEnd + 1 && hrefEnd > hrefStart + 1) {
        const href = safeHref(text.slice(hrefStart + 1, hrefEnd))
        if (href) {
          nodes.push(h('a', { href, target: href.startsWith('http') ? '_blank' : undefined, rel: 'noreferrer' }, text.slice(i + 1, labelEnd)))
          i = hrefEnd + 1
          continue
        }
      }
    }

    const nextSpecials = ['**', '`', '[']
      .map(marker => text.indexOf(marker, i + 1))
      .filter(pos => pos >= 0)
    const next = nextSpecials.length ? Math.min(...nextSpecials) : text.length
    pushText(text.slice(i, next))
    i = next
  }

  return nodes
}

export default defineComponent({
  name: 'ChatMarkdown',
  props: {
    markdown: {
      type: String,
      required: true,
    },
    evidenceItems: {
      type: Array as PropType<ToolTraceItem[]>,
      default: () => [],
    },
  },
  setup(props) {
    const activeTooltip = ref<ActiveTooltip | null>(null)

    function tooltipPosition(event: MouseEvent | FocusEvent): { x: number; y: number } {
      if ('clientX' in event && event.clientX > 0) {
        return { x: event.clientX + 12, y: event.clientY + 16 }
      }
      const target = event.target instanceof HTMLElement ? event.target : null
      const rect = target?.getBoundingClientRect()
      return {
        x: (rect?.left ?? 0) + 12,
        y: (rect?.bottom ?? 0) + 8,
      }
    }

    function showTooltip(
      event: MouseEvent | FocusEvent,
      token: string,
      hints: EvidenceHint[],
    ) {
      activeTooltip.value = { ...tooltipPosition(event), token, hints }
    }

    function moveTooltip(event: MouseEvent) {
      if (!activeTooltip.value) return
      activeTooltip.value = {
        ...activeTooltip.value,
        ...tooltipPosition(event),
      }
    }

    function hideTooltip() {
      activeTooltip.value = null
    }

    return () => {
      const evidenceIndex = buildEvidenceIndex(props.evidenceItems)
      return h('div', { class: 'chat-markdown spec-prose text-sm' }, [
        ...parseBlocks(props.markdown).map((block, index) => {
          switch (block.kind) {
            case 'heading':
              return h(`h${block.level}`, { key: index }, renderInline(
                block.text,
                evidenceIndex,
                showTooltip,
                moveTooltip,
                hideTooltip,
              ))
            case 'ul':
            case 'ol':
              return h(block.kind, { key: index }, block.items.map((item, itemIndex) => h(
                'li',
                { key: itemIndex },
                renderInline(item, evidenceIndex, showTooltip, moveTooltip, hideTooltip),
              )))
            case 'code':
              return h('pre', { key: index }, [h('code', block.code)])
            case 'paragraph':
              return h('p', { key: index }, renderInline(
                block.text,
                evidenceIndex,
                showTooltip,
                moveTooltip,
                hideTooltip,
              ))
          }
        }),
        activeTooltip.value && h('div', {
          class: 'chat-evidence-popover',
          role: 'tooltip',
          style: {
            left: `${Math.min(activeTooltip.value.x, window.innerWidth - 340)}px`,
            top: `${Math.min(activeTooltip.value.y, window.innerHeight - 260)}px`,
          },
        }, [
          h('div', { class: 'chat-evidence-popover-token' }, activeTooltip.value.token),
          ...activeTooltip.value.hints.map((hint, index) => h(
            'div',
            { class: 'chat-evidence-popover-item', key: `${hint.title}-${index}` },
            [
              h('div', { class: 'chat-evidence-popover-title' }, hint.title),
              h('div', { class: 'chat-evidence-popover-source' }, hint.source),
              ...hint.lines.slice(0, 4).map((line, lineIndex) => h(
                'div',
                { class: 'chat-evidence-popover-line', key: `${line}-${lineIndex}` },
                line,
              )),
            ],
          )),
        ]),
      ])
    }
  },
})
</script>

<style scoped>
.chat-evidence-token {
  display: inline;
}

.chat-evidence-token code {
  cursor: help;
}

.chat-evidence-popover {
  background: var(--ui-bg);
  border: 1px solid var(--ui-border);
  border-radius: 0.5rem;
  box-shadow: 0 12px 32px rgb(0 0 0 / 16%);
  color: var(--ui-text);
  font-size: 0.75rem;
  line-height: 1.4;
  max-height: min(18rem, 50vh);
  overflow-y: auto;
  padding: 0.625rem;
  position: fixed;
  width: min(24rem, calc(100vw - 1rem));
  z-index: 1000;
}

.chat-evidence-popover-token {
  color: var(--ui-primary);
  font-family: var(--font-mono);
  font-size: 0.75rem;
  font-weight: 600;
  margin-bottom: 0.375rem;
}

.chat-evidence-popover-item + .chat-evidence-popover-item {
  border-top: 1px solid var(--ui-border);
  margin-top: 0.5rem;
  padding-top: 0.5rem;
}

.chat-evidence-popover-title {
  color: var(--ui-text-highlighted);
  font-size: 0.75rem;
  font-weight: 600;
}

.chat-evidence-popover-source {
  color: var(--ui-primary);
  font-size: 0.6875rem;
  margin-bottom: 0.25rem;
}

.chat-evidence-popover-line {
  overflow-wrap: anywhere;
  white-space: pre-wrap;
}
</style>
