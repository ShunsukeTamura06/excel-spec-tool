/**
 * Backend (FastAPI) のレスポンス型. Pydantic モデルと対応する.
 *
 * backend/core/models.py の Pydantic 定義に合わせる. Pydantic 側を変えたら
 * こちらも合わせて変える (現状は手動同期. 将来は OpenAPI 生成に切替え可).
 */

export type JobStatus = 'uploaded' | 'extracted' | 'analyzed' | 'failed'

export interface JobMeta {
  job_id: string
  filename: string
  created_at: string
  status: JobStatus
  file_sha256?: string | null
  file_size?: number | null
}

export interface SpecResponse {
  spec_md: string
  meta: JobMeta
}

export interface ReferenceItem {
  kind: 'formula' | 'vba' | 'chart' | 'pivot' | 'power_query'
  /** Pydantic 側で `from_` を alias `from` で吐く. JSON では `from`. */
  from: string
  to: string
  code: string
}

export interface ChatMessage {
  role: 'user' | 'assistant' | 'system'
  content: string
  timestamp: string
}

export interface ChatSessionMeta {
  session_id: string
  title: string
  created_at: string
  updated_at: string
  archived: boolean
  last_message_preview: string
  message_count: number
}

// ---------- feedback ----------

export type FeedbackKind =
  | 'thumbs_up'
  | 'thumbs_down'
  | 'improvement'
  | 'bug'
  | 'other'

export interface FeedbackInput {
  kind: FeedbackKind
  comment?: string
  page?: string
  job_id?: string | null
  target_id?: string | null
  target_excerpt?: string
  user_label?: string
}

export interface FeedbackResponse {
  ok: boolean
  id: string
}

export interface ToolTraceItem {
  name: string
  arguments: Record<string, unknown>
  result_preview: string
  result?: unknown
}

export interface ChatReply {
  reply: string
  history: ChatMessage[]
  tool_trace?: ToolTraceItem[]
}

export type ChatStreamEventName = 'status' | 'tool_start' | 'tool_result' | 'final' | 'error'

export interface ChatProgressEvent {
  message: string
  tool_name?: string
  iteration?: number
}

export interface ChatSessionResponse {
  session: ChatSessionMeta
}

export interface ExtractResponse {
  job_id: string
  duplicate?: boolean
  duplicate_of?: string | null
}

export interface AnalyzeResponse {
  status: string
}

export interface DeleteResponse {
  deleted: boolean
}

// ---------- workbook structure ----------

export interface CellFormula {
  coord: string
  formula: string
  refs: string[]
  annotation: string
  external_functions: string[]
}

export interface NamedRange {
  name: string
  refers_to: string
}

export interface ConditionalFormat {
  range: string
  rule: string
}

export interface ExcelTable {
  name: string
  ref: string
  header_row_count: number
}

export interface ChartSeriesItem {
  name: string
  values_ref: string
  categories_ref: string
}

export interface ChartObjectItem {
  name: string
  chart_type: string
  title: string
  anchor: string
  series: ChartSeriesItem[]
}

export interface PivotTableItem {
  name: string
  anchor: string
  cache_id: string
  source_type: string
  source_sheet: string
  source_ref: string
  source_name: string
  row_fields: string[]
  column_fields: string[]
  value_fields: string[]
  filter_fields: string[]
}

export interface PowerQueryItem {
  name: string
  kind: 'power_query' | 'connection'
  connection_id: string
  connection_type: string
  description: string
  refresh_on_load: boolean
  target_sheet: string
  target_name: string
  source: string
  command: string
  m_code: string
  confidence: 'explicit' | 'inferred' | 'unknown'
}

export interface AnalysisRiskItem {
  category:
    | 'dynamic_vba'
    | 'runtime_state'
    | 'dynamic_formula'
    | 'external_dependency'
    | 'event_macro'
    | 'unknown_object_dependency'
  severity: 'high' | 'medium' | 'low'
  location: string
  evidence: string
  description: string
  recommendation: string
  confidence: 'explicit' | 'inferred' | 'unknown'
}

export interface DataValidationItem {
  range: string
  type: string
  formula: string
  operator: string
  prompt: string
  error: string
  allow_blank: boolean
}

export interface FormControlItem {
  kind: string
  name: string
  text: string
  macro: string
  anchor: string
}

export interface SheetInfo {
  name: string
  rows: number
  cols: number
  formulas: CellFormula[]
  named_ranges: NamedRange[]
  conditional_formats: ConditionalFormat[]
  tables: ExcelTable[]
  charts: ChartObjectItem[]
  pivot_tables: PivotTableItem[]
  merged_ranges: string[]
  data_validations: DataValidationItem[]
  form_controls: FormControlItem[]
  preview_rows: (string | null)[][]
  preview_origin: string
  // 構造化 LLM 注釈 (P1-2). いずれも空可.
  purpose: string
  inputs: string[]
  outputs: string[]
  main_calculations: string[]
  usage_scenario: string
}

export interface VbaProcedure {
  name: string
  kind: 'Sub' | 'Function' | 'Property'
  start_line: number
  end_line: number
  code: string
  // 構造化 LLM 注釈 (P1-2).
  annotation: string
  side_effects: string[]
  triggers: string[]
  calls: string[]
}

export interface VbaModule {
  name: string
  type: 'Module' | 'Class' | 'Form' | 'Document'
  code: string
  procedures: VbaProcedure[]
}

export interface WorkbookData {
  filename: string
  sheets: SheetInfo[]
  vba_modules: VbaModule[]
  external_links: string[]
  power_queries: PowerQueryItem[]
  analysis_risks: AnalysisRiskItem[]
}

// ---------- external functions ----------

export interface ExternalFunctionParam {
  name: string
  description: string
  required: boolean
  type: string
}

export interface ExternalFunction {
  name: string
  vendor: string
  short: string
  long: string
  signature: string
  params: ExternalFunctionParam[]
  returns: string
  examples: string[]
  notes: string[]
  doc_url: string
}

export interface ExternalFunctionRegistry {
  functions: ExternalFunction[]
  vendors: string[]
}

export interface ExternalFunctionUsageLocation {
  sheet: string
  coord: string
  formula: string
}

export interface ExternalFunctionUsageItem {
  name: string
  vendor: string
  short: string
  count: number
  locations: ExternalFunctionUsageLocation[]
  registered: boolean
}

export interface ExternalFunctionsUsed {
  items: ExternalFunctionUsageItem[]
  total_kinds: number
  total_uses: number
}

// ---------- diagrams ----------

export type DiagramNodeKind = 'sheet' | 'module' | 'procedure'
export type DiagramEdgeKind = 'formula' | 'call'

export interface DiagramNode {
  id: string
  label: string
  kind: DiagramNodeKind
  meta: Record<string, string | number>
}

export interface DiagramEdge {
  src: string
  dst: string
  weight: number
  kind: DiagramEdgeKind
}

export interface Diagram {
  kind: 'sheet_deps' | 'vba_calls'
  nodes: DiagramNode[]
  edges: DiagramEdge[]
}

export interface DiagramSet {
  sheet_deps: Diagram
  vba_calls: Diagram
}

/** Backend からの 4xx/5xx を表す. composable 側で投げ直す. */
export class BackendError extends Error {
  constructor(
    public status: number,
    public detail: string,
  ) {
    super(`backend error ${status}: ${detail}`)
    this.name = 'BackendError'
  }
}
