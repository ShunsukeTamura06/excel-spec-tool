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
}

export interface SpecResponse {
  spec_md: string
  meta: JobMeta
}

export interface ReferenceItem {
  kind: 'formula' | 'vba'
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

export interface ChatReply {
  reply: string
  history: ChatMessage[]
}

export interface ExtractResponse {
  job_id: string
}

export interface AnalyzeResponse {
  status: string
}

export interface DeleteResponse {
  deleted: boolean
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
