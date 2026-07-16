/**
 * Backend (FastAPI) を叩く typed クライアント.
 *
 * Nuxt の `$fetch` (ofetch) を薄くラップし、エラーを `BackendError` に
 * 正規化する。生の例外メッセージや stacktrace をフロントに出さず、
 * 必ず日本語のフレンドリーな detail を載せる責務を持つ。
 *
 * 関連メモリ: feedback_no_traceback_in_ui.md
 */

import {
  BackendError,
  type AnalyzeResponse,
  type ChatMessage,
  type ChatProgressEvent,
  type ChatReply,
  type ChatSessionMeta,
  type ChatSessionResponse,
  type ChangeBrief,
  type ChatStreamEventName,
  type DeleteResponse,
  type DiagramSet,
  type ExternalFunctionRegistry,
  type ExternalFunctionsUsed,
  type ExtractResponse,
  type FeedbackInput,
  type FeedbackResponse,
  type FormulaFixRequest,
  type FormulaFixResponse,
  type JobMeta,
  type NamedRangeFixRequest,
  type NamedRangeFixResponse,
  type MutationPlanData,
  type SafeChangePlanData,
  type WorkbookDiagnosis,
  type ReferenceItem,
  type SpecResponse,
  type WorkbookData,
  type WorkbookDiffData,
} from '~/types/api'

// 「重い」エンドポイント (extract / analyze / chat) は LLM や全セル抽出を含むため
// 数分かかることがある. ブラウザ側で早めにタイムアウトすると体感が悪く、
// ユーザーは「壊れた」と感じてしまうので寛容に取る.
const HEAVY_TIMEOUT_MS = 10 * 60 * 1000 // 10 分
const DEFAULT_TIMEOUT_MS = 60 * 1000 // 1 分

interface ChatStreamHandlers {
  onEvent?: (event: ChatStreamEventName, data: ChatProgressEvent | ChatReply) => void
}

interface FetchErrorLike {
  name?: string
  statusCode?: number
  status?: number
  data?: { detail?: unknown } | string
  message?: string
  cause?: { name?: string; message?: string }
}

/**
 * 任意の例外を BackendError に正規化する.
 *
 * 重要: 生の fetch メッセージ・例外名・URL をユーザーに出さない.
 * (TraceBack 非表示の原則 — feedback_no_traceback_in_ui.md)
 */
function toBackendError(err: unknown, context: string): BackendError {
  const e = (err ?? {}) as FetchErrorLike
  const status = e.statusCode ?? e.status ?? 0
  const rawMsg = `${e.message ?? ''} ${e.cause?.message ?? ''}`.toLowerCase()
  const isAbort = e.name === 'AbortError' || e.cause?.name === 'AbortError'
  const isTimeout = isAbort || rawMsg.includes('timeout') || rawMsg.includes('timed out')

  // 開発者向けに console には残す (ユーザー UI には出さない)
  // eslint-disable-next-line no-console
  console.warn(`[backend] ${context} failed:`, err)

  // ネットワーク到達不能 / タイムアウト系 (HTTP レスポンスなし)
  if (status === 0) {
    if (isTimeout) {
      return new BackendError(
        0,
        'サーバーの応答がタイムアウトしました。大きなファイルや LLM 呼び出しは数分かかります。'
        + ' もう一度試すか、しばらく時間を置いてから実行してください。',
      )
    }
    return new BackendError(
      0,
      'Backend に接続できません。サーバーが起動しているか、ネットワーク設定を確認してください。',
    )
  }

  // HTTP レスポンスはあるがエラー (4xx / 5xx)
  // backend は FastAPI HTTPException の detail を返してくれるのでそれを優先.
  let detail = ''
  if (typeof e.data === 'string') {
    detail = e.data
  } else if (e.data && typeof e.data === 'object' && 'detail' in e.data) {
    const d = (e.data as { detail?: unknown }).detail
    if (typeof d === 'string') {
      detail = d
    } else if (Array.isArray(d) && d.length > 0) {
      // FastAPI Pydantic バリデーションエラー: detail は
      // [{ loc: [...], msg: "...", type: "..." }, ...] 形式の配列
      // 先頭エラーの "loc: msg" を取り出す
      const first = d[0] as { msg?: unknown; loc?: unknown }
      const msg = typeof first?.msg === 'string' ? first.msg : ''
      const loc = Array.isArray(first?.loc) ? first.loc.filter(x => x !== 'body').join('.') : ''
      if (msg) {
        detail = loc ? `${loc}: ${msg}` : msg
      }
    } else if (d && typeof d === 'object') {
      const structured = d as {
        message?: unknown
        verification?: { status?: unknown }
      }
      if (structured.verification?.status === 'failed') {
        detail = '構造検証で予定外の差分を検出したため、修正版の提供を停止しました。原本は変更されていません。'
      } else if (typeof structured.message === 'string') {
        detail = structured.message
      }
    }
  }

  if (!detail) {
    if (status >= 500) {
      detail = 'サーバー側で予期しないエラーが発生しました。時間を置いて再度お試しください。'
    } else if (status === 404) {
      detail = '対象が見つかりませんでした。'
    } else if (status === 409) {
      detail = 'この操作はまだ実行できません。前のステップが完了しているか確認してください。'
    } else if (status === 413) {
      detail = 'ファイルが大きすぎます。サイズを確認してください。'
    } else if (status === 422) {
      detail = '入力形式に問題があります。'
    } else if (status === 400) {
      detail = '入力に問題があります。'
    } else {
      detail = '処理に失敗しました。'
    }
  }
  return new BackendError(status, detail)
}

async function backendErrorFromResponse(response: Response): Promise<BackendError> {
  let detail = ''
  try {
    const data = await response.json() as { detail?: unknown }
    if (typeof data.detail === 'string') {
      detail = data.detail
    }
  } catch {
    // JSON でないエラー本文はユーザーに出さない。
  }
  if (!detail) {
    detail = response.status >= 500
      ? 'サーバー側で予期しないエラーが発生しました。時間を置いて再度お試しください。'
      : '処理に失敗しました。'
  }
  return new BackendError(response.status, detail)
}

function parseSseEvent(raw: string): { event: ChatStreamEventName; data: unknown } | null {
  let event: ChatStreamEventName = 'status'
  const dataLines: string[] = []
  for (const line of raw.split('\n')) {
    if (line.startsWith('event:')) {
      const value = line.slice('event:'.length).trim()
      if (
        value === 'status'
        || value === 'tool_start'
        || value === 'tool_result'
        || value === 'final'
        || value === 'error'
      ) {
        event = value
      }
    } else if (line.startsWith('data:')) {
      dataLines.push(line.slice('data:'.length).trimStart())
    }
  }
  if (dataLines.length === 0) return null
  return { event, data: JSON.parse(dataLines.join('\n')) }
}

export function useBackend() {
  const config = useRuntimeConfig()
  const baseURL = config.public.backendUrl

  /** 共通呼び出しラッパ. 失敗時は BackendError を throw する. */
  async function call<T>(
    context: string,
    path: string,
    opts: Parameters<typeof $fetch>[1] = {},
  ): Promise<T> {
    try {
      return await $fetch<T>(path, { baseURL, ...opts })
    } catch (e) {
      throw toBackendError(e, context)
    }
  }

  return {
    /** GET /health — 接続確認. throw せず boolean を返す. */
    async health(): Promise<boolean> {
      try {
        await $fetch('/health', { baseURL, timeout: 5_000 })
        return true
      } catch {
        return false
      }
    },

    /** GET /system/llm-status — LLM が設定済か返す. throw せず null を返す
     *  (backend ダウン時はバナーを出さない判断に使えるよう). */
    async llmStatus(): Promise<{
      configured: boolean
      mode: string
      pro_model: string
      fast_model: string
    } | null> {
      try {
        // URL を string 扱いにして typed-routes の再帰的な型解決 (TS2321) を回避する
        return await $fetch<{
          configured: boolean
          mode: string
          pro_model: string
          fast_model: string
        }>('/system/llm-status' as string, { baseURL, timeout: 5_000 })
      } catch {
        return null
      }
    },

    /** GET /jobs */
    async listJobs(): Promise<JobMeta[]> {
      const res = await call<{ jobs: JobMeta[] }>('listJobs', '/jobs', {
        timeout: DEFAULT_TIMEOUT_MS,
      })
      return res.jobs ?? []
    },

    /** POST /extract — multipart アップロード. 大きなファイルだと数分かかる. */
    async extract(file: File): Promise<ExtractResponse> {
      const form = new FormData()
      form.append('file', file)
      return await call<ExtractResponse>('extract', '/extract', {
        method: 'POST',
        body: form,
        timeout: HEAVY_TIMEOUT_MS,
      })
    },

    /** POST /analyze/{job_id} — LLM 注釈付与. 数分かかる場合がある. */
    async analyze(jobId: string): Promise<AnalyzeResponse> {
      return await call<AnalyzeResponse>('analyze', `/analyze/${jobId}`, {
        method: 'POST',
        timeout: HEAVY_TIMEOUT_MS,
      })
    },

    /** GET /spec/{job_id} */
    async getSpec(jobId: string): Promise<SpecResponse> {
      return await call<SpecResponse>('getSpec', `/spec/${jobId}`, {
        timeout: DEFAULT_TIMEOUT_MS,
      })
    },

    /** GET /diagnosis/{job_id} — 根拠付きの一般ユーザー向け Excel 診断 */
    async getDiagnosis(jobId: string): Promise<WorkbookDiagnosis> {
      return await call<WorkbookDiagnosis>('getDiagnosis', `/diagnosis/${jobId}`, {
        timeout: DEFAULT_TIMEOUT_MS,
      })
    },

    /** POST /change-request/{job_id} — 業務要望を改修依頼書へ整理 */
    async createChangeBrief(
      jobId: string,
      requestedOutcome: string,
      featureId?: string,
    ): Promise<ChangeBrief> {
      return await call<ChangeBrief>('createChangeBrief', `/change-request/${jobId}`, {
        method: 'POST',
        body: {
          requested_outcome: requestedOutcome,
          feature_id: featureId || null,
        },
        timeout: DEFAULT_TIMEOUT_MS,
      })
    },

    /** GET /references/{job_id}?target=... */
    async getReferences(jobId: string, target: string): Promise<ReferenceItem[]> {
      const res = await call<{ refs: ReferenceItem[] }>(
        'getReferences',
        `/references/${jobId}`,
        { query: { target }, timeout: DEFAULT_TIMEOUT_MS },
      )
      return res.refs ?? []
    },

    /** POST /chat/{job_id} — LLM 応答 (tool ループあり). 数分かかる場合がある. */
    async chat(jobId: string, message: string, sessionId = 'default'): Promise<ChatReply> {
      return await call<ChatReply>('chat', `/chat/${jobId}`, {
        method: 'POST',
        body: { message },
        query: { session_id: sessionId },
        timeout: HEAVY_TIMEOUT_MS,
      })
    },

    /** POST /chat/{job_id}/stream — 進捗イベントを受け取りながら LLM 応答を待つ. */
    async chatStream(
      jobId: string,
      message: string,
      sessionId = 'default',
      handlers: ChatStreamHandlers = {},
    ): Promise<ChatReply> {
      try {
        const root = String(baseURL || '').replace(/\/$/, '')
        const query = new URLSearchParams({ session_id: sessionId })
        const response = await fetch(`${root}/chat/${jobId}/stream?${query.toString()}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ message }),
        })
        if (!response.ok) {
          throw await backendErrorFromResponse(response)
        }
        if (!response.body) {
          throw new BackendError(0, 'サーバーから進捗ストリームを取得できませんでした。')
        }

        const reader = response.body.getReader()
        const decoder = new TextDecoder()
        let buffer = ''
        let finalReply: ChatReply | null = null
        let streamError = ''

        const dispatch = (raw: string) => {
          const parsed = parseSseEvent(raw)
          if (!parsed) return
          if (parsed.event === 'final') {
            finalReply = parsed.data as ChatReply
          } else if (parsed.event === 'error') {
            const data = parsed.data as Partial<ChatProgressEvent>
            streamError = data.message || '応答生成中にエラーが発生しました。'
          }
          handlers.onEvent?.(parsed.event, parsed.data as ChatProgressEvent | ChatReply)
        }

        while (true) {
          const { done, value } = await reader.read()
          if (done) break
          buffer += decoder.decode(value, { stream: true }).replace(/\r\n/g, '\n')
          let boundary = buffer.indexOf('\n\n')
          while (boundary >= 0) {
            const raw = buffer.slice(0, boundary)
            buffer = buffer.slice(boundary + 2)
            dispatch(raw)
            boundary = buffer.indexOf('\n\n')
          }
        }
        buffer += decoder.decode().replace(/\r\n/g, '\n')
        if (buffer.trim()) dispatch(buffer)

        if (streamError) {
          throw new BackendError(500, streamError)
        }
        if (!finalReply) {
          throw new BackendError(0, '応答の受信が完了しませんでした。もう一度お試しください。')
        }
        return finalReply
      } catch (e) {
        if (e instanceof BackendError) throw e
        throw toBackendError(e, 'chatStream')
      }
    },

    /** GET /chat/{job_id}/history */
    async getChatHistory(jobId: string, sessionId = 'default'): Promise<ChatMessage[]> {
      const res = await call<{ history: ChatMessage[] }>(
        'getChatHistory',
        `/chat/${jobId}/history`,
        { query: { session_id: sessionId }, timeout: DEFAULT_TIMEOUT_MS },
      )
      return res.history ?? []
    },

    /** GET /chat/{job_id}/sessions */
    async listChatSessions(jobId: string, includeArchived = true): Promise<ChatSessionMeta[]> {
      const res = await call<{ sessions: ChatSessionMeta[] }>(
        'listChatSessions',
        `/chat/${jobId}/sessions`,
        { query: { include_archived: includeArchived }, timeout: DEFAULT_TIMEOUT_MS },
      )
      return res.sessions ?? []
    },

    /** POST /chat/{job_id}/sessions */
    async createChatSession(jobId: string, title = '新しい相談'): Promise<ChatSessionMeta> {
      const res = await call<ChatSessionResponse>('createChatSession', `/chat/${jobId}/sessions`, {
        method: 'POST',
        body: { title },
        timeout: DEFAULT_TIMEOUT_MS,
      })
      return res.session
    },

    /** PATCH /chat/{job_id}/sessions/{session_id} */
    async updateChatSession(
      jobId: string,
      sessionId: string,
      patch: { title?: string; archived?: boolean },
    ): Promise<ChatSessionMeta> {
      const res = await call<ChatSessionResponse>(
        'updateChatSession',
        `/chat/${jobId}/sessions/${sessionId}`,
        {
          method: 'PATCH',
          body: patch,
          timeout: DEFAULT_TIMEOUT_MS,
        },
      )
      return res.session
    },

    /** DELETE /jobs/{job_id} */
    async deleteJob(jobId: string): Promise<boolean> {
      const res = await call<DeleteResponse>('deleteJob', `/jobs/${jobId}`, {
        method: 'DELETE',
        timeout: DEFAULT_TIMEOUT_MS,
      })
      return res.deleted ?? false
    },

    /** GET /diagrams/{job_id} */
    async getDiagrams(jobId: string): Promise<DiagramSet> {
      return await call<DiagramSet>('getDiagrams', `/diagrams/${jobId}`, {
        timeout: DEFAULT_TIMEOUT_MS,
      })
    },

    /** GET /workbook/{job_id} — 抽出済み構造 (Sheets/VBA) */
    async getWorkbook(jobId: string): Promise<WorkbookData> {
      return await call<WorkbookData>('getWorkbook', `/workbook/${jobId}`, {
        timeout: DEFAULT_TIMEOUT_MS,
      })
    },

    /** GET /external-functions — 全ベンダーの登録済み関数定義 */
    async getExternalFunctions(): Promise<ExternalFunctionRegistry> {
      return await call<ExternalFunctionRegistry>(
        'getExternalFunctions',
        '/external-functions',
        { timeout: DEFAULT_TIMEOUT_MS },
      )
    },

    /** GET /external-functions/used/{job_id} — このブックでの使用箇所 */
    async getExternalFunctionsUsed(jobId: string): Promise<ExternalFunctionsUsed> {
      return await call<ExternalFunctionsUsed>(
        'getExternalFunctionsUsed',
        `/external-functions/used/${jobId}`,
        { timeout: DEFAULT_TIMEOUT_MS },
      )
    },

    /** GET /diff?before_job_id=...&after_job_id=... — 2ジョブ間の構造差分 + 波及範囲 */
    async getDiff(beforeJobId: string, afterJobId: string): Promise<WorkbookDiffData> {
      const res = await call<{ diff: WorkbookDiffData }>('getDiff', '/diff', {
        query: { before_job_id: beforeJobId, after_job_id: afterJobId },
        timeout: HEAVY_TIMEOUT_MS,
      })
      return res.diff
    },

    /** POST /jobs/{jobId}/named-range-fix — 名前定義修正を適用し、新ジョブと自己検証diffを得る */
    async applyNamedRangeFix(
      jobId: string,
      body: NamedRangeFixRequest,
    ): Promise<NamedRangeFixResponse> {
      return await call<NamedRangeFixResponse>(
        'applyNamedRangeFix',
        `/jobs/${jobId}/named-range-fix`,
        { method: 'POST', body, timeout: HEAVY_TIMEOUT_MS },
      )
    },

    /** POST /jobs/{jobId}/formula-fix — 固定参照置換/範囲拡張を適用し、新ジョブと自己検証diffを得る */
    async applyFormulaFix(jobId: string, body: FormulaFixRequest): Promise<FormulaFixResponse> {
      return await call<FormulaFixResponse>('applyFormulaFix', `/jobs/${jobId}/formula-fix`, {
        method: 'POST',
        body,
        timeout: HEAVY_TIMEOUT_MS,
      })
    },

    /** POST /jobs/{jobId}/change-plan — 原本を書き換えず、範囲拡張の予定差分を作る */
    async createSafeChangePlan(
      jobId: string,
      oldRef: string,
      newRef: string,
    ): Promise<SafeChangePlanData> {
      return await call<SafeChangePlanData>('createSafeChangePlan', `/jobs/${jobId}/change-plan`, {
        method: 'POST',
        body: { kind: 'range_expansion', old_ref: oldRef, new_ref: newRef },
        timeout: DEFAULT_TIMEOUT_MS,
      })
    },

    /** POST /jobs/{jobId}/change-plan/execute — 画面で確認した同一計画を実行する */
    async executeSafeChangePlan(
      jobId: string,
      plan: MutationPlanData,
    ): Promise<FormulaFixResponse> {
      return await call<FormulaFixResponse>(
        'executeSafeChangePlan',
        `/jobs/${jobId}/change-plan/execute`,
        { method: 'POST', body: { plan }, timeout: HEAVY_TIMEOUT_MS },
      )
    },

    /** GET /jobs/{jobId}/download — 修正版ExcelをBlobで取得する */
    async downloadWorkbook(jobId: string): Promise<Blob> {
      try {
        const root = String(baseURL || '').replace(/\/$/, '')
        const response = await fetch(`${root}/jobs/${jobId}/download`)
        if (!response.ok) throw await backendErrorFromResponse(response)
        return await response.blob()
      } catch (e) {
        if (e instanceof BackendError) throw e
        throw toBackendError(e, 'downloadWorkbook')
      }
    },

    /** POST /feedback — フィードバック 1 件を送信. */
    async submitFeedback(input: FeedbackInput): Promise<FeedbackResponse> {
      return await call<FeedbackResponse>('submitFeedback', '/feedback', {
        method: 'POST',
        body: input,
        timeout: DEFAULT_TIMEOUT_MS,
      })
    },
  }
}
