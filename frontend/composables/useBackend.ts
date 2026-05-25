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
  type ChatReply,
  type ChatSessionMeta,
  type ChatSessionResponse,
  type DeleteResponse,
  type DiagramSet,
  type ExternalFunctionRegistry,
  type ExternalFunctionsUsed,
  type ExtractResponse,
  type FeedbackInput,
  type FeedbackResponse,
  type JobMeta,
  type ReferenceItem,
  type SpecResponse,
  type WorkbookData,
} from '~/types/api'

// 「重い」エンドポイント (extract / analyze / chat) は LLM や全セル抽出を含むため
// 数分かかることがある. ブラウザ側で早めにタイムアウトすると体感が悪く、
// ユーザーは「壊れた」と感じてしまうので寛容に取る.
const HEAVY_TIMEOUT_MS = 10 * 60 * 1000 // 10 分
const DEFAULT_TIMEOUT_MS = 60 * 1000 // 1 分

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
