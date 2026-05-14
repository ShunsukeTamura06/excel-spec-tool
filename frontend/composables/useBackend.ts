/**
 * Backend (FastAPI) を叩く typed クライアント.
 *
 * Nuxt の `$fetch` (ofetch) を薄くラップし、エラーを `BackendError` に正規化する。
 * Frontend は本 composable 経由でのみ Backend にアクセスする (旧 Streamlit の
 * frontend_streamlit/api_client.py に対応する役割).
 */

import {
  BackendError,
  type AnalyzeResponse,
  type ChatMessage,
  type ChatReply,
  type DeleteResponse,
  type DiagramSet,
  type ExtractResponse,
  type JobMeta,
  type ReferenceItem,
  type SpecResponse,
} from '~/types/api'

interface FetchErrorLike {
  statusCode?: number
  status?: number
  data?: { detail?: unknown } | string
  message?: string
}

function toBackendError(err: unknown): BackendError {
  const e = err as FetchErrorLike
  const status = e?.statusCode ?? e?.status ?? 0
  let detail = ''
  if (typeof e?.data === 'string') {
    detail = e.data
  } else if (e?.data && typeof e.data === 'object' && 'detail' in e.data) {
    const d = (e.data as { detail?: unknown }).detail
    detail = typeof d === 'string' ? d : JSON.stringify(d)
  } else {
    detail = e?.message ?? 'unknown error'
  }
  return new BackendError(status, detail)
}

export function useBackend() {
  const config = useRuntimeConfig()
  const baseURL = config.public.backendUrl

  /** 共通呼び出しラッパ. 失敗時は BackendError を throw する. */
  async function call<T>(
    path: string,
    opts: Parameters<typeof $fetch>[1] = {},
  ): Promise<T> {
    try {
      return await $fetch<T>(path, { baseURL, ...opts })
    } catch (e) {
      throw toBackendError(e)
    }
  }

  return {
    /** GET /health — 接続確認. throw せず boolean を返す. */
    async health(): Promise<boolean> {
      try {
        await $fetch('/health', { baseURL })
        return true
      } catch {
        return false
      }
    },

    /** GET /jobs */
    async listJobs(): Promise<JobMeta[]> {
      const res = await call<{ jobs: JobMeta[] }>('/jobs')
      return res.jobs ?? []
    },

    /** POST /extract — multipart アップロード. */
    async extract(file: File): Promise<string> {
      const form = new FormData()
      form.append('file', file)
      const res = await call<ExtractResponse>('/extract', {
        method: 'POST',
        body: form,
      })
      return res.job_id
    },

    /** POST /analyze/{job_id} */
    async analyze(jobId: string): Promise<AnalyzeResponse> {
      return await call<AnalyzeResponse>(`/analyze/${jobId}`, { method: 'POST' })
    },

    /** GET /spec/{job_id} */
    async getSpec(jobId: string): Promise<SpecResponse> {
      return await call<SpecResponse>(`/spec/${jobId}`)
    },

    /** GET /references/{job_id}?target=... */
    async getReferences(jobId: string, target: string): Promise<ReferenceItem[]> {
      const res = await call<{ refs: ReferenceItem[] }>(
        `/references/${jobId}`,
        { query: { target } },
      )
      return res.refs ?? []
    },

    /** POST /chat/{job_id} */
    async chat(jobId: string, message: string): Promise<ChatReply> {
      return await call<ChatReply>(`/chat/${jobId}`, {
        method: 'POST',
        body: { message },
      })
    },

    /** GET /chat/{job_id}/history */
    async getChatHistory(jobId: string): Promise<ChatMessage[]> {
      const res = await call<{ history: ChatMessage[] }>(`/chat/${jobId}/history`)
      return res.history ?? []
    },

    /** DELETE /jobs/{job_id} */
    async deleteJob(jobId: string): Promise<boolean> {
      const res = await call<DeleteResponse>(`/jobs/${jobId}`, { method: 'DELETE' })
      return res.deleted ?? false
    },

    /** GET /diagrams/{job_id} */
    async getDiagrams(jobId: string): Promise<DiagramSet> {
      return await call<DiagramSet>(`/diagrams/${jobId}`)
    },
  }
}
