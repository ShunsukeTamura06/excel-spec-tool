/**
 * 現在選択中のジョブを Pinia で管理する.
 *
 * - `currentJobId` を `localStorage` に永続化 (リロード耐性)
 * - `jobs` リストはサーバから都度取得し、Pinia には最新版をキャッシュ
 * - ページコンポーネントは `selectJob` で URL ナビゲーションも兼ねる
 */

import { defineStore } from 'pinia'
import type { JobMeta } from '~/types/api'

const STORAGE_KEY = 'excel-spec-tool/current-job-id'

export const useJobStore = defineStore('job', () => {
  const currentJobId = ref<string | null>(null)
  const jobs = ref<JobMeta[]>([])
  const loading = ref(false)
  const error = ref<string | null>(null)

  // 初期化時に localStorage から復元 (client only)
  if (import.meta.client) {
    const saved = window.localStorage.getItem(STORAGE_KEY)
    if (saved) currentJobId.value = saved
  }

  const currentJob = computed<JobMeta | null>(() => {
    if (!currentJobId.value) return null
    return jobs.value.find(j => j.job_id === currentJobId.value) ?? null
  })

  function setCurrentJobId(jobId: string | null) {
    currentJobId.value = jobId
    if (import.meta.client) {
      if (jobId) {
        window.localStorage.setItem(STORAGE_KEY, jobId)
      } else {
        window.localStorage.removeItem(STORAGE_KEY)
      }
    }
  }

  async function refreshJobs() {
    loading.value = true
    error.value = null
    try {
      jobs.value = await useBackend().listJobs()
      // 削除されたジョブを選択していたらクリア
      if (currentJobId.value && !jobs.value.some(j => j.job_id === currentJobId.value)) {
        setCurrentJobId(null)
      }
    } catch (e) {
      error.value = e instanceof Error ? e.message : String(e)
    } finally {
      loading.value = false
    }
  }

  return {
    currentJobId,
    currentJob,
    jobs,
    loading,
    error,
    setCurrentJobId,
    refreshJobs,
  }
})
