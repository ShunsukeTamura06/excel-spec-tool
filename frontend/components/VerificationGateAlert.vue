<script setup lang="ts">
/** 変更後ファイルに対する構造検証policy gateの判定を表示する. */

import type { VerificationReportData } from '~/types/api'

const props = defineProps<{
  report: VerificationReportData
  newJobId: string
}>()

const presentation = computed(() => {
  if (props.report.status === 'passed') {
    return {
      color: 'success' as const,
      icon: 'i-lucide-shield-check',
      title: '構造検証に合格',
      description:
        `${props.report.expected_change_count}件の予定変更と実差分が一致しました。` +
        `新しいジョブ (${props.newJobId}) を作成しました。`,
    }
  }
  if (props.report.status === 'needs_review') {
    return {
      color: 'warning' as const,
      icon: 'i-lucide-shield-alert',
      title: '適用済み・確認が必要',
      description:
        `${props.report.warnings.join(' ')} ` +
        `新しいジョブ (${props.newJobId}) の差分を確認してください。`,
    }
  }
  return {
    color: 'error' as const,
    icon: 'i-lucide-shield-x',
    title: '構造検証に不合格',
    description: props.report.violations.map(item => item.message).join(' '),
  }
})
</script>

<template>
  <UAlert
    :color="presentation.color"
    variant="subtle"
    :icon="presentation.icon"
    :title="presentation.title"
    :description="presentation.description"
  />
</template>
