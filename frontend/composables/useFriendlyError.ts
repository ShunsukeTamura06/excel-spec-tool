/**
 * 任意の例外をユーザー向けの 1〜2 文の日本語に変換する.
 *
 * 原則: 生の例外メッセージ・スタックトレース・URL を出さない.
 * BackendError なら .detail (useBackend が既にフレンドリーに整形済) を返す.
 * その他は汎用フォールバック.
 *
 * 関連メモリ: feedback_no_traceback_in_ui.md
 */

import { BackendError } from '~/types/api'

export function friendlyMessage(err: unknown): string {
  if (err instanceof BackendError) {
    return err.detail || '処理に失敗しました。'
  }
  // それ以外は素の内容を漏らさず汎用文言で返す.
  // 元の例外は呼び出し側で console に残しているはず.
  return '予期しないエラーが発生しました。時間を置いて再度お試しください。'
}
