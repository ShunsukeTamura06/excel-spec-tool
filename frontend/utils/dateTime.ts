/** Backendが返すUTC日時を、日本時間の安定した表示へ変換する. */

const JST_TIME_ZONE = 'Asia/Tokyo'

function parseUtcTimestamp(value: string): Date | null {
  const trimmed = value.trim()
  if (!trimmed) return null
  const timestampWithZone = /(?:Z|[+-]\d{2}:?\d{2})$/i.test(trimmed)
    ? trimmed
    : `${trimmed}Z`
  const date = new Date(timestampWithZone)
  return Number.isNaN(date.getTime()) ? null : date
}

function jstParts(value: string): Record<string, string> | null {
  const date = parseUtcTimestamp(value)
  if (!date) return null
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: JST_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hourCycle: 'h23',
  }).formatToParts(date)
  return Object.fromEntries(parts.map(part => [part.type, part.value]))
}

/** UTC日時を `YYYY-MM-DD HH:mm JST` 形式で返す. */
export function formatJstDateTime(value: string): string {
  const parts = jstParts(value)
  if (!parts) return ''
  return `${parts.year}-${parts.month}-${parts.day} ${parts.hour}:${parts.minute} JST`
}

/** UTC日時を `HH:mm JST` 形式で返す. */
export function formatJstTime(value: string): string {
  const parts = jstParts(value)
  if (!parts) return ''
  return `${parts.hour}:${parts.minute} JST`
}
