import { normalize } from './number.js'

const MONTH = /(?<!\d)(20\d{2})(?:\s*[-/.年]\s*)(0?[1-9]|1[0-2])(?!\d)(?:\s*月)?/gu
const SHORT_MONTH = /(?<![\d.])(0?[1-9]|1[0-2])\s*月(?!\s*\d)/gu
const EN_MONTH = /\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)(?:\.?\s*,?\s*(20\d{2}))?\b/gu
const EN_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

export function periods(text) {
  return [...new Set([...normalize(text).matchAll(MONTH)].map(
    match => match[1] + '-' + match[2].padStart(2, '0'),
  ))]
}

export function signature(text) {
  return normalize(text).toLowerCase().replace(/[\s:：/|]+/gu, ' ').trim()
}

export function monthSignature(text, period) {
  return signature(normalize(text).replace(MONTH, (original, year, month) =>
    year + '-' + month.padStart(2, '0') === period ? '<month>' : original))
}

function slidePeriodTokens(text, selected, substitute) {
  const year = selected.slice(0, 4)
  const month = n => String(n).padStart(2, '0')
  return normalize(text)
    .replace(MONTH, (original, y, m) => substitute(original, y + '-' + month(m)))
    .replace(SHORT_MONTH, (original, m) => substitute(original, year + '-' + month(m)))
    .replace(EN_MONTH, (original, name, y) => substitute(original, (y || year) + '-' + month(EN_MONTHS.indexOf(name.slice(0, 3)) + 1)))
}

export function slidePeriods(text, selected) {
  const matches = new Set()
  slidePeriodTokens(text, selected, (_, period) => { matches.add(period); return '' })
  return [...matches]
}

export function slideSignature(text, selected) {
  return signature(slidePeriodTokens(text, selected, (original, period) => period === selected ? '<month>' : original))
}

export function validMonth(text) {
  return /^20\d{2}-(?:0[1-9]|1[0-2])$/u.test(text)
}
