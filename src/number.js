import Decimal from 'decimal.js'

const Num = Decimal.clone({ precision: 60, rounding: Decimal.ROUND_HALF_UP })
export const UNIT = Object.freeze({
  '元': ['cny', '1'], '万元': ['cny', '10000'], '亿元': ['cny', '100000000'],
  '%': ['rate', '0.01'], '百分点': ['rateDelta', '0.01'],
  '个': ['count', '1'], '项': ['item', '1'], '人': ['person', '1'], '': ['number', '1'],
})
const UNITS = '亿元|万元|元|个百分点|百分点|percentage points?|pp|%|个|项|人'
const MAGNITUDE = '(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d+)?'
const NUMBER = '[+-]?' + MAGNITUDE
const SIGNED = '(?:\\(\\s*(?:[¥￥]\\s*)?' + MAGNITUDE + '\\s*(?:' + UNITS + ')?\\s*\\)|(?:[+-]\\s*)?(?:[¥￥]\\s*)?' + MAGNITUDE + ')'
const tokenPattern = new RegExp('(?<![\\p{N}.,₀₁₂₃₄₅₆₇₈₉])' + SIGNED + '(?:\\s*(?:' + UNITS + ')(?![a-z]))?(?![\\d.,])', 'giu')

export function normalize(text) {
  return text.normalize('NFKC').replace(/[−–—]/gu, '-').trim()
}

export function canonicalUnit(unit) {
  if (/^(?:个百分点|百分点|pp|percentage points?)$/iu.test(unit)) return '百分点'
  return Object.hasOwn(UNIT, unit) ? unit : ''
}

export function unitIn(text) {
  const s = normalize(text)
  const full = canonicalUnit(s)
  if (full) return full
  const label = /(?:单位\s*[:：]?\s*|\()\s*(亿元|万元|元|百分点|个百分点|%|个|项|人)\s*(?:\)|$)/u.exec(s)
  return label ? canonicalUnit(label[1]) : ''
}

export function numberTokens(text) {
  const normalizedText = normalize(text)
  const hits = []
  for (const token of normalizedText.matchAll(tokenPattern)) {
    const shown = token[0]
    const match = new RegExp(NUMBER, 'u').exec(shown.replace(/\s+/gu, '').replace(/[¥￥]/gu, ''))
    const literal = match[0].replaceAll(',', '')
    const raw = shown.startsWith('(') ? '-' + literal : literal
    if (raw.length > 38) continue
    const suffix = new RegExp('(?:' + UNITS + ')(?=\\s*\\)?$)', 'iu').exec(shown)
    const unit = suffix ? canonicalUnit(suffix[0]) : /[¥￥]/u.test(shown) ? '元' : ''
    hits.push({ shown, raw, unit, precision: raw.split('.')[1]?.length ?? 0,
      start: token.index, end: token.index + shown.length })
  }
  return { text: normalizedText, hits }
}

export function excelDecimal(value) {
  const s = value.trim()
  const match = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:e([-+]?\d{1,3}))?$/iu.exec(s)
  if (!match || s.length > 38 || Math.abs(Number(match[1] ?? 0)) > 100) return null
  return s
}

export function compareNumbers(slide, cell) {
  if (cell.reason) return { status: 'verify', reason: cell.reason }
  const source = UNIT[slide.unit]
  const target = UNIT[cell.unit]
  if (!source || !target || source[0] !== target[0]) return { status: 'verify', reason: 'unitMismatch' }
  const workbook = new Num(cell.raw).times(cell.factor).times(target[1]).div(source[1])
  const rounded = workbook.toDecimalPlaces(slide.precision)
  const delta = new Num(slide.raw).minus(workbook)
  const differenceUnit = slide.unit === '%' ? '百分点' : slide.unit
  return {
    status: new Num(slide.raw).eq(rounded) ? 'match' : 'difference',
    workbook: workbook.toString(),
    rounded: rounded.toFixed(slide.precision),
    precision: slide.precision,
    delta: delta.toFixed(Math.max(slide.precision, 1)),
    exactDelta: delta.toString(),
    direction: delta.cmp(0),
    unit: slide.unit,
    differenceUnit,
  }
}
