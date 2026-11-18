import { all, attr, linked, one, problem, relId, value, withOfficePackage } from './ooxml.js'
import { monthSignature, periods, signature } from './meaning.js'
import { canonicalUnit, excelDecimal, unitIn } from './number.js'

const SHEET_LIMIT = 64
const CELL_LIMIT = 35_000
const ROW_LIMIT = 5000
const COL_LIMIT = 128
const addressPattern = /^([A-Z]{1,3})([1-9]\d{0,6})$/u

function column(name) {
  return [...name].reduce((index, char) => index * 26 + char.charCodeAt(0) - 64, 0)
}

function richText(node) {
  return value(one(node, 't')) + all(node, 'r').map(run => value(one(run, 't'))).join('')
}

function formatFor(id, code, original) {
  if (id === 9 || id === 10) return { type: 'number', unit: '%', factor: '100', code }
  if (code === undefined) {
    if ([14, 15, 16, 17, 22].includes(id)) return { type: 'date' }
    if ([0, 1, 2, 3, 4, 37, 38, 39, 40].includes(id)) return { type: 'number', unit: '', factor: '1' }
    return { type: 'unsupported' }
  }
  if (code.length > 512 || /\[\s*[<>=]/u.test(code)) return { type: 'unsupported', code }
  const sections = []
  let start = 0
  for (const token of code.matchAll(/"[^"]*"|\\.|;/gu)) {
    if (token[0] !== ';') continue
    sections.push(code.slice(start, token.index))
    start = token.index + 1
  }
  sections.push(code.slice(start))
  if (!sections.length || sections.length > 4) return { type: 'unsupported', code }
  const number = Number(original)
  const format = sections[number < 0 && sections.length >= 2 ? 1 : number === 0 && sections.length >= 3 ? 2 : 0]
  const literals = [...format.matchAll(/"([^"]*)"|\\(.)/gu)].map(m => m[1] ?? m[2]).join('')
  const syntax = format.replace(/"[^"]*"|\\.|[_*].|\[[^\]]*\]/gu, '')
  if (/[yd]/iu.test(syntax) && !/[0#%?]/u.test(syntax)) return { type: 'date', code }
  if (id === 0 && /^general$/iu.test(syntax)) return { type: 'number', unit: '', factor: '1', code }
  if (!/[0#]/u.test(syntax) || /[ymdhse?@/]/iu.test(syntax) || /[0#],[^0#]|[0#],$/u.test(syntax) ||
      (syntax.match(/%/gu)?.length ?? 0) > 1 ||
      /(?:[$€£]|[A-Z]{3})/u.test(literals.replace(/[¥￥]/gu, '')) ||
      /[$€£]/u.test(syntax) || /\[\$[^¥￥-]/u.test(format)) {
    return { type: 'unsupported', code }
  }
  const literalUnit = canonicalUnit(literals.trim()) || (/[¥￥]/u.test(format) ? '元' : '')
  if (literals.trim() && !literalUnit && literals.trim() !== '-') return { type: 'unsupported', code }
  if (syntax.includes('%')) return literalUnit ? { type: 'unsupported', code }
    : { type: 'number', unit: '%', factor: '100', code }
  return { type: 'number', unit: literalUnit,
    factor: '1', code }
}

function dateMonth(serial, date1904) {
  const days = Number(serial)
  if (!Number.isFinite(days) || days < 0 || days > 100_000) return ''
  const offset = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, days < 60 ? 31 : 30)
  return new Date(offset + Math.floor(days) * 86_400_000).toISOString().slice(0, 7)
}

function extractCell(node, address, shared, styles, date1904) {
  const type = attr(node, 't') ?? 'n'
  const original = value(one(node, 'v')).trim()
  const formulaNode = one(node, 'f')
  const formula = value(formulaNode).slice(0, 1000)
  const raw = excelDecimal(original)
  const styleIndex = Number(attr(node, 's') ?? 0)
  const config = styles[styleIndex] ?? (styleIndex === 0 && !styles.length ? { id: 0 } : { id: -1 })
  const style = formatFor(config.id, config.code, original)
  if (formulaNode) {
    return { address, raw: original, formula: formula || '(shared/array)', format: style.code,
      unit: style.unit ?? '', factor: style.factor ?? '1',
      reason: /\[[^\]]+\]/u.test(formula) ? 'externalFormula' : 'unverifiedFormula' }
  }
  if (type === 'e') return { address, raw: original, reason: 'cellError' }
  if (type === 's') {
    if (!/^\d+$/u.test(original) || shared[Number(original)] === undefined) throw problem('invalidOffice', address)
    return { address, text: shared[Number(original)] }
  }
  if (type === 'inlineStr') return { address, text: richText(one(node, 'is')) }
  if (type === 'str' || type === 'd') return { address, text: type === 'd' ? original.slice(0, 7) : original }
  if (!original) return {}
  if (type !== 'n') return { address, raw: original, reason: 'nonNumeric' }
  if (raw === null) return { address, raw: original, reason: 'nonNumeric' }
  if (style.type === 'date') return { address, text: dateMonth(raw, date1904), raw }
  if (style.type === 'unsupported') return { address, raw, format: style.code, reason: 'unknownFormat' }
  return { address, raw, unit: style.unit, factor: style.factor, format: style.code }
}

function worksheetCells(root, sheet, shared, styles, date1904) {
  if (root?.name !== 'worksheet') throw problem('invalidOffice', sheet)
  const matrix = []
  const used = new Set()
  for (const row of all(one(root, 'sheetData'), 'row')) {
    const cells = []
    for (const node of all(row, 'c')) {
      const address = attr(node, 'r')
      const match = addressPattern.exec(address ?? '')
      if (!match || used.has(address)) throw problem('invalidOffice', address ?? sheet)
      used.add(address)
      if (used.size > CELL_LIMIT) throw problem('sheetLimit', sheet)
      const position = column(match[1])
      const rowNumber = Number(match[2])
      if (rowNumber > ROW_LIMIT || position > COL_LIMIT) throw problem('sheetLimit', sheet)
      if (cells.length && cells[0].row !== rowNumber) throw problem('invalidOffice', address)
      if (attr(row, 'r') !== undefined && Number(attr(row, 'r')) !== rowNumber) throw problem('invalidOffice', address)
      const cell = extractCell(node, address, shared, styles, date1904)
      if (cell.text?.length > 1024) throw problem('sheetLimit', address)
      cells.push({ ...cell, col: position, row: rowNumber })
    }
    if (cells.length) matrix.push(cells.sort((a, b) => a.col - b.col))
  }
  matrix.sort((a, b) => a[0].row - b[0].row)
  if (matrix.some((row, index) => index && row[0].row === matrix[index - 1][0].row)) {
    throw problem('invalidOffice', sheet)
  }
  const headings = new Map()
  const cells = []
  let section = ''
  let precedingLabels = new Map()
  let labelRow = 0
  for (const row of matrix) {
    const periodCells = row.filter(c => c.text && periods(c.text).length)
    const hasValues = row.some(c => c.raw !== undefined && c.text === undefined || c.reason)
    if (!hasValues && periodCells.length) {
      headings.clear()
      section = ''
      for (const cell of periodCells) {
        const upper = labelRow === cell.row - 1 ? precedingLabels.get(cell.col) : ''
        const colTitle = [upper, cell.text].filter(Boolean).join(' / ')
        const seen = periods(colTitle)
        headings.set(cell.col, { colTitle, period: seen.length === 1 ? seen[0] : '',
          colUnits: [unitIn(upper ?? ''), unitIn(cell.text)] })
      }
    }
    if (!hasValues && !periodCells.length) {
      // A subheading such as Actual or Budget underneath the month is part of the column identity.
      for (const [col, heading] of headings) {
        const label = row.find(cell => cell.col === col && cell.text)?.text
        if (label) {
          const colTitle = heading.colTitle + ' / ' + label
          if (colTitle.length > 2048) throw problem('sheetLimit', sheet)
          headings.set(col, { ...heading, colTitle, colUnits: [...heading.colUnits, unitIn(label)] })
        }
      }
      const left = Math.min(...headings.keys())
      const labels = row.filter(cell => cell.col < left && cell.text?.trim())
      if (headings.size && labels.length) section = labels.map(cell => cell.text.trim()).join(' / ')
    }
    if (!hasValues) {
      labelRow = row[0].row
      precedingLabels = new Map(row.filter(cell => cell.text?.trim()).map(cell => [cell.col, cell.text.trim()]))
      continue
    }
    precedingLabels.clear()
    const firstData = Math.min(...headings.keys(), ...row.filter(c => c.raw !== undefined && c.text === undefined || c.reason).map(c => c.col))
    for (const cell of row) {
      if (cell.text !== undefined || (cell.raw === undefined && !cell.reason)) continue
      const labels = row.filter(c => c.col < Math.min(firstData, cell.col) && c.text?.trim()).map(c => c.text.trim())
      const label = labels.at(-1)
      const metric = labels.filter(s => !canonicalUnit(s) && !/^单位\s*[:：]/u.test(s)).at(-1) ?? ''
      const heading = headings.get(cell.col) ?? { colTitle: '', period: '', colUnits: [] }
      const rowTitle = [section, ...labels].filter(Boolean).join(' / ')
      if (rowTitle.length > 2048) throw problem('sheetLimit', sheet)
      const declared = [cell.unit, unitIn(section), ...labels.map(unitIn), ...heading.colUnits].filter(Boolean)
      const unit = declared[0] ?? ''
      const reason = cell.reason || (new Set(declared).size > 1 ? 'unitConflict' : '')
      cells.push({
        ...cell, unit, reason, sheet, metric, rowTitle, colTitle: heading.colTitle, period: heading.period,
        semantic: JSON.stringify([sheet, signature(rowTitle), monthSignature(heading.colTitle, heading.period), unit]),
        reusable: Boolean(label && metric && heading.period && heading.colTitle),
      })
    }
  }
  return { name: sheet, cells, scanned: used.size, rows: matrix.length }
}

function workbookInfo(root, links) {
  if (root?.name !== 'workbook') throw problem('wrongFile')
  const sheets = all(one(root, 'sheets'), 'sheet')
  if (sheets.length > SHEET_LIMIT) throw problem('workbookLimit')
  const names = new Set()
  return sheets.map(node => {
    const name = attr(node, 'name')
    const id = relId(node)
    if (!name || names.has(name) || !links.has(id)) throw problem('invalidOffice', name)
    names.add(name)
    return { name, id, supported: links.get(id).type === 'worksheet' && !links.get(id).external,
      hidden: attr(node, 'state') === 'hidden' || attr(node, 'state') === 'veryHidden' }
  })
}

export async function openWorkbook(bytes, sheetName) {
  return withOfficePackage(bytes, async ({ main, relations, xml }) => {
    const root = await xml(main)
    const links = await relations(main)
    const sheets = workbookInfo(root, links)
    if (sheetName === undefined) return { sheets }
    const selected = sheets.find(sheet => sheet.name === sheetName && sheet.supported)
    if (!selected) throw problem('noSheet', sheetName)
    const readType = async type => {
      const match = [...links.values()].find(r => r.type === type && !r.external)
      return match ? xml(match.part) : undefined
    }
    const formats = await readType('styles')
    const codes = new Map(all(one(formats, 'numFmts'), 'numFmt')
      .map(node => [Number(attr(node, 'numFmtId')), attr(node, 'formatCode')]))
    const styles = all(one(formats, 'cellXfs'), 'xf').map(node => {
      const id = Number(attr(node, 'numFmtId') ?? 0)
      return { id, code: codes.get(id) }
    })
    const shared = all(await readType('sharedStrings'), 'si').map(richText)
    const date1904 = ['1', 'true'].includes(attr(one(root, 'workbookPr'), 'date1904'))
    const table = await xml(linked(links, selected.id, 'worksheet'))
    return { sheets, sheet: worksheetCells(table, selected.name, shared, styles, date1904) }
  })
}
