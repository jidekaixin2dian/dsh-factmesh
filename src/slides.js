import { all, attr, linked, one, problem, relId, value, withOfficePackage } from './ooxml.js'
import { slidePeriods, slideSignature, signature } from './meaning.js'
import { normalize, numberTokens, unitIn } from './number.js'

const SLIDE_LIMIT = 180
const OBJECT_LIMIT = 6000
const NUMBER_LIMIT = 8000
const TEXT_LIMIT = 16_384

function paragraphs(body) {
  return all(body, 'p').map(p => p.children.map(child => {
    if (child.name === 'br') return '\n'
    if (child.name === 'r' || child.name === 'fld') return value(one(child, 't'))
    return ''
  }).join('')).join('\n')
}

function textMetrics(line, previous = '', contextUnit = '') {
  const s = normalize(line)
  const suffix = s.replace(/(?<!\d)(?:20\d{2}\s*[-/.年]\s*(?:0?[1-9]|1[0-2])(?!\d)(?:\s*月)?|(?:0?[1-9]|1[0-2])\s*月(?!\s*\d))/gu, ' ')
  if (!/\d/u.test(suffix)) return { hits: [], unknown: false }
  const { hits, text } = numberTokens(suffix)
  if (hits.length !== 1) return { hits: [], unknown: true }
  const hit = hits[0]
  const prefix = text.slice(0, hit.start).trim()
  const label = prefix.replace(/[\s:：=]+$/u, '').trim() || previous
  const tail = text.slice(hit.end).trim()
  if (!label || !/[\p{L}]/u.test(label) || (tail && !/^[。.!！]$/u.test(tail)) ||
      /[\d.,]$/u.test(prefix) || (!hit.unit && !contextUnit && !/[:：=]\s*$/u.test(prefix) && !previous)) {
    return { hits: [], unknown: true }
  }
  return { hits: [{ ...hit, metric: label, unit: hit.unit || contextUnit || unitIn(label) }], unknown: false }
}

function shapeName(shape) {
  const info = shape.children.find(child => child.name.startsWith('nv'))
  const props = one(info, 'cNvPr')
  return { shapeId: attr(props, 'id'), shapeName: attr(props, 'name'),
    placeholder: attr(one(one(info, 'nvPr'), 'ph'), 'type') }
}

function flatten(tree) {
  const objects = []
  for (const child of tree?.children ?? []) {
    if (child.name === 'grpSp') objects.push(...flatten(child))
    else if (!['nvGrpSpPr', 'grpSpPr', 'extLst'].includes(child.name)) objects.push(child)
  }
  return objects
}

function tableMetrics(table, period, add, skip) {
  const rows = all(table, 'tr').map(row => all(row, 'tc'))
  if (rows.some(row => row.length !== rows[0].length || row.some(cell =>
    ['gridSpan', 'rowSpan', 'hMerge', 'vMerge'].some(key => ![undefined, '0', '1'].includes(attr(cell, key))) ||
    ['hMerge', 'vMerge'].some(key => ['1', 'true'].includes(attr(cell, key)))))) {
    skip('mergedTable')
    return
  }
  const texts = rows.map(row => row.map(cell => paragraphs(one(cell, 'txBody')).trim()))
  if (texts.some(row => row.some(cell => cell.length > 1024))) {
    skip('unreadableNumber')
    return
  }
  for (const [row, cells] of texts.entries()) {
    for (const [col, text] of cells.entries()) {
      if (!text || !/\d/u.test(text)) continue
      const ownPeriods = slidePeriods(text, period)
      const { hits } = numberTokens(text)
      if (ownPeriods.length && (row === 0 || col === 0)) continue
      if (col === 0 || hits.length !== 1) {
        skip('unreadableNumber')
        continue
      }
      const num = hits[0]
      const labels = cells.slice(0, col).filter(label => label && !numberTokens(label).hits.length)
      const rowLabel = labels.filter(label => /[\p{L}]/u.test(label) && !unitIn(label)).at(-1)
      const colTitle = row > 0 ? texts[0][col] : ''
      const match = textMetrics(text, rowLabel, unitIn(colTitle) || labels.map(unitIn).find(Boolean))
      if (!match.hits.length || (num.unit && unitIn(colTitle) && num.unit !== unitIn(colTitle))) {
        skip('unreadableNumber')
        continue
      }
      add(match.hits[0], text, { row: row + 1, col: col + 1, rowLabel: labels.join(' / '), colTitle, periodText: colTitle })
    }
  }
}

export async function openSlides(bytes, runPeriod) {
  return withOfficePackage(bytes, async ({ main, relations, xml }) => {
    const deck = await xml(main)
    if (deck?.name !== 'presentation') throw problem('wrongFile')
    const ordered = all(one(deck, 'sldIdLst'), 'sldId')
    if (ordered.length > SLIDE_LIMIT) throw problem('slideLimit')
    const links = await relations(main)
    const numbers = []
    const coverage = { slides: ordered.length, textBoxes: 0, tables: 0, omitted: [] }
    let count = 0
    for (const [index, entry] of ordered.entries()) {
      const page = index + 1
      const part = linked(links, relId(entry), 'slide')
      const root = await xml(part)
      if (root?.name !== 'sld') throw problem('invalidOffice', part)
      const objects = flatten(one(one(root, 'cSld'), 'spTree'))
      count += objects.length
      if (count > OBJECT_LIMIT) throw problem('slideLimit', part)
      const boxes = objects.filter(shape => shape.name === 'sp' && one(shape, 'txBody'))
      const titleBox = boxes.find(shape => ['title', 'ctrTitle'].includes(shapeName(shape).placeholder))
        ?? boxes.find(shape => {
          const s = paragraphs(one(shape, 'txBody')).trim()
          return s && s.length <= 120 && !s.includes('\n') && !textMetrics(s).hits.length
        })
      const rawTitle = titleBox ? paragraphs(one(titleBox, 'txBody')).trim() : ''
      const title = rawTitle.length <= 240 ? rawTitle : ''
      let pageSection = ''
      for (const [indexInPage, shape] of objects.entries()) {
        const id = shapeName(shape)
        const spot = { page, title, ...id }
        const omittedKinds = new Set()
        const skip = kind => {
          if (omittedKinds.has(kind)) return
          omittedKinds.add(kind)
          coverage.omitted.push({ ...spot, kind })
        }
        const add = (num, context, table = {}, textSection = '') => {
          if (numbers.length >= NUMBER_LIMIT) throw problem('slideLimit', part)
          const explicit = [...new Set([title, pageSection, textSection, context, table.periodText ?? '']
            .flatMap(label => slidePeriods(label, runPeriod)))]
          const periodOkay = explicit.length === 0 || (explicit.length === 1 && explicit[0] === runPeriod)
          numbers.push({
            ...num, ...spot, ...table, context, section: [pageSection, textSection].filter(Boolean).join(' / '),
            explicit, periodOkay,
            id: [page, indexInPage, numbers.length].join('.'),
            semantic: JSON.stringify([slideSignature(title, runPeriod), slideSignature(pageSection, runPeriod),
              slideSignature(textSection, runPeriod), slideSignature(num.metric, runPeriod),
              signature(table.rowLabel ?? ''), slideSignature(table.colTitle ?? '', runPeriod), num.unit]),
          })
        }
        if (shape.name === 'sp' && one(shape, 'txBody')) {
          coverage.textBoxes++
          const text = paragraphs(one(shape, 'txBody'))
          if (text.length > TEXT_LIMIT) { skip('unreadableNumber'); continue }
          let section = ''
          let accepted = false
          for (const line of text.split('\n')) {
            if (!line.trim()) continue
            const parsed = textMetrics(line, section, unitIn(section))
            if (parsed.unknown) skip('unreadableNumber')
            for (const num of parsed.hits) {
              accepted = true
              add(num, section ? section + '\n' + line : line, {}, section)
            }
            if (!parsed.hits.length && !parsed.unknown) section = line.trim().slice(0, 512)
          }
          if (shape !== titleBox && !accepted && !omittedKinds.size && section.length <= 120) pageSection = section
        } else if (shape.name === 'graphicFrame') {
          const data = one(one(shape, 'graphic'), 'graphicData')
          const table = one(data, 'tbl')
          if (table) {
            coverage.tables++
            tableMetrics(table, runPeriod, add, skip)
          } else skip(one(data, 'chart') ? 'chart' : 'unsupportedObject')
        } else if (shape.name === 'pic') skip('image')
        else if (shape.name !== 'sp' && shape.name !== 'cxnSp') skip('unsupportedObject')
      }
    }
    return { numbers, coverage }
  })
}
