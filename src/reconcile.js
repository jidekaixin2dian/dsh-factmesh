import { signature } from './meaning.js'
import { compareNumbers, UNIT } from './number.js'

const LIMIT = 2000

export function createProjectStore(storage, scope) {
  const key = 'dsh-factmesh:v1:' + JSON.stringify(scope)
  function read() {
    const serialized = storage.getItem(key)
    if (!serialized) return []
    const { schema, links } = JSON.parse(serialized)
    if (schema !== 1 || !Array.isArray(links) || links.length > LIMIT ||
        links.some(link => !link || typeof link.id !== 'string' || !link.source || !link.target || !link.origin)) {
      throw new Error('Invalid FactMesh project')
    }
    return links
  }
  return {
    key, read,
    save(link, replacedId) {
      const links = read().filter(item => item.id !== replacedId &&
        !(item.target.sheet === link.target.sheet && item.source.semantic === link.source.semantic && item.reusable && link.reusable) &&
        !(item.origin.deck === link.origin.deck && item.source.id === link.source.id &&
          item.origin.period === link.origin.period && item.target.sheet === link.target.sheet))
      if (links.length >= LIMIT) throw new Error('FactMesh project limit')
      links.push(link)
      storage.setItem(key, JSON.stringify({ schema: 1, links }))
      return links
    },
  }
}

function fileIdentity(file) {
  return JSON.stringify([file.absolutePath, file.version])
}

function compatible(slide, cell) {
  return UNIT[slide.unit]?.[0] === UNIT[cell.unit]?.[0]
}

export function candidateIndex(cells) {
  const byMetric = new Map()
  const byAddress = new Map()
  for (const cell of cells) {
    const key = signature(cell.metric)
    if (!byMetric.has(key)) byMetric.set(key, [])
    byMetric.get(key).push(cell)
    byAddress.set(cell.address, cell)
  }
  return { byMetric, byAddress }
}

export function candidatesFor(slide, index, period) {
  const label = signature(slide.metric)
  const options = new Set(index.byMetric.get(label) ?? [])
  if (!options.size) {
    for (const [key, cells] of index.byMetric) {
      if (key.length >= 2 && (label.endsWith(' ' + key) || label.endsWith(key) && /[^\x00-\x7f]/u.test(key))) {
        for (const cell of cells) options.add(cell)
      }
    }
  }
  return [...options].sort((a, b) => {
    const score = cell => (cell.period === period ? 4 : 0) + (compatible(slide, cell) ? 2 : 0) +
      (signature(cell.metric) === label ? 1 : 0)
    return score(b) - score(a)
  })
}

export function confirmationProblem(slide, cell, period) {
  if (!slide.periodOkay) return 'slidePeriod'
  if (cell.period !== period) return 'cellPeriod'
  if (!cell.metric || !cell.colTitle) return 'headingsMissing'
  if (!compatible(slide, cell) && !cell.reason) return 'unitMismatch'
  return ''
}

function groups(items, key) {
  const index = new Map()
  for (const item of items) {
    const value = key(item)
    const group = index.get(value)
    if (group) group.push(item)
    else index.set(value, [item])
  }
  return index
}

export function reviewRun(run, links) {
  const bySlide = groups(run.slides.numbers, slide => slide.semantic)
  const current = run.sheet.cells.filter(cell => cell.period === run.period)
  const byCell = groups(current, cell => cell.semantic)
  const deck = fileIdentity(run.files.pptx)
  const book = fileIdentity(run.files.xlsx)
  const applicable = links.filter(link => link.target.sheet === run.sheet.name)
  const bySource = groups(applicable.filter(link => link.reusable), link => link.source.semantic)
  const rows = run.slides.numbers.map(slide => ({
    slide, status: slide.periodOkay ? 'confirm' : 'verify',
    reason: slide.periodOkay ? '' : 'slidePeriod', linkIds: [],
  }))
  const byId = new Map(rows.map(row => [row.slide.id, row]))
  const missing = []
  for (const link of applicable) {
    const sources = bySlide.get(link.source.semantic) ?? []
    const targets = byCell.get(link.target.semantic) ?? []
    const sameSnapshot = link.origin.deck === deck && link.origin.book === book && link.origin.period === run.period
    const source = sameSnapshot ? sources.find(slide => slide.id === link.source.id)
      : sources.length === 1 && bySource.get(link.source.semantic)?.length === 1 ? sources[0] : undefined
    const target = sameSnapshot ? targets.find(cell => cell.address === link.target.address)
      : targets.length === 1 ? targets[0] : undefined
    if (!sources.length) {
      missing.push(link)
      continue
    }
    if (!source) {
      for (const slide of sources) byId.get(slide.id).linkIds.push(link.id)
      continue
    }
    const row = byId.get(source.id)
    row.linkIds.push(link.id)
    if (!target || (!sameSnapshot && !link.reusable) || confirmationProblem(source, target, run.period)) continue
    const result = compareNumbers(source, target)
    Object.assign(row, { ...result, target, link, reused: !sameSnapshot })
  }
  for (const row of rows) {
    if (!row.target && row.linkIds.length && row.status === 'confirm') row.status = 'reconfirm'
  }
  return { rows, missing }
}

export function makeLink(run, slide, cell) {
  const issue = confirmationProblem(slide, cell, run.period)
  if (issue) throw Object.assign(new Error(issue), { code: issue })
  const targetMatches = run.sheet.cells.filter(other =>
    other.period === run.period && other.semantic === cell.semantic).length
  const sourceMatches = run.slides.numbers.filter(other => other.semantic === slide.semantic).length
  return {
    id: crypto.randomUUID(),
    source: { id: slide.id, semantic: slide.semantic, title: slide.title, metric: slide.metric,
      page: slide.page, unit: slide.unit },
    target: { semantic: cell.semantic, sheet: cell.sheet, address: cell.address, rowTitle: cell.rowTitle,
      colTitle: cell.colTitle, unit: cell.unit },
    origin: { deck: fileIdentity(run.files.pptx), book: fileIdentity(run.files.xlsx),
      period: run.period, periodRule: 'selected-month' },
    reusable: cell.reusable && targetMatches === 1 && sourceMatches === 1,
  }
}
