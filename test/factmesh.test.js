import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'
import { TextReader, Uint8ArrayWriter, ZipWriter } from '@zip.js/zip.js'
import { compareNumbers, numberTokens } from '../src/number.js'
import { openSlides } from '../src/slides.js'
import { openWorkbook } from '../src/workbook.js'
import { candidateIndex, candidatesFor, createProjectStore, makeLink, reviewRun } from '../src/reconcile.js'
import { readSource } from '../src/workspace.js'

const sheetName = '经营指标'
async function sample(month) {
  const [pptx, xlsx] = await Promise.all(['pptx', 'xlsx'].map(ext =>
    readFile(new URL('../samples/' + month + '.' + ext, import.meta.url)),
  ))
  return {
    period: month,
    files: Object.fromEntries(['pptx', 'xlsx'].map(ext => [ext, { absolutePath: '/sample/' + month + '.' + ext, version: 'first' }])),
    slides: await openSlides(pptx, month),
    sheet: (await openWorkbook(xlsx, sheetName)).sheet,
  }
}

test('real Office samples relocate independently of slide order, old cell and matching values', async () => {
  const [sept, oct] = await Promise.all(['2026-09', '2026-10'].map(sample))
  assert.equal(sept.slides.coverage.tables, 1)
  assert.deepEqual(sept.slides.coverage.omitted.map(x => x.kind), ['chart'])
  assert.equal(sept.slides.numbers.length, 7)
  assert.ok(reviewRun(sept, []).rows.every(row => row.status === 'confirm'))
  const index = candidateIndex(sept.sheet.cells)
  const revenue = sept.slides.numbers.find(x => x.metric === '收入同比')
  const candidates = candidatesFor(revenue, index, sept.period)
  assert.deepEqual(candidates.filter(x => x.period === sept.period).map(x => x.address), ['F17', 'F18'])
  const byMetric = { 收入同比: 'F17', 收入环比: 'F19', 营业收入: 'F20',
    毛利率: 'F21', 净现金流: 'F22', 续费率: 'F23', 应收比率: 'F24' }
  const links = sept.slides.numbers.map(slide => makeLink(sept, slide, index.byAddress.get(byMetric[slide.metric])))
  assert.equal(reviewRun(sept, links).rows.filter(row => row.status === 'match').length, 6)
  assert.equal(oct.sheet.cells.find(cell => cell.address === 'F17').metric, '物流费用同比')
  const next = new Map(reviewRun(oct, links).rows.map(row => [row.slide.metric, row]))
  assert.deepEqual([...next.values()].filter(row => row.status === 'difference').map(row => row.slide.metric),
    ['收入同比', '净现金流'])
  const yoy = next.get('收入同比')
  assert.deepEqual([yoy.slide.page, yoy.target.address, yoy.target.raw, yoy.delta, yoy.differenceUnit],
    [3, 'D18', '0.108', '2.0', '百分点'])
  assert.equal(next.get('营业收入').workbook, '1.2')
  assert.equal(next.get('毛利率').workbook, '12.84')
  assert.equal(next.get('毛利率').rounded, '12.8')
  assert.equal(next.get('应收比率').reason, 'unverifiedFormula')
  assert.equal(next.get('续费率').status, 'reconfirm')
  assert.equal([...next.values()].filter(row => row.reused).length, 6)
})

test('percentage points, currency factors and signed half-up rounding use exact decimals', () => {
  const cell = (raw, unit, factor = '1') => ({ raw, unit, factor })
  const slide = (raw, unit, precision) => ({ raw, unit, precision })
  assert.equal(compareNumbers(slide('12.8', '%', 1), cell('12.8', '百分点')).reason, 'unitMismatch')
  assert.equal(compareNumbers(slide('2', '百分点', 0), cell('2', '百分点')).status, 'match')
  assert.equal(compareNumbers(slide('-1.3', '亿元', 1), cell('-12500', '万元')).status, 'match')
  assert.equal(compareNumbers(slide('12.8', '%', 1), cell('0.1285', '%', '100')).status, 'difference')
  assert.equal(numberTokens('现金流（9,000 万元）').hits[0].raw, '-9000')
})

const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const packageNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
const s = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const rels = (...items) => '<Relationships xmlns="' + packageNs + '">' + items.map(([id, type, target]) =>
  '<Relationship Id="' + id + '" Type="' + r + '/' + type + '" Target="' + target + '"/>').join('') + '</Relationships>'
async function zip(files) {
  const writer = new ZipWriter(new Uint8ArrayWriter(), { useWebWorkers: false })
  for (const [part, content] of Object.entries(files)) await writer.add(part, new TextReader(content))
  return writer.close()
}

async function workbookVariant({ month, upper, section, number = '-25', format }) {
  const inline = (addr, value) => '<c r="' + addr + '" t="inlineStr"><is><t>' + value + '</t></is></c>'
  return (await openWorkbook(await zip({
    '_rels/.rels': rels(['office', 'officeDocument', 'xl/workbook.xml']),
    'xl/workbook.xml': '<workbook xmlns="' + s + '" xmlns:r="' + r + '"><sheets><sheet name="Demo" r:id="sheet"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': rels(['sheet', 'worksheet', 'worksheets/grid.xml'], ['styles', 'styles', 'styles.xml']),
    'xl/styles.xml': '<styleSheet xmlns="' + s + '"><numFmts><numFmt numFmtId="164" formatCode="' + format +
      '"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="164"/></cellXfs></styleSheet>',
    'xl/worksheets/grid.xml': '<worksheet xmlns="' + s + '"><sheetData><row r="1">' + inline('B1', upper) +
      '</row><row r="2">' + inline('B2', month) + '</row><row r="3">' + inline('A3', section) +
      '</row><row r="4">' + inline('A4', '收入') + '<c r="B4" s="1"><v>' + number +
      '</v></c></row></sheetData></worksheet>',
  }), 'Demo')).sheet.cells[0]
}

test('worksheet identity retains business sections and headings above the month', async () => {
  const format = '0&quot;元&quot;;(0&quot;万元&quot;)'
  const first = await workbookVariant({ month: '2026-09', upper: '实际', section: '电商', format })
  const next = await workbookVariant({ month: '2026-10', upper: '实际', section: '电商', format })
  const budget = await workbookVariant({ month: '2026-10', upper: '预算', section: '电商', format })
  const retail = await workbookVariant({ month: '2026-10', upper: '实际', section: '零售', format })
  assert.equal(first.semantic, next.semantic)
  assert.notEqual(first.semantic, budget.semantic)
  assert.notEqual(first.semantic, retail.semantic)
  assert.equal(first.unit, '万元')
  const conditional = await workbookVariant({ month: '2026-10', upper: '实际', section: '零售', format: '[&gt;0]0%;0' })
  assert.equal(conditional.reason, 'unknownFormat')
})

test('ZIP parts reject XML declarations with DTDs and external relationships cannot select a deck', async () => {
  await assert.rejects(openSlides(await zip({
    '_rels/.rels': '<!DOCTYPE Relationships>' + rels(['office', 'officeDocument', 'ppt/presentation.xml']),
  }), '2026-09'), { code: 'unsupportedXml' })
  await assert.rejects(openSlides(await zip({
    '_rels/.rels': '<Relationships xmlns="' + packageNs + '"><Relationship Id="r1" Type="' + r +
      '/officeDocument" Target="https://example.invalid/a" TargetMode="External"/></Relationships>',
  }), '2026-09'), { code: 'invalidOffice' })
})

test('a slide that still says 9月 cannot be compared against October', async () => {
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  const shape = text => '<p:sp><p:txBody><a:p><a:r><a:t>' + text + '</a:t></a:r></a:p></p:txBody></p:sp>'
  const data = await zip({
    '_rels/.rels': rels(['office', 'officeDocument', 'ppt/presentation.xml']),
    'ppt/presentation.xml': '<p:presentation xmlns:p="' + p + '" xmlns:r="' + r + '">' +
      '<p:sldIdLst><p:sldId r:id="first"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': rels(['first', 'slide', 'slides/slide1.xml']),
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="' + p + '" xmlns:a="' + a + '"><p:cSld><p:spTree>' +
      shape('9月经营情况') + shape('收入同比：10.8%') + '</p:spTree></p:cSld></p:sld>',
  })
  assert.equal((await openSlides(data, '2026-09')).numbers[0].periodOkay, true)
  const october = await openSlides(data, '2026-10')
  assert.equal(october.numbers[0].periodOkay, false)
  assert.deepEqual(october.numbers[0].explicit, ['2026-09'])
})

test('a newly unique link can be reused despite older links that needed position in their snapshot', async () => {
  const first = await sample('2026-09')
  const revenue = first.slides.numbers.find(x => x.metric === '收入同比')
  const copy = { ...revenue, id: revenue.id + '-duplicate' }
  const duplicates = { ...first, slides: { numbers: [revenue, copy] } }
  const cell = first.sheet.cells.find(item => item.address === 'F17')
  const store = createProjectStore({
    getItem: key => storage.get(key),
    setItem: (key, value) => storage.set(key, value),
  }, '/samples')
  const storage = new Map()
  store.save(makeLink(duplicates, revenue, cell))
  store.save(makeLink(duplicates, copy, cell))
  const next = await sample('2026-10')
  const newSlide = next.slides.numbers.find(x => x.metric === '收入同比')
  const newCell = next.sheet.cells.find(item => item.address === 'D18')
  assert.equal(reviewRun(next, store.read()).rows.find(row => row.slide === newSlide).status, 'reconfirm')
  store.save(makeLink(next, newSlide, newCell))
  const replacement = { ...next, files: { ...next.files, xlsx: { ...next.files.xlsx, version: 'new' } } }
  assert.equal(reviewRun(replacement, store.read()).rows.find(row => row.slide === newSlide).reused, true)
})

test('workspace reads recheck selected path, byte window, and the post-read version', async () => {
  const ok = value => ({ ok: true, value })
  const file = { absolutePath: 'C:\\workspace\\report.xlsx', version: 'v1', bytes: 3 }
  const signal = new AbortController().signal
  let version = 'v1'
  let reads = 0
  const remote = {
    list: async () => ok({ entries: [{ name: 'report.xlsx', type: 'file' }] }),
    stat: async () => ok({ ...file, version }),
    readBytes: async () => {
      reads++
      return ok({ ...file, offset: 0, eof: true, data: Uint8Array.of(1, 2, 3) })
    },
  }
  assert.deepEqual((await readSource(remote, 's', 'C:\\workspace', 'report.xlsx', signal)).data, Uint8Array.of(1, 2, 3))
  version = 'v2'
  await assert.rejects(readSource(remote, 's', 'C:\\workspace', 'report.xlsx', signal), { code: 'stale' })
  const before = reads
  await assert.rejects(readSource(remote, 's', 'C:\\outside', 'report.xlsx', signal), { code: 'workspaceError' })
  await assert.rejects(readSource(remote, 's', 'C:\\workspace', '../report.xlsx', signal), { code: 'workspaceError' })
  assert.equal(reads, before)
})
