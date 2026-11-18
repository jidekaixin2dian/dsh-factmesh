import fs from 'node:fs/promises'
import path from 'node:path'
import { pathToFileURL } from 'node:url'

const modules = process.env.FACTMESH_ARTIFACT_MODULES
if (!modules) throw new Error('Set FACTMESH_ARTIFACT_MODULES to the bundled workspace node_modules directory')
const { Presentation, PresentationFile, Workbook, SpreadsheetFile } = await import(
  pathToFileURL(path.join(modules, '@oai/artifact-tool/dist/artifact_tool.mjs')).href
)
const samples = path.resolve(process.argv[2] ?? 'samples')
const renders = process.env.FACTMESH_SAMPLE_RENDERS
await fs.mkdir(samples, { recursive: true })
if (renders) await fs.mkdir(renders, { recursive: true })

const ink = '#172637'
const muted = '#526479'
const paper = '#F5F8FC'
const blue = '#23518E'
const font = 'Microsoft YaHei'
const periods = ['2026-09', '2026-10']
const metrics = [
  ['电商', '收入同比', '%', 0.128, 0.108],
  ['零售', '收入同比', '%', 0.108, 0.108],
  ['电商', '收入环比', '%', 0.108, 0.108],
  ['集团', '营业收入', '万元', 12000, 12000],
  ['集团', '毛利率', '%', 0.1284, 0.1284],
  ['零售', '净现金流', '万元', -4500, -5200],
  ['集团', '续费率', '%', 0.923, 0.917],
  ['集团', '应收比率', '%', null, null],
]

async function workbook(index) {
  const book = Workbook.create()
  const sheet = book.worksheets.add('经营指标')
  sheet.showGridLines = false
  const october = index === 1
  const current = october ? 'D' : 'F'
  const other = october ? 'F' : 'E'
  const first = october ? 18 : 17
  sheet.getRange('A2').values = [['云栖零售｜经营指标']]
  sheet.getRange('A2').format.font = { name: font, bold: true, size: 17, color: ink }
  sheet.getRange('A4').values = [['演示数据 · 按月更新']]
  sheet.getRange('A4').format.font = { name: font, color: muted }
  sheet.getRange('A15:C15').values = [['业务线', '指标', '单位']]
  sheet.getRange(current + '15').values = [[periods[index] + ' / 当期实际']]
  sheet.getRange(other + '15').values = [[(october ? '2026-09' : '2026-08') + ' / 当期实际']]
  const header = sheet.getRange('A15:F15')
  header.format.fill = blue
  header.format.font = { name: font, color: '#FFFFFF', bold: true }
  header.format.rowHeight = 24
  sheet.getRange('A:A').format.columnWidth = 16
  sheet.getRange('B:B').format.columnWidth = 27
  sheet.getRange('C:C').format.columnWidth = 12
  sheet.getRange('D:F').format.columnWidth = 24
  function writeRow(row, branch, metric, unit, currentValue, pastValue) {
    sheet.getRange('A' + row + ':C' + row).values = [[branch, metric, unit]]
    sheet.getRange(current + row).values = [[currentValue]]
    sheet.getRange(other + row).values = [[pastValue]]
    sheet.getRange('D' + row + ':F' + row).format.numberFormat = unit === '%' ? '0.00%' : '#,##0;(#,##0)'
    sheet.getRange('A' + row + ':F' + row).format.fill = row % 2 ? '#EAF0F8' : '#FFFFFF'
  }
  if (october) writeRow(17, '供应链', '物流费用同比', '%', 0.024, 0.031)
  for (const [offset, [branch, metric, unit, sep, oct]] of metrics.entries()) {
    const row = first + offset
    writeRow(row, branch, october && metric === '续费率' ? '续费率（新口径）' : metric, unit,
      october ? oct : sep, october ? sep : sep === null ? null : sep - (unit === '%' ? 0.006 : 200))
  }
  const formulaAddress = current + (first + metrics.length - 1)
  // A cached error deliberately requires source verification in the sample.
  sheet.getRange(formulaAddress).formulas = [['=1/0']]
  const notes = book.worksheets.add('口径说明')
  notes.getRange('A1:A3').values = [['月报口径'], ['金额使用万元，负数表示净流出。'], ['百分比单元格存储比率原值。']]
  notes.getRange('A:A').format.columnWidth = 48
  book.recalculate()
  const check = await book.inspect({ kind: 'table', range: '经营指标!A15:F' + (first + 7),
    include: 'values,formulas', tableMaxRows: 20, tableMaxCols: 8 })
  if (renders) {
    await fs.writeFile(path.join(renders, periods[index] + '-sheet.ndjson'), check.ndjson)
    const png = await book.render({ sheetName: '经营指标', range: 'A15:F' + (first + 7), scale: 1 })
    await fs.writeFile(path.join(renders, periods[index] + '-sheet.png'), new Uint8Array(await png.arrayBuffer()))
  }
  await (await SpreadsheetFile.exportXlsx(book)).save(path.join(samples, periods[index] + '.xlsx'))
}

function text(slide, name, value, x, y, w, h, options = {}) {
  const shape = slide.shapes.add({ name, geometry: 'textbox', fill: 'none',
    line: { fill: 'none', width: 0 },
    position: { left: x, top: y, width: w, height: h } })
  shape.text = value
  shape.text.style = { typeface: font, fontSize: options.size ?? 32, bold: Boolean(options.bold),
    color: options.color ?? ink, verticalAlignment: 'middle', wrap: 'none',
    insets: { top: 0, bottom: 0, left: 0, right: 0 } }
  return shape
}

async function powerpoint(index) {
  const month = periods[index]
  const october = index === 1
  const deck = Presentation.create({ slideSize: { width: 960, height: 540 } })
  function page(title, figures = [], note = '') {
    const slide = deck.slides.add()
    slide.background.fill = paper
    text(slide, '标题', title + '｜' + month, 66, 62, 828, 62, { size: 42, bold: true })
    for (const [row, [label, display]] of figures.entries()) {
      text(slide, label, label + '：' + display, 68, 170 + row * 80, 820, 69, { size: 42, bold: true, color: blue })
    }
    if (note) text(slide, '说明', note, 68, 468, 820, 35, { size: 20, color: muted })
    return slide
  }
  page('云栖零售经营月报', [], '经营分析组 · 虚构演示数据')
  if (october) page('供应链观察', [], '本期新增物流专题')
  const amount = () => page('经营规模', [['营业收入', '1.20 亿元'], ['净现金流', '-0.45 亿元']], '负数表示净流出')
  const core = () => page('电商核心指标', [['收入同比', '12.8%'], ['收入环比', '10.8%']])
  const table = () => {
    const slide = page('盈利能力')
    const element = slide.tables.add({ left: 68, top: 170, width: 824, height: 166, rows: 2, columns: 2,
      values: [['指标', '本期'], ['毛利率', '12.8%']] })
    for (let row = 0; row < 2; row++) {
      for (let col = 0; col < 2; col++) {
        const cell = element.getCell(row, col)
        cell.fill = row === 0 ? blue : '#FFFFFF'
        cell.text.style = { typeface: font, fontSize: 26, bold: true, color: row === 0 ? '#FFFFFF' : ink }
      }
    }
  }
  const memberships = () => page('会员与回款', [['续费率', october ? '91.7%' : '92.3%'], ['应收比率', '5.0%']])
  const chart = () => {
    const slide = page('订单渠道', [], '渠道订单量 · 单位为件')
    slide.charts.add('bar', {
      position: { left: 70, top: 153, width: 795, height: 299 },
      categories: ['线上', '门店', '合作方'],
      series: [{ name: '订单量', values: october ? [135, 81, 44] : [127, 79, 45], fill: blue }],
      hasLegend: false,
      dataLabels: { showValue: true, position: 'outEnd' },
    })
  }
  if (october) {
    core(); amount(); table(); memberships(); chart()
  } else {
    amount(); table(); memberships(); chart(); core()
  }
  if (renders) {
    for (const [i, slide] of deck.slides.items.entries()) {
      const png = await deck.export({ slide, format: 'png', scale: 0.75 })
      await fs.writeFile(path.join(renders, month + '-slide-' + (i + 1) + '.png'), new Uint8Array(await png.arrayBuffer()))
    }
  }
  await (await PresentationFile.exportPptx(deck)).save(path.join(samples, month + '.pptx'))
}

for (const index of periods.keys()) {
  await workbook(index)
  await powerpoint(index)
  for (const extension of ['xlsx', 'pptx']) {
    await fs.rm(path.join(samples, periods[index] + '.' + extension + '.inspect.ndjson'), { force: true })
  }
}
