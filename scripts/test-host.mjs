import assert from 'node:assert/strict'
import { spawn, execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { chromium } from 'playwright'

const root = path.resolve(import.meta.dirname, '..')
const output = path.resolve(process.env.FACTMESH_HOST_RESULTS ?? path.join(root, 'work/host-results'))
const tarball = path.resolve(process.argv[2] ?? path.join(root, 'dist/jidekaixin2dian-dsh-factmesh-0.1.0.tgz'))
const installSource = process.env.FACTMESH_INSTALL_SOURCE ?? tarball
const cli = process.env.FACTMESH_DSH_CLI || (process.platform === 'win32'
  ? path.join(process.env.APPDATA, 'npm/node_modules/@deepseek-ai/dsh/lib/bin.js')
  : null)
assert.ok(cli, 'Set FACTMESH_DSH_CLI to the installed DSH CLI bin.js')
await fs.mkdir(output, { recursive: true })
const isolation = await fs.mkdtemp(path.join(os.tmpdir(), 'factmesh-host-'))
const home = path.join(isolation, 'home')
const workspace = path.join(isolation, 'workspace')
await fs.mkdir(workspace)
const documents = path.join(isolation, 'Documents')
await fs.mkdir(documents)
const overlay = path.join(isolation, 'web-e2e.yml')
await fs.writeFile(overlay, [
  '- id: workspace-controller',
  '  config:',
  '    documentsDirectory: ' + JSON.stringify(documents),
  '- id: directory-picker',
  '  disabled: true',
  '- insert:',
  '    - id: directory-picker-browse',
  "      name: '@deepseek-ai/dsh-host-directory-picker-browse'",
  '    - id: ui-directory-picker-browse',
  "      name: '@deepseek-ai/dsh-client-ui-directory-picker-browse'",
  '',
].join('\n'))
const env = { ...process.env, DSH_HOME: home, DSH_AGENTS_HOME: path.join(home, 'agents'),
  DSH_TELEMETRY_DISABLED: '1', DEEPSEEK_API_KEY: '' }
const version = execFileSync(process.execPath, [cli, '--version'], { env, encoding: 'utf8' }).trim()
const files = (await fs.readdir(path.join(root, 'samples'))).filter(name => /\.(xlsx|pptx)$/u.test(name))
assert.equal(files.length, 4)
const sha256 = async filename => createHash('sha256').update(await fs.readFile(filename)).digest('hex')
const hashes = async directory => Object.fromEntries(await Promise.all(files.map(
  async file => [file, await sha256(path.join(directory, file))],
)))
for (const file of files) await fs.copyFile(path.join(root, 'samples', file), path.join(workspace, file))
const before = await hashes(workspace)
const install = execFileSync(process.execPath, [cli, 'plugin', '--profile', 'web', 'add', installSource],
  { cwd: workspace, env, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] })
await fs.writeFile(path.join(output, 'install.log'), install)
const manifest = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'))
const installed = path.join(home, 'profiles/web/node_modules', manifest.name)
assert.equal(JSON.parse(await fs.readFile(path.join(installed, 'package.json'), 'utf8')).version, manifest.version)
const clientSha256 = await sha256(path.join(installed, 'lib/client.js'))
assert.equal(clientSha256, await sha256(path.join(root, 'lib/client.js')))
const host = spawn(process.execPath, [cli, 'web', '--patch', overlay, '--host', '127.0.0.1', '--port', '0', '--no-open'],
  { cwd: workspace, env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
let hostLog = ''
const hostReady = new Promise((resolve, reject) => {
  const timer = setTimeout(() => reject(new Error('DSH did not open the Web profile')), 40_000)
  for (const pipe of [host.stdout, host.stderr]) pipe.on('data', chunk => {
    hostLog += chunk.toString()
    const url = /dsh web: (http:\/\/127\.0\.0\.1:\d+\/\?token=[\w-]+)/u.exec(hostLog)?.[1]
    if (url) { clearTimeout(timer); resolve(url) }
  })
  host.once('error', error => { clearTimeout(timer); reject(error) })
  host.once('exit', code => { clearTimeout(timer); reject(new Error('DSH exited: ' + code)) })
})

let browser
const pageErrors = []
try {
  const url = await hostReady
  browser = await chromium.launch({ headless: true, channel: process.env.FACTMESH_BROWSER_CHANNEL })
  const locale = process.env.FACTMESH_BROWSER_LOCALE ?? 'zh-CN'
  const colorScheme = process.env.FACTMESH_COLOR_SCHEME ?? 'light'
  const context = await browser.newContext({ locale, colorScheme, viewport: { width: 1440, height: 1050 } })
  const page = await context.newPage()
  page.setDefaultTimeout(15_000)
  page.on('pageerror', error => pageErrors.push(String(error)))
  await page.goto(url.replace(/[\r\n]/gu, ''))
  const panel = page.locator('[data-factmesh]')
  const nav = page.getByRole('button', { name: /FactMesh/ })
  const add = page.getByRole('button', { name: /添加工作区|Add workspace/iu })
  await nav.waitFor()
  const welcome = page.getByRole('button', { name: /^(继续|Continue)$/u })
  await welcome.click()
  async function skipProviderSetup() {
    await page.getByRole('button', { name: /^(添加 API Key|Add API Key)$/iu }).click()
    await page.getByRole('button', { name: /稍后配置|Set up later|Configure later/iu }).click()
  }
  await skipProviderSetup()
  if (await add.count()) {
    await add.first().click()
    const edit = page.getByRole('button', { name: /编辑路径|Edit path/iu })
    await edit.first().click()
    const entry = page.getByRole('textbox', { name: /编辑路径|Edit path/iu })
    await entry.fill(workspace)
    await entry.press('Enter')
    await entry.waitFor({ state: 'hidden' })
    const choose = page.getByRole('dialog').getByRole('button', { name: /^(打开|Open)$/u }).last()
    await choose.click()
  }
  await nav.click()
  await panel.getByRole('heading', { name: 'FactMesh' }).waitFor()
  async function selectWorkspace() {
    const workspaceSelect = panel.getByLabel(/工作区会话|Workspace session/u)
    const scopeOption = workspaceSelect.locator('option').filter({ hasText: workspace }).first()
    await scopeOption.waitFor({ state: 'attached' })
    await workspaceSelect.selectOption(await scopeOption.getAttribute('value'))
  }
  await selectWorkspace()

  async function run(month, readSheets = true) {
    await panel.getByLabel(/PowerPoint (汇报|report)/iu).selectOption(month + '.pptx')
    await panel.getByLabel(/Excel (底表|workbook)/iu).selectOption(month + '.xlsx')
    await panel.getByLabel(/核对期间|Reporting month/u).fill(month)
    if (readSheets) await panel.getByRole('button', { name: /读取工作表目录|Read worksheet list/u }).click()
    await panel.getByLabel(/^(工作表|Worksheet)$/u).selectOption('经营指标')
    await panel.getByRole('button', { name: /^(开始核对|Reconcile)$/u }).click()
    await panel.locator('[data-results] .fm-row').first().waitFor()
  }
  const row = metric => panel.locator('.fm-row[data-metric="' + metric + '"]')
  const status = async (metric, value) => assert.equal(await row(metric).getAttribute('data-status'), value, metric)
  const candidates = panel.getByLabel(/候选单元格|Candidate cells/u)
  const confirmation = panel.getByRole('button', { name: /^(确认关系|Confirm link)$/u })
  async function confirm(metric, address, expected = 'match') {
    await row(metric).click()
    await candidates.selectOption(address)
    await confirmation.click()
    await row(metric).and(panel.locator('[data-status="' + expected + '"]')).waitFor()
  }

  await run('2026-09')
  assert.equal(await panel.locator('.fm-row').count(), 7)
  assert.equal(await panel.locator('[data-status="difference"]').count(), 0)
  await row('收入同比').click()
  assert.equal(await candidates.inputValue(), '')
  const names = await candidates.locator('option').allTextContents()
  assert.ok(names.some(name => /F17.+电商/u.test(name)))
  assert.ok(names.some(name => /F18.+零售/u.test(name)))
  assert.ok(names.every(name => !name.includes('环比')))
  await confirm('收入同比', 'F17')
  const switchCandidate = async (address, expected) => {
    await candidates.selectOption(address)
    await panel.getByRole('button', { name: /^(更换为此候选|Switch to this cell)$/u }).click()
    await row('收入同比').and(panel.locator('[data-status="' + expected + '"]')).waitFor()
  }
  await switchCandidate('F18', 'difference')
  await switchCandidate('F17', 'match')
  await confirm('收入环比', 'F19')
  await confirm('营业收入', 'F20')
  assert.match(await panel.locator('[data-comparison]').innerText(), /12000.+10000.+100000000.+1\.2亿元/u)
  await confirm('毛利率', 'F21')
  assert.match(await panel.locator('[data-comparison]').innerText(), /12\.84%/u)
  await confirm('净现金流', 'F22')
  await confirm('续费率', 'F23')
  await confirm('应收比率', 'F24', 'verify')
  assert.match(await panel.locator('[data-cell]').innerText(), /#DIV\/0!/u)
  await panel.locator('.fm-coverage summary').first().click()
  assert.match(await panel.locator('.fm-omitted').innerText(), /嵌入图表|Embedded chart/u)
  await panel.locator('.fm-coverage summary').first().click()
  assert.equal(await panel.locator('[data-status="match"]').count(), 6)
  assert.deepEqual(await hashes(workspace), before)
  await page.screenshot({ path: path.join(output, '01-confirmed.png') })

  // A page reload verifies durable mappings before the files for month two are chosen.
  await page.reload()
  await skipProviderSetup()
  await nav.click()
  await selectWorkspace()
  await run('2026-10')
  await row('收入同比').click()
  assert.match(await row('收入同比').innerText(), /第 3 页|Slide 3/u)
  assert.equal(await panel.locator('[data-cell]').getAttribute('data-cell'), '经营指标!D18')
  assert.match(await panel.locator('[data-comparison]').innerText(), /2\.0.*(?:百分点|percentage points)/u)
  await status('收入同比', 'difference')
  await status('净现金流', 'difference')
  await status('续费率', 'reconfirm')
  assert.match(await panel.locator('[data-reused]').innerText(), /\b6\b/u)
  assert.equal(await panel.locator('[data-status="difference"]').count(), 2)
  assert.equal(await panel.locator('[data-status="match"]').count(), 3)
  await page.screenshot({ path: path.join(output, '02-rerun.png') })
  await page.getByRole('button', { name: /^(收起侧边栏|Collapse sidebar)$/iu }).click()
  await page.setViewportSize({ width: 620, height: 915 })
  await page.locator('[data-shortcut-modal="sidebar"][aria-modal="true"]').waitFor({ state: 'hidden' })
  await panel.locator('.fm-split').evaluate(element => {
    if (getComputedStyle(element).gridTemplateColumns.trim().split(/\s+/u).length !== 1) {
      throw new Error('The review layout did not stack at the narrow viewport')
    }
  })
  await page.screenshot({ path: path.join(output, '03-narrow.png') })
  await row('续费率').click()
  await panel.getByLabel(/手选单元格地址|Cell address/u).fill('D24')
  await panel.getByRole('button', { name: /^(查看|Inspect)$/u }).click()
  assert.equal(await panel.locator('[data-cell]').getAttribute('data-cell'), '经营指标!D24')
  await confirmation.click()
  await row('续费率').and(panel.locator('[data-status="match"]')).waitFor()
  await row('收入环比').click()
  await panel.getByRole('button', { name: /^(跳过此项|Skip item)$/u }).click()
  await status('收入环比', 'skipped')
  await panel.getByRole('button', { name: /^(取消跳过|Undo skip)$/u }).click()
  await status('收入环比', 'match')
  assert.deepEqual(await hashes(workspace), before)

  // Simulate an external replacement only on the isolated copy.
  await fs.copyFile(path.join(workspace, '2026-09.xlsx'), path.join(workspace, '2026-10.xlsx'))
  await panel.locator('[data-stale="true"]').waitFor()
  assert.match(await panel.getByRole('alert').innerText(), /已过期|out of date/u)
  assert.equal(await panel.locator('[data-comparison]').count(), 0)
  assert.deepEqual(await hashes(path.join(root, 'samples')), before)
  assert.deepEqual(pageErrors, [])
  await fs.writeFile(path.join(output, 'receipt.json'), JSON.stringify({
    host: version, platform: process.platform, locale, colorScheme, installSource,
    packageSha256: await sha256(tarball), clientSha256, files: before,
    september: { slides: 6, figures: 7, matches: 6, discrepancies: 0, verify: 1, omittedCharts: 1 },
    october: { slides: 7, figures: 7, plantedDiscrepancies: 2, recovered: 2, falsePositives: 0,
      reused: 6, manualRebinds: 1, formulaVerification: 1, omittedCharts: 1 },
    replacedSourceDetected: true, officeHashesUnchangedDuringReview: true,
  }, null, 2) + '\n')
  console.log('FactMesh verified through DSH Web ' + version)
} catch (error) {
  const page = browser?.contexts()[0]?.pages()[0]
  if (page) {
    await page.screenshot({ path: path.join(output, 'failure.png') }).catch(() => {})
    const dom = await page.locator('body').innerText().catch(() => '')
    await fs.writeFile(path.join(output, 'failure.txt'), [String(error), ...pageErrors, dom].join('\n\n'))
  }
  throw error
} finally {
  await browser?.close()
  host.kill()
  await fs.writeFile(path.join(output, 'host.log'), hostLog.replace(/([?&]token=)[\w-]+/gu, '$1<redacted>'))
}
