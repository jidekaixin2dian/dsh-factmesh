import assert from 'node:assert/strict'
import { execFile, execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import fs from 'node:fs/promises'
import http from 'node:http'
import os from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { promisify } from 'node:util'

const root = path.resolve(import.meta.dirname, '..')
const market = process.env.FACTMESH_MARKET_SOURCE
assert.ok(market, 'Set FACTMESH_MARKET_SOURCE to the dshmarket 1.73.20 source checkout')
const npm = process.env.npm_execpath
assert.ok(npm, 'Run this with npm run test:update')
const cli = process.env.FACTMESH_DSH_CLI || (process.platform === 'win32'
  ? path.join(process.env.APPDATA, 'npm/node_modules/@deepseek-ai/dsh/lib/bin.js') : null)
assert.ok(cli, 'Set FACTMESH_DSH_CLI to the installed DSH CLI bin.js')
const output = path.resolve(process.env.FACTMESH_UPDATE_RESULTS ?? path.join(root, 'work/update-results'))
await fs.mkdir(output, { recursive: true })
const current = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'))
const marketPackage = JSON.parse(await fs.readFile(path.join(market, 'package.json'), 'utf8'))
assert.equal(marketPackage.version, '1.73.20')
const prior = { ...current, version: '0.0.9' }
const old = prior.version
const target = current.version
assert.notEqual(old, target)
const approved = new Map([[old, prior], [target, current]])
const temp = await fs.mkdtemp(path.join(os.tmpdir(), 'factmesh-update-'))
const fixture = path.join(temp, 'old')
const home = path.join(temp, 'home')
const profile = path.join(home, 'profiles/web')
const npmrc = path.join(temp, '.npmrc')
await fs.mkdir(fixture)
await fs.writeFile(npmrc, '')
await fs.cp(path.join(root, 'lib'), path.join(fixture, 'lib'), { recursive: true })
await fs.copyFile(path.join(root, 'cordis.patch.yml'), path.join(fixture, 'cordis.patch.yml'))
await fs.writeFile(path.join(fixture, 'package.json'), JSON.stringify(prior))
const npmPack = args => execFileSync(process.execPath, [npm, 'pack', '--silent', ...args], { cwd: root, encoding: 'utf8' }).trim()
const oldName = npmPack([fixture, '--pack-destination', temp])
const newName = npmPack(['--pack-destination', temp])
const artifacts = new Map([[old, await fs.readFile(path.join(temp, oldName))],
  [target, await fs.readFile(path.join(temp, newName))]])
const hash = (bytes, algorithm, encoding) => createHash(algorithm).update(bytes).digest(encoding)
const integrity = version => 'sha512-' + hash(artifacts.get(version), 'sha512', 'base64')
let latest = old
let base
const registryLog = []
function manifest(version) {
  return { ...approved.get(version), dist: {
    tarball: base + '/tarballs/' + version + '.tgz',
    integrity: integrity(version),
    shasum: hash(artifacts.get(version), 'sha1', 'hex'),
  } }
}
const server = http.createServer((request, response) => {
  const route = decodeURIComponent(new URL(request.url, base).pathname)
  registryLog.push(route)
  const archive = /^\/tarballs\/([\d.]+)\.tgz$/u.exec(route)?.[1]
  if (archive && artifacts.has(archive)) {
    response.setHeader('content-type', 'application/octet-stream')
    response.end(artifacts.get(archive))
    return
  }
  const endpoint = '/' + current.name
  response.setHeader('content-type', 'application/json')
  if (route === endpoint || route === endpoint + '/') {
    response.end(JSON.stringify({ name: current.name, 'dist-tags': { latest },
      versions: Object.fromEntries([...approved.keys()].filter(v => v === old || latest === target).map(v => [v, manifest(v)])),
      time: { [old]: '2026-01-01T00:00:00.000Z', ...(latest === target ? { [target]: '2026-01-02T00:00:00.000Z' } : {}) },
    }))
  } else if (route === endpoint + '/latest') {
    response.end(JSON.stringify(manifest(latest)))
  } else {
    response.writeHead(404)
    response.end('{}')
  }
})
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve))
base = 'http://127.0.0.1:' + server.address().port

const cliEnv = { ...process.env, DSH_HOME: home, DSH_AGENTS_HOME: path.join(home, 'agents'),
  DSH_TELEMETRY_DISABLED: '1', DEEPSEEK_API_KEY: '', npm_config_userconfig: npmrc, npm_config_registry: base,
  NO_PROXY: '127.0.0.1,localhost', no_proxy: '127.0.0.1,localhost' }
process.env.DSH_HOME = home
process.env.NO_PROXY = cliEnv.NO_PROXY
process.env.no_proxy = cliEnv.no_proxy
const exec = promisify(execFile)
try {
  const imported = name => import(pathToFileURL(path.join(market, 'src', name + '.ts')).href)
  const { checkUpdates, setUpdateRegistry } = await imported('updates')
  const { hasLoadableEntry } = await imported('profile')
  setUpdateRegistry(base)
  async function install(version) {
    const result = await exec(process.execPath, [cli, 'plugin', '--profile', 'web', 'add',
      current.name + '@' + version, '--registry=' + base, '--config.minimum-release-age=0'],
    { env: cliEnv, cwd: temp, timeout: 45_000, windowsHide: true })
    await fs.writeFile(path.join(output, 'install-' + version + '.log'), result.stdout + result.stderr)
    const installedDir = path.join(profile, 'node_modules', current.name)
    assert.equal(JSON.parse(await fs.readFile(path.join(installedDir, 'package.json'), 'utf8')).version, version)
    assert.ok(hasLoadableEntry(profile, current.name))
    assert.deepEqual(await fs.readFile(path.join(installedDir, 'lib/client.js')), await fs.readFile(path.join(root, 'lib/client.js')))
  }
  const check = async force => (await checkUpdates('web', force, profile))[current.name]
  await install(old)
  const baseline = await check(true)
  assert.deepEqual([baseline.kind, baseline.version, baseline.latest, baseline.updateAvailable], ['npm', old, old, false])
  latest = target
  const available = await check(true)
  assert.deepEqual([available.version, available.latest, available.updateAvailable], [old, target, true])
  await install(target)
  const cached = await check(false)
  const fresh = await check(true)
  for (const result of [cached, fresh]) {
    assert.deepEqual([result.version, result.latest, result.updateAvailable], [target, target, false])
  }
  const commit = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: market, encoding: 'utf8' }).trim()
  await fs.writeFile(path.join(output, 'receipt.json'), JSON.stringify({
    marketVersion: marketPackage.version, marketCommit: commit, registry: 'isolated loopback npm protocol',
    installation: 'DSH official Web/CLI package manager', packageSha256: hash(artifacts.get(target), 'sha256', 'hex'),
    baseline, available, cached, fresh, registryLog,
  }, null, 2) + '\n')
  console.log('Local registry upgrade verified: ' + old + ' → ' + target)
} finally {
  server.closeAllConnections()
  server.close()
}
