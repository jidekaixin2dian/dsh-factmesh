import { MAX_FILE_BYTES, problem } from './ooxml.js'

const CHUNK = 1024 * 1024
const TIMEOUT_MS = 15_000

async function unwrap(promise) {
  const result = await promise
  if (!result.ok) throw problem('workspaceError', result.error?.code ?? '')
  return result.value
}

function deadline(signal) {
  return AbortSignal.any([signal, AbortSignal.timeout(TIMEOUT_MS)])
}

export function childPath(dir, name) {
  if (!name || /[\\/\u0000]/u.test(name) || /^[a-z]:/iu.test(name) || name === '.' || name === '..') {
    throw problem('workspaceError')
  }
  return dir === '.' ? name : dir + '/' + name
}

export function browse(files, session, dir, signal) {
  return unwrap(files.list(session, dir, deadline(signal)))
}

function absoluteFor(cwd, relative) {
  const root = cwd.replaceAll('\\', '/').replace(/\/+$/u, '')
  const parts = relative.split('/')
  for (const part of parts) childPath('.', part)
  if (!/^(?:\/|[a-z]:$|[a-z]:\/)/iu.test(root) && root !== '') throw problem('workspaceError', relative)
  return root + '/' + parts.join('/')
}

function pathMatches(cwd, expected, reported) {
  const actual = reported.replaceAll('\\', '/')
  return /^(?:[a-z]:|\\\\|\/\/)/iu.test(cwd) ? actual.toLowerCase() === expected.toLowerCase() : actual === expected
}

async function selectable(files, session, path, signal) {
  const split = path.lastIndexOf('/') + 1
  const parent = split ? path.slice(0, split - 1) : '.'
  const listing = await browse(files, session, parent, signal)
  if (!listing.entries.some(entry => entry.name === path.slice(split) && entry.type === 'file')) {
    throw problem('workspaceError', path)
  }
}

export async function unchanged(files, session, selected, signal) {
  await Promise.all(Object.values(selected).map(async record => {
    const result = await unwrap(files.stat(session, record.path, deadline(signal)))
    if (!same(result, record)) throw problem('stale', record.path)
  }))
}

function same(a, b) {
  return a.absolutePath === b.absolutePath && a.version === b.version && a.bytes === b.bytes
}

export async function readSource(files, session, cwd, path, signal) {
  const request = deadline(signal)
  const expected = absoluteFor(cwd, path)
  await selectable(files, session, path, request)
  const info = await unwrap(files.stat(session, path, request))
  if (!pathMatches(cwd, expected, info.absolutePath)) throw problem('workspaceError', path)
  if (!Number.isSafeInteger(info.bytes) || info.bytes <= 0 || info.bytes > MAX_FILE_BYTES) {
    throw problem('fileLimit', path)
  }
  const data = new Uint8Array(info.bytes)
  for (let offset = 0; offset < data.length; offset += CHUNK) {
    const length = Math.min(CHUNK, data.length - offset)
    const chunk = await unwrap(files.readBytes(session, path, { range: { offset, length } }, request))
    if (!same(info, chunk) || chunk.offset !== offset || chunk.data?.length !== length ||
        chunk.eof !== (offset + length === data.length)) throw problem('stale', path)
    data.set(chunk.data, offset)
  }
  const file = { ...info, path }
  await selectable(files, session, path, request)
  await unchanged(files, session, { file }, request)
  return { file, data }
}

export async function watchSources(files, session, selected, signal, onStale) {
  await Promise.all(Object.values(selected).map(async record => {
    try {
      for await (const frame of files.changes(session, record.path, signal)) {
        if (signal.aborted || !frame.ok) return
        const message = frame.value
        if (message.kind === 'ready' || message.kind === 'change') {
          await unchanged(files, session, { record }, signal)
        }
      }
    } catch (error) {
      if (!signal.aborted) onStale(record.path, error.code === 'workspaceError' && !error.detail.endsWith('/not-found')
        ? 'refreshFailed' : 'stale')
    }
  }))
}
