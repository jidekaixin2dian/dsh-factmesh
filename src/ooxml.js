import { Uint8ArrayReader, ZipReader } from '@zip.js/zip.js'
import { SaxesParser } from 'saxes'

export const MAX_FILE_BYTES = 24 * 1024 * 1024
const PART_BYTES = 8 * 1024 * 1024
const TOTAL_BYTES = 64 * 1024 * 1024
const MAX_ENTRIES = 3000
const MAX_NODES = 220_000
const MAX_DEPTH = 80
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const STRICT_REL = 'http://purl.oclc.org/ooxml/officeDocument/relationships'

export function problem(code, detail = '') {
  return Object.assign(new Error(detail ? code + ': ' + detail : code), { code, detail })
}

export const all = (node, name) => node?.children.filter(child => child.name === name) ?? []
export const one = (node, name) => node?.children.find(child => child.name === name)
export const attr = (node, name) => node?.attributes[name]
export const relId = (node, name = 'id') => node?.relationships[name]
export const value = (node) => node?.text ?? ''

function parseXml(bytes, name, budget) {
  const stack = []
  let root
  const parser = new SaxesParser({ xmlns: true, fileName: name })
  parser.on('doctype', () => { throw problem('unsupportedXml', name) })
  parser.on('opentag', tag => {
    if (--budget.nodes < 0 || stack.length >= MAX_DEPTH) throw problem('xmlLimit', name)
    const node = {
      name: tag.local, namespace: tag.uri, children: [], text: '',
      attributes: Object.create(null), relationships: Object.create(null),
    }
    for (const a of Object.values(tag.attributes)) {
      if (!a.uri) node.attributes[a.local] = a.value
      if (a.uri === REL || a.uri === STRICT_REL) node.relationships[a.local] = a.value
    }
    if (stack.length) stack.at(-1).children.push(node)
    else root = node
    stack.push(node)
  })
  parser.on('closetag', () => { stack.pop() })
  const append = text => { if (stack.length) stack.at(-1).text += text }
  parser.on('text', append)
  parser.on('cdata', append)
  try {
    const encoding = bytes[0] === 0xff && bytes[1] === 0xfe ? 'utf-16le'
      : bytes[0] === 0xfe && bytes[1] === 0xff ? 'utf-16be' : 'utf-8'
    parser.write(new TextDecoder(encoding, { fatal: true }).decode(bytes)).close()
  } catch (error) {
    if (error.code) throw error
    throw problem('invalidXml', name)
  }
  return root
}

function resolvePart(source, target) {
  if (!target || /[\\?#\u0000-\u001f]/u.test(target)) throw problem('invalidOffice', source)
  let decoded
  try { decoded = decodeURIComponent(target) } catch { throw problem('invalidOffice', source) }
  if (/[\\?#:\u0000-\u001f]/u.test(decoded) || decoded.startsWith('//')) throw problem('invalidOffice', source)
  const segments = decoded.startsWith('/') ? [] : source.split('/').slice(0, -1)
  for (const segment of decoded.replace(/^\//u, '').split('/')) {
    if (!segment || segment === '.') continue
    if (segment === '..') {
      if (!segments.length) throw problem('invalidOffice', source)
      segments.pop()
    } else segments.push(segment)
  }
  return segments.join('/')
}

export async function withOfficePackage(bytes, read) {
  if (!bytes.length || bytes.length > MAX_FILE_BYTES) throw problem('fileLimit')
  const reader = new ZipReader(new Uint8ArrayReader(bytes), {
    useWebWorkers: false, checkSignature: true,
  })
  const entries = new Map()
  const budget = { bytes: TOTAL_BYTES, nodes: MAX_NODES }
  const parsed = new Map()
  async function xml(name, optional = false) {
    if (parsed.has(name)) return parsed.get(name)
    const entry = entries.get(name)
    if (!entry && optional) return undefined
    if (!entry) throw problem('missingPart', name)
    const size = entry.uncompressedSize
    if (size > PART_BYTES || size > budget.bytes) throw problem('xmlLimit', name)
    budget.bytes -= size
    const output = new Uint8Array(size)
    let offset = 0
    const writer = new WritableStream({
      write(chunk) {
        if (offset + chunk.length > size) throw problem('xmlLimit', name)
        output.set(chunk, offset)
        offset += chunk.length
      },
    })
    try { await entry.getData(writer) } catch (error) {
      if (error.code === 'xmlLimit') throw error
      throw problem('invalidZip', name)
    }
    if (offset !== size) throw problem('invalidZip', name)
    const tree = parseXml(output, name, budget)
    parsed.set(name, tree)
    return tree
  }
  async function relations(source) {
    const split = source.lastIndexOf('/') + 1
    const name = source.slice(0, split) + '_rels/' + source.slice(split) + '.rels'
    const root = await xml(name, Boolean(source))
    const map = new Map()
    for (const node of all(root, 'Relationship')) {
      const id = attr(node, 'Id')
      const external = attr(node, 'TargetMode') === 'External'
      const type = attr(node, 'Type')?.split('/').at(-1)
      if (!id || !type || map.has(id)) throw problem('invalidOffice', name)
      map.set(id, { type, external, part: external ? null : resolvePart(source, attr(node, 'Target')) })
    }
    return map
  }
  try {
    try {
      let count = 0
      for await (const entry of reader.getEntriesGenerator()) {
        if (++count > MAX_ENTRIES) throw problem('archiveLimit')
        if (entry.directory) continue
        if (entries.has(entry.filename) || entry.encrypted ||
            !Number.isSafeInteger(entry.uncompressedSize) || entry.uncompressedSize < 0) {
          throw problem('invalidZip', entry.filename)
        }
        entries.set(entry.filename, entry)
      }
    } catch (error) {
      if (error.code === 'archiveLimit' || error.code === 'invalidZip') throw error
      throw problem('invalidZip')
    }
    const roots = [...(await relations('')).values()].filter(r => r.type === 'officeDocument' && !r.external)
    if (roots.length !== 1) throw problem('invalidOffice')
    return await read({ main: roots[0].part, relations, xml })
  } finally {
    await reader.close()
  }
}

export function linked(relations, id, type) {
  const match = relations.get(id)
  if (!match || match.type !== type || match.external) throw problem('invalidRelation', id)
  return match.part
}
