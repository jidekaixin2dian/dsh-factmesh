import { build } from 'esbuild'
import { writeFile } from 'node:fs/promises'

const result = await build({
  entryPoints: ['src/client.jsx'],
  outfile: 'lib/client.js',
  bundle: true,
  format: 'cjs',
  platform: 'browser',
  target: 'es2022',
  external: ['react', '@deepseek-ai/dsh-client-ui-primitives'],
  loader: { '.css': 'text' },
  legalComments: 'eof',
  write: false,
  banner: { js: `;window.__ModuleLoader__.load({ id: '@jidekaixin2dian/dsh-factmesh', factory: (require) => {
const module = { exports: {} };` },
  footer: { js: `Object.defineProperty(module.exports, Symbol.toStringTag, { value: 'Module' });
return module.exports;
} });` },
})
const [client] = result.outputFiles
await writeFile(client.path, client.text.replace(/^([ \t]*\/\*\*)[ \t]+$/gmu, '$1'))
