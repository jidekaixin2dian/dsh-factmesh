import React, { useEffect, useMemo, useRef, useState } from 'react'
import { Button } from '@deepseek-ai/dsh-client-ui-primitives'
import { en, zh } from './locale.js'
import { validMonth } from './meaning.js'
import { UNIT } from './number.js'
import { problem } from './ooxml.js'
import { candidateIndex, candidatesFor, confirmationProblem, createProjectStore, makeLink, reviewRun } from './reconcile.js'
import { openSlides } from './slides.js'
import { browse, childPath, readSource, unchanged, watchSources } from './workspace.js'
import { openWorkbook } from './workbook.js'
import css from './panel.css'

const ID = 'dsh-factmesh'
const sorts = ['difference', 'reconfirm', 'verify', 'confirm', 'match', 'skipped']
const basename = path => path.split('/').at(-1)
const location = cell => cell.sheet + '!' + cell.address

function message(error, t) {
  return t(error?.code && Object.hasOwn(en, error.code) ? error.code : 'generic') +
    (error?.detail ? ' ' + error.detail : '')
}

function Icon({ size = 20 }) {
  return <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor"
    strokeWidth="1.65" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
    <path d="M4 5h16v14H4zM4 10h16M10 10v9m3-4 2 2 3-4" />
  </svg>
}

function Panel({ t, useSessions, fileRemote }) {
  const sessions = useSessions(snapshot => snapshot)
  const [session, setSession] = useState('')
  const byDirectory = new Map()
  for (const id of sessions.ids) {
    const item = sessions.byId[id]
    if (item?.cwd && !item.origin && !byDirectory.has(item.cwd)) byDirectory.set(item.cwd, item)
  }
  const options = [...byDirectory.values()]
  const selected = options.find(item => item.id === session) ?? (options.length === 1 ? options[0] : null)
  return <section className="fm-panel" data-factmesh>
    <style>{css}</style>
    <header className="fm-head">
      <div className="fm-mark"><Icon size={24} /></div>
      <div><h1>FactMesh</h1><p>{t('subtitle')}</p></div>
    </header>
    {!fileRemote ? <p role="alert">{t('unavailable')}</p> : <>
      <label className="fm-field fm-workspace">{t('workspace')}
        <select value={selected?.id ?? ''} onChange={event => setSession(event.target.value)}>
          <option value="">{t('choose')}</option>
          {options.map(item => <option key={item.id} value={item.id}>{item.cwd}</option>)}
        </select>
      </label>
      {selected ? <Project key={selected.id} t={t} session={selected} remote={fileRemote} />
        : <p className="fm-note">{t('noWorkspace')}</p>}
    </>}
  </section>
}

function FileSelect({ kind, t, dir, listing, selected, onChange }) {
  const choices = (listing?.entries ?? []).filter(entry => entry.type === 'file' &&
    !entry.name.startsWith('~$') && entry.name.toLowerCase().endsWith('.' + kind))
    .map(entry => childPath(dir, entry.name))
  if (selected && !choices.includes(selected)) choices.unshift(selected)
  return <label className="fm-field">{t(kind)}
    <select value={selected} onChange={event => onChange(event.target.value)}>
      <option value="">{t('choose')}</option>
      {choices.map(path => <option key={path} value={path}>{path}</option>)}
    </select>
  </label>
}

function Project({ t, session, remote }) {
  const [dir, setDir] = useState('.')
  const [listing, setListing] = useState(null)
  const [listRefresh, refresh] = useState(0)
  const [listError, setListError] = useState(null)
  const [xlsx, setXlsx] = useState('')
  const [pptx, setPptx] = useState('')
  const [period, setPeriod] = useState('')
  const [project, setProject] = useState(t('projectDefault'))
  const [stage, setStage] = useState(null)
  const [sheet, setSheet] = useState('')
  const [run, setRun] = useState(null)
  const [links, setLinks] = useState([])
  const [stale, setStale] = useState(null)
  const [error, setError] = useState(null)
  const [busy, setBusy] = useState(false)
  const operation = useRef(null)
  const invalidated = useRef(false)
  const store = useMemo(() => createProjectStore({
    getItem: (...args) => localStorage.getItem(...args),
    setItem: (...args) => localStorage.setItem(...args),
  }, [session.cwd, project.trim()]), [session.cwd, project])

  useEffect(() => {
    const controller = new AbortController()
    setListing(null)
    setListError(null)
    browse(remote, session.id, dir, controller.signal).then(
      data => { if (!controller.signal.aborted) setListing({ dir, ...data }) },
      reason => { if (!controller.signal.aborted) setListError(reason) },
    )
    return () => controller.abort()
  }, [dir, listRefresh, remote, session.id])

  useEffect(() => () => operation.current?.abort(), [])
  useEffect(() => {
    if (!run || stale) return
    const controller = new AbortController()
    const mark = (path, code = 'stale') => {
      if (!controller.signal.aborted) {
        invalidated.current = true
        setStale(problem(code, path))
      }
    }
    void watchSources(remote, session.id, run.files, controller.signal, mark)
    let checking = false
    async function verify() {
      if (checking || controller.signal.aborted) return
      checking = true
      try { await unchanged(remote, session.id, run.files, controller.signal) }
      catch (reason) { mark(reason.detail, versionProblem(reason)) }
      finally { checking = false }
    }
    const interval = setInterval(verify, 3000)
    window.addEventListener('focus', verify)
    return () => {
      controller.abort()
      clearInterval(interval)
      window.removeEventListener('focus', verify)
    }
  }, [run, stale, remote, session.id])
  useEffect(() => {
    if (!run) return
    function onStorage(event) {
      if (event.key !== store.key && event.key !== null) return
      try { setLinks(store.read()); setError(null) }
      catch { setError(problem('storageError')) }
    }
    window.addEventListener('storage', onStorage)
    return () => window.removeEventListener('storage', onStorage)
  }, [run, store])

  function start(task) {
    operation.current?.abort()
    const controller = new AbortController()
    operation.current = controller
    setError(null)
    setBusy(true)
    void Promise.resolve().then(() => task(controller.signal)).catch(reason => {
      if (!controller.signal.aborted) setError(reason)
    }).finally(() => {
      if (operation.current !== controller) return
      operation.current = null
      setBusy(false)
    })
  }

  function invalidate(full = true) {
    operation.current?.abort()
    operation.current = null
    invalidated.current = true
    setBusy(false)
    if (full) {
      setStage(null)
      setSheet('')
    }
    setRun(null)
    setStale(null)
    setError(null)
  }

  function readCatalog() {
    invalidate()
    start(async signal => {
      const book = await readSource(remote, session.id, session.cwd, xlsx, signal)
      const catalog = await openWorkbook(book.data)
      const supported = catalog.sheets.filter(candidate => candidate.supported)
      if (!supported.length) throw problem('noSheet')
      if (!signal.aborted) {
        setStage(catalog)
        setSheet(supported.find(candidate => !candidate.hidden)?.name ?? supported[0].name)
      }
    })
  }

  function reconcile() {
    invalidate(false)
    start(async signal => {
      const [deck, book] = await Promise.all([
        readSource(remote, session.id, session.cwd, pptx, signal),
        readSource(remote, session.id, session.cwd, xlsx, signal),
      ])
      const files = { pptx: deck.file, xlsx: book.file }
      const slides = await openSlides(deck.data, period)
      const table = await openWorkbook(book.data, sheet)
      await unchanged(remote, session.id, files, signal)
      if (signal.aborted) return
      let saved
      try { saved = store.read() } catch { throw problem('storageError') }
      invalidated.current = false
      setLinks(saved)
      setRun({ files, period, sheet: table.sheet, slides })
    })
  }

  function confirm(slide, cell, replaceId) {
    start(async signal => {
      try { await unchanged(remote, session.id, run.files, signal) } catch (reason) {
        if (!signal.aborted) {
          invalidated.current = true
          setStale(problem(versionProblem(reason), reason.detail))
        }
        return
      }
      const link = makeLink(run, slide, cell)
      function write() {
        if (signal.aborted || invalidated.current) return
        try { setLinks(store.save(link, replaceId)) } catch { throw problem('storageError') }
      }
      if (navigator.locks) await navigator.locks.request(store.key, { signal }, write)
      else write()
    })
  }

  const valid = pptx && xlsx && validMonth(period) && project.trim()
  const directories = listing?.dir === dir ? listing.entries.filter(entry => entry.type === 'directory') : []
  return <>
    <div className="fm-setup">
      <div className="fm-fields">
        <label className="fm-field">{t('project')}
          <input value={project} maxLength={100} title={t('projectHint')}
            onChange={event => { invalidate(false); setProject(event.target.value) }} />
        </label>
        <label className="fm-field">{t('period')}
          <input placeholder="2026-09" value={period} maxLength={7}
            onChange={event => { invalidate(false); setPeriod(event.target.value.trim()) }} />
        </label>
      </div>
      <div className="fm-folder">
        <span className="fm-path" title={dir}>{dir === '.' ? t('root') : dir}</span>
        <Button size="sm" variant="outline" disabled={dir === '.'}
          onClick={() => setDir(dir.includes('/') ? dir.slice(0, dir.lastIndexOf('/')) : '.')}>{t('parent')}</Button>
        <Button size="sm" variant="outline" onClick={() => refresh(listRefresh + 1)}>{t('reload')}</Button>
        {directories.length > 0 && <select aria-label={t('folder')} value=""
          onChange={event => setDir(childPath(dir, event.target.value))}>
          <option value="" disabled>{t('folder')}</option>
          {directories.map(entry => <option key={entry.name} value={entry.name}>{entry.name}</option>)}
        </select>}
      </div>
      {listing?.truncated && <p className="fm-note">{t('truncated')}</p>}
      {listError && <p role="alert">{message(listError, t)}</p>}
      <div className="fm-fields">
        <FileSelect kind="pptx" t={t} dir={dir} listing={listing?.dir === dir ? listing : null} selected={pptx}
          onChange={value => { invalidate(false); setPptx(value) }} />
        <FileSelect kind="xlsx" t={t} dir={dir} listing={listing?.dir === dir ? listing : null} selected={xlsx}
          onChange={value => { invalidate(); setXlsx(value) }} />
      </div>
      <div className="fm-actions">
        <Button variant={stage ? 'outline' : 'primary'} disabled={!valid || busy} onClick={readCatalog}>{t('sheets')}</Button>
        {stage && <>
          <label className="fm-field">{t('sheet')}
            <select aria-label={t('sheet')} value={sheet} onChange={event => { invalidate(false); setSheet(event.target.value) }}>
              {stage.sheets.filter(item => item.supported).map(item => <option key={item.name} value={item.name}>
                {item.name}{item.hidden ? ' (' + t('hidden') + ')' : ''}
              </option>)}
            </select>
          </label>
          <Button variant="primary" disabled={!valid || !sheet || busy} onClick={reconcile}>{t(run ? 'rerun' : 'run')}</Button>
        </>}
        {busy && <span role="status" className="fm-note">{t('busy')}</span>}
      </div>
    </div>
    {error && <p className="fm-alert" role="alert">{message(error, t)}</p>}
    {stale && <p className="fm-alert" role="alert">{message(stale, t)}</p>}
    {run ? <Results key={JSON.stringify([run.files, run.period, run.sheet.name])}
      run={run} links={links} t={t} disabled={Boolean(busy || stale)} stale={Boolean(stale)} onConfirm={confirm} />
      : !stage && <p className="fm-note">{t('selectFiles')}</p>}
  </>
}

function versionProblem(reason) {
  return reason?.code === 'stale' || reason?.detail?.endsWith('/not-found') ? 'stale' : 'refreshFailed'
}

function Coverage({ run, t }) {
  const { slides, textBoxes, tables, omitted } = run.slides.coverage
  const counts = [[slides, 'pages'], [textBoxes, 'boxes'], [tables, 'tables'],
    [run.slides.numbers.length, 'numbers'], [run.sheet.scanned, 'scanned']]
  return <details className="fm-coverage">
    <summary>{t('scope')} <span>{counts.map(([count, label]) => count + ' ' + t(label)).join(' · ')}</span>
      <span data-omitted>{t('omitted')} {omitted.length}</span>
    </summary>
    <p className="fm-note">{basename(run.files.pptx.path)} · {basename(run.files.xlsx.path)} / {run.sheet.name} · {run.period}</p>
    {omitted.length ? <ul className="fm-omitted">{omitted.slice(0, 100).map((item, i) => <li key={i}>
      {t('atPage', { page: item.page })} · {item.title || item.shapeName} · {t(item.kind)}
    </li>)}{omitted.length > 100 && <li>{t('omittedMore', { count: omitted.length - 100 })}</li>}</ul>
      : <p className="fm-note">{t('noOmitted')}</p>}
  </details>
}

function Results({ run, links, t, disabled, stale, onConfirm }) {
  const [selected, setSelected] = useState('')
  const [filter, setFilter] = useState('all')
  const [skips, setSkips] = useState(new Set())
  const { rows, missing } = useMemo(() => reviewRun(run, links), [run, links])
  const index = useMemo(() => candidateIndex(run.sheet.cells), [run.sheet.cells])
  const rowsWithSkips = rows.map(row => skips.has(row.slide.id) ? { ...row, status: 'skipped' } : row)
  const counts = Object.fromEntries(sorts.map(key => [key, 0]))
  for (const row of rowsWithSkips) counts[row.status]++
  const ordered = rowsWithSkips.filter(row => filter === 'all' || row.status === filter)
    .sort((a, b) => sorts.indexOf(a.status) - sorts.indexOf(b.status) || a.slide.page - b.slide.page)
  const current = ordered.find(item => item.slide.id === selected) ?? ordered[0]
  const reusable = rows.filter(row => row.reused).length
  return <div data-results data-stale={stale || undefined}>
    <Coverage run={run} t={t} />
    <div className="fm-filters" aria-label={t('selected')}>
      {['all', ...sorts].map(status => <Button size="sm" variant={filter === status ? 'outline' : 'ghost'} key={status}
        aria-pressed={filter === status} onClick={() => setFilter(status)}>
        {t(status)} {status === 'all' ? rows.length : counts[status]}
      </Button>)}
      {reusable > 0 && <span className="fm-note" data-reused>{t('reused')} {reusable}</span>}
    </div>
    {missing.length > 0 && <details className="fm-coverage">
      <summary>{t('missing')} · {missing.length}</summary>
      <ul>{missing.map(item => <li key={item.id}>{item.source.title} / {item.source.metric} → {location(item.target)}</li>)}</ul>
    </details>}
    <div className="fm-split">
      <div className="fm-list">
        <h2>{t('slide')}</h2>
        {!ordered.length && <p className="fm-note">{t('emptyFilter')}</p>}
        {ordered.map(row => <button type="button" key={row.slide.id} className="fm-row"
          data-status={row.status} data-metric={row.slide.metric} aria-pressed={current?.slide.id === row.slide.id}
          onClick={() => setSelected(row.slide.id)}>
          <span>{t('atPage', { page: row.slide.page })} · {row.slide.title}</span>
          <strong>{row.slide.metric} <b>{row.slide.shown}</b></strong>
          <small className="fm-state">{t(row.status)}{row.reused ? ' · ' + t('reused') : ''}</small>
        </button>)}
      </div>
      {current && <Detail key={current.slide.id + ':' + (current.link?.id ?? '')} row={current} t={t}
        index={index} period={run.period} disabled={disabled} stale={stale}
        missing={missing} onConfirm={onConfirm} onSkip={() => setSkips(previous => {
          const next = new Set(previous)
          if (next.has(current.slide.id)) next.delete(current.slide.id)
          else next.add(current.slide.id)
          return next
        })} />}
    </div>
  </div>
}

function Detail({ row, t, index, period, disabled, stale, missing, onConfirm, onSkip }) {
  const [selection, setSelection] = useState('')
  const [address, setAddress] = useState('')
  const [missingCell, setMissingCell] = useState(false)
  const [replace, setReplace] = useState('')
  const { slide } = row
  const candidates = candidatesFor(slide, index, period)
  const sameMonth = candidates.filter(cell => cell.period === period)
  const fallback = row.target ?? (sameMonth.length === 1 ? sameMonth[0] : undefined)
  const cell = index.byAddress.get(selection) ?? fallback
  const reason = cell ? confirmationProblem(slide, cell, period) || cell.reason : ''
  const options = [...new Map([...(fallback ? [fallback] : []), ...candidates.slice(0, 60),
    ...(cell ? [cell] : [])].map(item => [item.address, item])).values()]
  const confirmed = row.target && cell?.address === row.target.address
  const block = cell ? confirmationProblem(slide, cell, period) : ''
  const comparison = confirmed && !stale && row.rounded !== undefined
  const factor = comparison && cell.factor !== '1' ? ' × ' + cell.factor : ''
  const conversion = comparison && cell.unit !== slide.unit ? ' × ' + UNIT[cell.unit][1] + ' ÷ ' + UNIT[slide.unit][1] : ''
  return <article className="fm-detail">
    <h2>{t('target')}</h2>
    <p className="fm-loc">{t('atPage', { page: slide.page })} · {slide.title}
      {slide.section ? ' / ' + slide.section : ''}
      {slide.shapeName ? ' / ' + slide.shapeName : ''}{slide.row ? ' · R' + slide.row + 'C' + slide.col : ''}</p>
    <blockquote>{slide.context}</blockquote>
    {row.reason === 'slidePeriod' && <p className="fm-note">{t(row.reason)}</p>}
    {row.status === 'reconfirm' && <p className="fm-note">{t('reconfirmHint')}</p>}
    {comparison && <div className="fm-comparison" data-comparison>
      <strong>{row.status === 'difference'
        ? t(row.direction >= 0 ? 'higher' : 'lower', { delta: row.delta.replace(/^-/, ''),
          unit: row.differenceUnit === '百分点' ? t('ppGap') : row.differenceUnit })
        : t(row.status === 'skipped' ? 'skipped' : 'match')}</strong>
      <p>{t('rawRule', { raw: cell.raw, operations: factor + conversion, converted: row.workbook, unit: row.unit })}</p>
      <p>{t('rounding', { precision: row.precision, value: row.rounded, unit: row.unit })}</p>
    </div>}
    <label className="fm-field">{t('candidates')}
      <select value={cell?.address ?? ''} onChange={event => { setSelection(event.target.value); setMissingCell(false) }}>
        <option value="" disabled>{t('choose')}</option>
        {options.map(option => <option key={option.address} value={option.address}>
          {location(option)} · {option.rowTitle} · {option.colTitle}
        </option>)}
      </select>
    </label>
    {!options.length && <p className="fm-note">{t('noCandidates')}</p>}
    {candidates.length > 60 && <p className="fm-note">{t('candidateMore')}</p>}
    <form className="fm-actions" onSubmit={event => {
      event.preventDefault()
      const key = address.trim().toUpperCase().replaceAll('$', '')
      const found = index.byAddress.has(key)
      setMissingCell(!found)
      if (found) setSelection(key)
    }}>
      <label className="fm-field">{t('chooseAddress')}
        <input placeholder={t('addressHint')} value={address} maxLength={14}
          onChange={event => setAddress(event.target.value)} />
      </label>
      <Button type="submit" variant="outline" disabled={!address.trim()}>{t('locate')}</Button>
    </form>
    {missingCell && <p className="fm-note" role="status">{t('noCell')}</p>}
    {cell && <div className="fm-evidence" data-cell={location(cell)}>
      <h3>{location(cell)}</h3>
      <dl>
        <dt>{t('rowTitle')}</dt><dd>{cell.rowTitle || t('unknown')}</dd>
        <dt>{t('colTitle')}</dt><dd>{cell.colTitle || t('unknown')}</dd>
        <dt>{t('cellPeriodLabel')}</dt><dd>{cell.period || t('unknown')}</dd>
        <dt>{t('unit')}</dt><dd>{cell.unit || t('unitless')}</dd>
        <dt>{t('sourceRaw')}</dt><dd><code>{cell.raw || '—'}</code></dd>
        {cell.format && <><dt>{t('format')}</dt><dd><code>{cell.format}</code></dd></>}
        {cell.formula && <><dt>{t('formula')}</dt><dd><code>{cell.formula}</code></dd></>}
      </dl>
      {reason && <p className="fm-note">{t(reason)}</p>}
      {confirmed && row.reused && <p className="fm-note">{t('before')} · {t('atPage', { page: row.link.source.page })}
        {' → ' + location(row.link.target) + ' · ' + row.link.origin.period}</p>}
    </div>}
    {!row.target && missing.length > 0 && <label className="fm-field">{t('rebind')}
      <select value={replace} onChange={event => setReplace(event.target.value)}>
        <option value="">{t('newLink')}</option>
        {missing.map(link => <option key={link.id} value={link.id}>{link.source.title} / {link.source.metric}</option>)}
      </select>
    </label>}
    <div className="fm-actions">
      {cell && !confirmed && <Button variant="primary" disabled={disabled || Boolean(block)}
        onClick={() => onConfirm(slide, cell, row.link?.id ?? (replace || (row.linkIds.length === 1 ? row.linkIds[0] : undefined)))}>
        {t(row.target ? 'changeLink' : 'confirmLink')}
      </Button>}
      <Button variant="outline" disabled={disabled} onClick={onSkip}>{t(row.status === 'skipped' ? 'undoSkip' : 'skip')}</Button>
    </div>
  </article>
}

export const name = ID
export const inject = ['slots', 'locale', 'remote', 'remote.workspaceFiles']

export function apply(ctx) {
  ctx.effect(() => ctx.locale.register(ID, { en, zh }), ID + ': locale')
  const t = ctx.locale.bind(ID)
  ctx.slots.inject('main', () => ctx.slots.register({
    name: 'main', key: ID, locale: ID, inject: () => ({ fileRemote: ctx.remote.workspaceFiles }),
  }, Panel))
  ctx.slots.inject('sidebar.panellist', () => ctx.slots.register({
    name: 'sidebar.panellist', id: ID, order: 25, locale: ID, label: () => t('nav'),
  }, Icon))
}
