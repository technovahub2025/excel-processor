import { useMemo, useState } from 'react'
import axios from 'axios'
import { FileSpreadsheet, UploadCloud, ArrowRight, CheckCircle2, Download, Plus, Trash2, RefreshCw, AlertCircle, Filter, GitMerge } from 'lucide-react'

const api = axios.create({ baseURL: 'http://127.0.0.1:8000' })

function App() {
  const [source, setSource] = useState(null)
  const [target, setTarget] = useState(null)
  const [pairs, setPairs] = useState([{ source: '', target: '' }])
  const [transferColumns, setTransferColumns] = useState([])
  const [busy, setBusy] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [error, setError] = useState('')
  const [result, setResult] = useState(null)
  const [filters, setFilters] = useState({})
  const [exportColumns, setExportColumns] = useState([])
  const [exportColumnDraft, setExportColumnDraft] = useState([])
  const [exportPickerOpen, setExportPickerOpen] = useState(false)
  const [exportColumnSearch, setExportColumnSearch] = useState('')

  async function upload(file, kind) {
    setBusy(true); setError(''); setResult(null)
    try {
      const body = new FormData(); body.append('file', file)
      const { data } = await api.post('/api/upload', body)
      const entry = { ...data, size: file.size }
      if (kind === 'source') { setSource(entry); setTransferColumns([]); setPairs([{ source: '', target: '' }]) }
      else { setTarget(entry); setPairs([{ source: '', target: '' }]) }
    } catch (e) { setError(e.response?.data?.detail || e.message || 'Upload failed') }
    finally { setBusy(false) }
  }

  async function changeSheet(kind, sheet) {
    const entry = kind === 'source' ? source : target
    if (!entry) return
    setBusy(true); setError(''); setResult(null)
    try {
      const body = new FormData(); body.append('file_id', entry.file_id); body.append('sheet_name', sheet)
      const { data } = await api.post('/api/sheet', body)
      const next = { ...entry, ...data, columns: data.columns ?? entry.columns }
      if (kind === 'source') { setSource(next); setTransferColumns([]); setPairs(old => old.map(() => ({ source: '', target: '' }))) }
      else { setTarget(next); setPairs(old => old.map(() => ({ source: '', target: '' }))) }
    } catch (e) { setError(e.response?.data?.detail || e.message) }
    finally { setBusy(false) }
  }

  function updatePair(index, key, value) {
    setPairs(old => old.map((pair, i) => i === index ? { ...pair, [key]: value } : pair)); setResult(null)
  }

  function toggleTransferColumn(column) {
    setTransferColumns(current => current.includes(column) ? current.filter(item => item !== column) : [...current, column])
    setResult(null)
  }

  async function findMatches() {
    if (!source || !target) return setError('Upload both the large/source and small/target Excel files.')
    if (!pairs.length || pairs.some(pair => !pair.source || !pair.target)) return setError('Choose a source and target column for every matching pair.')
    if (!transferColumns.length) return setError('Select at least one source column to add to the target.')
    setBusy(true); setError('')
    try {
      const body = new FormData()
      body.append('source_id', source.file_id); body.append('target_id', target.file_id)
      body.append('config_json', JSON.stringify({ matching_pairs: pairs, transfer_columns: transferColumns }))
      const { data } = await api.post('/api/match', body)
      setResult(data); setFilters({})
      const allColumns = data.columns.map((_, index) => index)
      setExportColumns(allColumns); setExportColumnDraft(allColumns)
    } catch (e) { setError(e.response?.data?.detail || e.message || 'Matching failed') }
    finally { setBusy(false) }
  }

  async function exportFile(format) {
    if (!result) return
    if (!exportColumns.length) return setError('Select at least one export column before downloading.')
    setExporting(true); setError('')
    try {
      const body = new FormData(); body.append('result_id', result.result_id); body.append('format', format); body.append('filters_json', JSON.stringify(filters)); body.append('columns_json', JSON.stringify(exportColumns))
      const response = await api.post('/api/export', body, { responseType: 'blob' })
      const url = URL.createObjectURL(response.data); const a = document.createElement('a')
      a.href = url; a.download = `matched_target.${format}`; a.click(); URL.revokeObjectURL(url)
    } catch (e) { setError('Export failed. ' + (e.response?.data?.detail || e.message)) }
    finally { setExporting(false) }
  }

  const columns = result?.columns || []
  const exportColumnOptions = columns.map((name, index) => ({ name, index, display: columns.slice(0, index).filter(column => column === name).length ? `${name} (${columns.slice(0, index).filter(column => column === name).length + 1})` : name }))
  const filteredExportColumns = exportColumnOptions.filter(column => column.display.toLowerCase().includes(exportColumnSearch.toLowerCase()))
  const visiblePreview = useMemo(() => (result?.preview || []).filter(row => Object.entries(filters).every(([key, value]) => !value || String(row[key] ?? '').toLowerCase().includes(value.toLowerCase()))), [result, filters])

  return <div className="app-shell">
    <header className="topbar"><div className="brand"><div className="brand-icon"><FileSpreadsheet size={22}/></div><div><strong>SheetFlow</strong><span>UNIVERSAL EXCEL PROCESSOR</span></div></div><div className="local-pill"><span className="green-dot"/> Local processing workspace</div></header>
    <main className="container">
      <section className="hero"><div className="eyebrow">SOURCE TO TARGET WORKFLOW</div><h1>Bring the right data<br/><em>into your spreadsheet.</em></h1><p>Match rows across two workbooks using the fields they share, then add a selected source column to every target row.</p><div className="hero-tags"><span>XLSX · XLS · XLSM · CSV</span><span>Multiple match fields</span><span>Target rows preserved</span></div></section>
      {error && <div className="alert"><AlertCircle size={18}/><span>{error}</span><button onClick={() => setError('')}>Dismiss</button></div>}

      <section className="section-heading"><div><span className="step">01</span><div><h2>Large / Source Excel</h2><p>Choose the workbook that contains the value you want to transfer.</p></div></div></section>
      <div className="upload-grid single-upload"><UploadCard title="Large / Source Excel" subtitle="Source of matching data and the columns to add" entry={source} onFile={file => upload(file, 'source')} onRemove={() => { setSource(null); setResult(null); setTransferColumns([]) }} onSheet={sheet => changeSheet('source', sheet)} busy={busy}/></div>

      <section className="section-heading"><div><span className="step">02</span><div><h2>Small / Target Excel</h2><p>Every row and existing column in this workbook is kept in the output.</p></div></div></section>
      <div className="upload-grid single-upload"><UploadCard title="Small / Target Excel" subtitle="The base workbook to enrich; it does not need the source primary key" entry={target} onFile={file => upload(file, 'target')} onRemove={() => { setTarget(null); setResult(null) }} onSheet={sheet => changeSheet('target', sheet)} busy={busy}/></div>

      {source && target && <section className="mapping-section">
        <div className="section-heading"><div><span className="step">03</span><div><h2>Match Records</h2><p>Map one or more corresponding fields. Source and target column names can differ.</p></div></div></div>
        <div className="mapping-card">
          <div className="pair-heading"><span>Source / Large Excel</span><span></span><span>Target / Small Excel</span></div>
          {pairs.map((pair, index) => <div className="mapping-controls pair-row" key={index}>
            <label><span>Source column</span><select value={pair.source} onChange={e => updatePair(index, 'source', e.target.value)}><option value="">Select source column</option>{source.columns.map(c => <option key={c}>{c}</option>)}</select></label>
            <div className="map-arrow"><ArrowRight size={19}/></div>
            <label><span>Target column</span><select value={pair.target} onChange={e => updatePair(index, 'target', e.target.value)}><option value="">Select target column</option>{target.columns.map(c => <option key={c}>{c}</option>)}</select></label>
            {pairs.length > 1 && <button className="icon-button pair-remove" title="Remove matching pair" onClick={() => setPairs(old => old.filter((_, i) => i !== index))}><Trash2 size={16}/></button>}
          </div>)}
          <button className="add-pair" onClick={() => setPairs(old => [...old, { source: '', target: '' }])}><Plus size={16}/> Add Matching Column</button>
          <div className="transfer-picker">
            <div className="transfer-picker-heading"><div><strong>04 · Columns to Add to Target</strong><span>{transferColumns.length} of {source.columns.length} selected</span></div><div className="transfer-actions"><button type="button" onClick={() => { setTransferColumns([...source.columns]); setResult(null) }}>Select All</button><button type="button" onClick={() => { setTransferColumns([]); setResult(null) }}>Clear All</button></div></div>
            <div className="transfer-column-list" role="group" aria-label="Source columns to add to target">
              {source.columns.map((column, index) => <label className={`transfer-option ${transferColumns.includes(column) ? 'selected' : ''}`} key={`${column}-${index}`}><input type="checkbox" checked={transferColumns.includes(column)} onChange={() => toggleTransferColumn(column)}/><span>{column}</span></label>)}
            </div>
            {transferColumns.length > 0 && <div className="selected-transfer-list"><span>Selected</span>{transferColumns.map(column => <span className="selected-transfer-chip" key={column}>{column}</span>)}</div>}
            <p>Source primary keys can be selected for transfer, but are never required in the target workbook. Matching and transfer both use one processing request.</p>
          </div>
        </div>
        <button className="primary-button" disabled={busy} onClick={findMatches}><GitMerge size={18}/>{busy ? 'Finding matches…' : 'Find Matches'}<ArrowRight size={18}/></button>
      </section>}

      {result && <section className="results-section">
        <div className="section-heading"><div><span className="step">05</span><div><h2>Matching Results</h2><p>Output is based on {result.target_filename}; unmatched and ambiguous target rows remain.</p></div></div><div className="result-count"><strong>{result.row_count.toLocaleString()}</strong><span>target records</span></div></div>
        <div className="stats-row"><div><CheckCircle2/><strong>{result.matched.toLocaleString()}</strong><span>Matched</span></div><div><AlertCircle/><strong>{result.unmatched.toLocaleString()}</strong><span>Not matched</span></div><div><AlertCircle/><strong>{result.ambiguous.toLocaleString()}</strong><span>Ambiguous</span></div></div>
        <div className="table-toolbar"><div><Filter size={17}/><strong>Data preview</strong><small>First 100 rows</small></div><div className="export-buttons export-controls"><div className="export-picker-wrap"><button className="secondary-button" type="button" aria-expanded={exportPickerOpen} onClick={() => { setExportColumnDraft(exportColumns); setExportPickerOpen(open => !open) }}>Export Columns ({exportColumns.length}) <span>▾</span></button>{exportPickerOpen && <div className="export-picker"><div className="export-picker-title"><strong>Export Columns</strong><span>{exportColumnDraft.length} selected</span></div><input className="export-search" placeholder="Search columns" value={exportColumnSearch} onChange={event => setExportColumnSearch(event.target.value)}/><div className="export-column-list">{filteredExportColumns.map(column => <label className="export-column-option" key={column.index}><input type="checkbox" checked={exportColumnDraft.includes(column.index)} onChange={() => setExportColumnDraft(current => current.includes(column.index) ? current.filter(index => index !== column.index) : [...current, column.index])}/><span>{column.display}</span></label>)}</div><div className="export-picker-actions"><button type="button" onClick={() => setExportColumnDraft(columns.map((_, index) => index))}>Select All</button><button type="button" onClick={() => setExportColumnDraft([])}>Clear All</button><button type="button" className="apply-export" onClick={() => { setExportColumns(exportColumnDraft); setExportPickerOpen(false) }}>Apply</button></div><div className="export-selected-names"><span>Selected:</span>{exportColumnDraft.length ? exportColumnOptions.filter(column => exportColumnDraft.includes(column.index)).map(column => <span className="selected-transfer-chip" key={column.index}>{column.display}</span>) : <em>None</em>}</div></div>}</div><button className="secondary-button" disabled={exporting} onClick={() => exportFile('csv')}><Download size={16}/> CSV</button><button className="primary-button small" disabled={exporting} onClick={() => exportFile('xlsx')}><Download size={16}/> Excel</button></div></div>
        <div className="filter-grid">{columns.map(column => <label key={column}><span>{column}</span><input placeholder={`Filter ${column}`} value={filters[column] || ''} onChange={e => setFilters(old => ({ ...old, [column]: e.target.value }))}/></label>)}</div>
        <div className="table-wrap"><table><thead><tr>{columns.map(column => <th key={column}>{column}</th>)}</tr></thead><tbody>{visiblePreview.map((row, index) => <tr key={index}>{columns.map(column => <td key={column}>{row[column] == null ? <span className="empty">—</span> : String(row[column])}</td>)}</tr>)}</tbody></table>{!visiblePreview.length && <div className="no-rows">No preview rows match these filters.</div>}</div>
        <p className="preview-note">Filtering applies to the exported file too. The preview shows up to 100 rows; export includes the full target dataset.</p>
        <div className="section-heading download-heading"><div><span className="step">06</span><div><h2>Download Output</h2><p>Download the enriched target workbook.</p></div></div></div>
      </section>}
      <footer><div className="brand small-brand"><div className="brand-icon"><FileSpreadsheet size={17}/></div><strong>SheetFlow</strong></div><span>Files are processed locally on your server</span></footer>
    </main>
    {busy && <div className="busy-overlay"><RefreshCw className="spin"/><strong>Working on your spreadsheet…</strong><span>Large files may take a little time.</span></div>}
  </div>
}

function UploadCard({ title, subtitle, entry, onFile, onRemove, onSheet, busy }) {
  return <div className={`upload-card ${entry ? 'has-file' : ''}`}><div className="upload-card-head"><div><h3>{title}</h3><p>{subtitle}</p></div>{entry && <button className="icon-button" onClick={onRemove} title="Remove file"><Trash2 size={16}/></button>}</div>
    {!entry ? <label className="dropzone"><input type="file" accept=".xlsx,.xls,.xlsm,.csv" disabled={busy} onChange={e => { if (e.target.files?.[0]) onFile(e.target.files[0]); e.target.value = '' }}/><UploadCloud size={27}/><strong>Choose a file</strong><span>Browse · XLSX, XLS, XLSM, CSV</span></label> : <>
      <div className="file-summary"><div className="file-icon"><FileSpreadsheet size={22}/></div><div className="file-info"><strong>{entry.filename}</strong><span>{entry.row_count.toLocaleString()} rows · {entry.columns.length} columns · {entry.size ? `${(entry.size / 1024 / 1024).toFixed(2)} MB` : ''}</span></div><CheckCircle2 className="file-check" size={19}/></div>
      {entry.sheets.length > 1 && <label className="sheet-select"><span>Worksheet</span><select value={entry.selected_sheet} onChange={e => onSheet(e.target.value)}>{entry.sheets.map(sheet => <option key={sheet}>{sheet}</option>)}</select></label>}
      <div className="column-chips">{entry.columns.map(column => <span key={column}>{column}</span>)}</div>
    </>}
  </div>
}

export default App
