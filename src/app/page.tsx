'use client'

import { useState, useEffect, useRef, useCallback } from 'react'
import * as XLSX from 'xlsx'
import {
  CatalogItem,
  parseXmlFeed,
  parseWeberExcel,
  parseBroilKingExcel,
  mergeExcelIntoCatalog,
  generateOffersJson,
  formatPrice,
  ExcelRow,
} from '@/lib/catalog'

const XML_FEED_URL = 'https://kaminhall-feed.vercel.app/content.xml'
const APP_PASSWORD = process.env.NEXT_PUBLIC_APP_PASSWORD ?? 'kaminhall2024'
const STORAGE_KEY = 'kaminhall_catalog_v1'

// ─── Password Gate ────────────────────────────────────────────────────────────

function PasswordGate({ onAuth }: { onAuth: () => void }) {
  const [value, setValue] = useState('')
  const [error, setError] = useState(false)

  const submit = () => {
    if (value === APP_PASSWORD) {
      sessionStorage.setItem('kh_auth', '1')
      onAuth()
    } else {
      setError(true)
      setValue('')
      setTimeout(() => setError(false), 1500)
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center">
      <div className="w-80 space-y-4">
        <div className="text-center mb-8">
          <div className="text-2xl font-semibold text-white mb-1">КАМІНХОЛ</div>
          <div className="text-sm text-neutral-500">Управління пропозиціями</div>
        </div>
        <input
          type="password"
          placeholder="Пароль"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && submit()}
          className={`w-full bg-neutral-900 border rounded-lg px-4 py-3 text-white outline-none transition-colors
            ${error ? 'border-red-500' : 'border-neutral-700 focus:border-neutral-400'}`}
          autoFocus
        />
        {error && <p className="text-red-400 text-sm text-center">Невірний пароль</p>}
        <button
          onClick={submit}
          className="w-full bg-white text-black rounded-lg py-3 font-medium hover:bg-neutral-200 transition-colors"
        >
          Увійти
        </button>
      </div>
    </div>
  )
}

// ─── Upload Status Badge ──────────────────────────────────────────────────────

function UploadBadge({ matched, total, vendor }: { matched: number; total: number; vendor: string }) {
  return (
    <span className="inline-flex items-center gap-1.5 text-xs px-2.5 py-1 rounded-full bg-yellow-400/10 text-yellow-400 border border-yellow-400/20">
      <span className="w-1.5 h-1.5 rounded-full bg-yellow-400 inline-block" />
      {vendor}: {matched}/{total} товарів
    </span>
  )
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function Home() {
  const [authed, setAuthed] = useState(false)
  const [catalog, setCatalog] = useState<CatalogItem[]>([])
  const [loading, setLoading] = useState(false)
  const [loadError, setLoadError] = useState('')
  const [publishing, setPublishing] = useState(false)
  const [publishResult, setPublishResult] = useState<{ ok: boolean; msg: string } | null>(null)
  const [uploadStats, setUploadStats] = useState<{ vendor: string; matched: number; total: number }[]>([])
  const [search, setSearch] = useState('')
  const [filterBrand, setFilterBrand] = useState<'all' | 'Broil King' | 'Weber'>('all')
  const [filterChanged, setFilterChanged] = useState(false)
  const bkInputRef = useRef<HTMLInputElement>(null)
  const weberInputRef = useRef<HTMLInputElement>(null)

  // Auth check
  useEffect(() => {
    if (sessionStorage.getItem('kh_auth') === '1') setAuthed(true)
  }, [])

  // Load catalog from localStorage or fetch XML
  useEffect(() => {
    if (!authed) return
    const saved = localStorage.getItem(STORAGE_KEY)
    if (saved) {
      try {
        setCatalog(JSON.parse(saved))
        return
      } catch { /* fall through to fetch */ }
    }
    fetchXml()
  }, [authed])

  // Persist catalog to localStorage whenever it changes
  useEffect(() => {
    if (catalog.length > 0) {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(catalog))
    }
  }, [catalog])

  const fetchXml = useCallback(async () => {
    setLoading(true)
    setLoadError('')
    try {
      const res = await fetch(XML_FEED_URL)
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      const text = await res.text()
      const items = parseXmlFeed(text)
      setCatalog(items)
    } catch (e) {
      setLoadError(`Помилка завантаження XML: ${e}`)
    } finally {
      setLoading(false)
    }
  }, [])

  // ─── Excel Upload ───────────────────────────────────────────────────────────

  const handleExcelUpload = useCallback(
    async (file: File, vendor: 'weber' | 'broil-king') => {
      const buffer = await file.arrayBuffer()
      const wb = XLSX.read(buffer, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, {
        header: 1,
        defval: null,
      })

      let excelRows: ExcelRow[]
      if (vendor === 'weber') {
        excelRows = parseWeberExcel(rows as (string | number | null)[][])
      } else {
        excelRows = parseBroilKingExcel(rows as (string | number | null)[][])
      }

      const { items, matchedCount } = mergeExcelIntoCatalog(catalog, excelRows)
      setCatalog(items)
      setUploadStats((prev) => {
        const filtered = prev.filter((s) => s.vendor !== (vendor === 'weber' ? 'Weber' : 'Broil King'))
        return [...filtered, {
          vendor: vendor === 'weber' ? 'Weber' : 'Broil King',
          matched: matchedCount,
          total: excelRows.length,
        }]
      })
    },
    [catalog]
  )

  // ─── Inline Edit ────────────────────────────────────────────────────────────

  const updateItem = useCallback((code: string, field: keyof CatalogItem, value: unknown) => {
    setCatalog((prev) =>
      prev.map((item) =>
        item.code === code
          ? { ...item, [field]: value, changed: true, source: 'manual' }
          : item
      )
    )
  }, [])

  // ─── Publish ────────────────────────────────────────────────────────────────

  const handlePublish = async () => {
    setPublishing(true)
    setPublishResult(null)
    try {
      const offersJson = generateOffersJson(catalog)
      const res = await fetch('/api/publish', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ content: offersJson }),
      })
      const data = await res.json()
      if (data.success) {
        setPublishResult({ ok: true, msg: `Опубліковано ${offersJson.total} пропозицій` })
        // Clear changed flags
        setCatalog((prev) => prev.map((i) => ({ ...i, changed: false })))
      } else {
        setPublishResult({ ok: false, msg: data.error ?? 'Помилка публікації' })
      }
    } catch (e) {
      setPublishResult({ ok: false, msg: String(e) })
    } finally {
      setPublishing(false)
    }
  }

  // ─── Filtered catalog ───────────────────────────────────────────────────────

  const filtered = catalog.filter((item) => {
    if (filterBrand !== 'all' && item.brand !== filterBrand) return false
    if (filterChanged && !item.changed) return false
    if (search) {
      const q = search.toLowerCase()
      return (
        item.title.toLowerCase().includes(q) ||
        item.code.toLowerCase().includes(q) ||
        item.vendor_code.toLowerCase().includes(q)
      )
    }
    return true
  })

  const changedCount = catalog.filter((i) => i.changed).length
  const withPriceCount = catalog.filter((i) => i.price !== null).length

  if (!authed) return <PasswordGate onAuth={() => setAuthed(true)} />

  return (
    <div className="min-h-screen p-4 md:p-6">
      {/* Header */}
      <div className="flex items-center justify-between mb-6 flex-wrap gap-3">
        <div>
          <h1 className="text-xl font-semibold text-white">КАМІНХОЛ — Пропозиції</h1>
          <p className="text-sm text-neutral-500 mt-0.5">
            {catalog.length} товарів · {withPriceCount} з цінами · {changedCount} змінено
          </p>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {uploadStats.map((s) => (
            <UploadBadge key={s.vendor} {...s} />
          ))}

          {/* Excel uploads */}
          <input
            ref={bkInputRef}
            type="file"
            accept=".xls,.xlsx"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0]
              if (f) handleExcelUpload(f, 'broil-king')
              e.target.value = ''
            }}
          />
          <input
            ref={weberInputRef}
            type="file"
            accept=".xls,.xlsx"
            className="hidden"
            onChange={(e) => {
              const f = e.target.files?.[0]
              if (f) handleExcelUpload(f, 'weber')
              e.target.value = ''
            }}
          />

          <button
            onClick={() => bkInputRef.current?.click()}
            className="text-sm px-3 py-1.5 rounded-lg border border-neutral-700 text-neutral-300 hover:border-neutral-500 hover:text-white transition-colors"
          >
            ↑ Broil King XLS
          </button>
          <button
            onClick={() => weberInputRef.current?.click()}
            className="text-sm px-3 py-1.5 rounded-lg border border-neutral-700 text-neutral-300 hover:border-neutral-500 hover:text-white transition-colors"
          >
            ↑ Weber XLSX
          </button>

          <button
            onClick={() => { if (window.confirm('Ви впевнені що треба скинути XML файл? Подумай добре!\n\nВсі ціни та залишки будуть очищені.')) { fetchXml() } }}
            disabled={loading}
            className="text-sm px-3 py-1.5 rounded-lg border border-neutral-700 text-neutral-400 hover:border-neutral-500 transition-colors"
            title="Перезавантажити XML каталог"
          >
            {loading ? '…' : '↻ XML'}
          </button>

          <button
            onClick={handlePublish}
            disabled={publishing || withPriceCount === 0}
            className="text-sm px-4 py-1.5 rounded-lg bg-white text-black font-medium hover:bg-neutral-200 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            {publishing ? 'Публікую…' : `Опублікувати (${withPriceCount})`}
          </button>
        </div>
      </div>

      {/* Alerts */}
      {loadError && (
        <div className="mb-4 px-4 py-3 rounded-lg bg-red-500/10 border border-red-500/20 text-red-400 text-sm">
          {loadError}
        </div>
      )}
      {publishResult && (
        <div
          className={`mb-4 px-4 py-3 rounded-lg text-sm border flex items-center justify-between
            ${publishResult.ok
              ? 'bg-green-500/10 border-green-500/20 text-green-400'
              : 'bg-red-500/10 border-red-500/20 text-red-400'}`}
        >
          <span>{publishResult.msg}</span>
          <button onClick={() => setPublishResult(null)} className="ml-4 opacity-60 hover:opacity-100">✕</button>
        </div>
      )}

      {/* Filters */}
      <div className="flex items-center gap-3 mb-4 flex-wrap">
        <input
          type="text"
          placeholder="Пошук за назвою, кодом, артикулом…"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="bg-neutral-900 border border-neutral-700 rounded-lg px-3 py-1.5 text-sm text-white outline-none focus:border-neutral-500 w-72"
        />
        <select
          value={filterBrand}
          onChange={(e) => setFilterBrand(e.target.value as typeof filterBrand)}
          className="bg-neutral-900 border border-neutral-700 rounded-lg px-3 py-1.5 text-sm text-neutral-300 outline-none"
        >
          <option value="all">Всі бренди</option>
          <option value="Broil King">Broil King</option>
          <option value="Weber">Weber</option>
        </select>
        <label className="flex items-center gap-2 text-sm text-neutral-400 cursor-pointer select-none">
          <input
            type="checkbox"
            checked={filterChanged}
            onChange={(e) => setFilterChanged(e.target.checked)}
            className="accent-yellow-400"
          />
          Тільки змінені
        </label>
        <span className="text-xs text-neutral-600">{filtered.length} з {catalog.length}</span>
      </div>

      {/* Table */}
      {loading ? (
        <div className="text-center text-neutral-500 py-20">Завантаження каталогу…</div>
      ) : catalog.length === 0 ? (
        <div className="text-center text-neutral-600 py-20">
          <p>Каталог порожній</p>
          <button onClick={fetchXml} className="mt-3 text-sm text-neutral-400 underline">
            Завантажити XML
          </button>
        </div>
      ) : (
        <div className="overflow-x-auto rounded-xl border border-neutral-800">
          <table className="w-full text-sm border-collapse">
            <thead>
              <tr className="border-b border-neutral-800">
                {['Код', 'Артикул', 'Бренд', 'Назва', 'Ціна', 'Стара ціна', 'Залишок', 'Наявність', 'Дж.'].map((h) => (
                  <th key={h} className="px-3 py-2.5 text-left text-xs font-medium text-neutral-500 whitespace-nowrap bg-neutral-900/80">
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((item) => (
                <tr
                  key={item.code}
                  className={`border-b border-neutral-800/60 hover:bg-neutral-800/30 transition-colors
                    ${item.changed ? 'changed-row' : ''}`}
                >
                  <td className="px-3 py-2 text-neutral-400 font-mono text-xs whitespace-nowrap">{item.code}</td>
                  <td className="px-3 py-2 text-neutral-500 font-mono text-xs whitespace-nowrap">{item.vendor_code}</td>
                  <td className="px-3 py-2 whitespace-nowrap">
                    <span className={`text-xs px-2 py-0.5 rounded-full
                      ${item.brand === 'Broil King'
                        ? 'bg-orange-500/10 text-orange-400'
                        : 'bg-blue-500/10 text-blue-400'}`}>
                      {item.brand === 'Broil King' ? 'BK' : 'W'}
                    </span>
                  </td>
                  <td className="px-3 py-2 text-neutral-200 max-w-xs">
                    <span className="block truncate" title={item.title}>{item.title}</span>
                  </td>

                  {/* Editable: price */}
                  <td className="px-3 py-2 whitespace-nowrap">
                    <input
                      type="number"
                      value={item.price ?? ''}
                      onChange={(e) => updateItem(item.code, 'price', e.target.value ? Number(e.target.value) : null)}
                      placeholder="—"
                      className="w-24 bg-transparent border-b border-neutral-700 focus:border-neutral-400 outline-none text-white text-right py-0.5 px-1"
                    />
                  </td>

                  {/* Editable: old_price */}
                  <td className="px-3 py-2 whitespace-nowrap">
                    <input
                      type="number"
                      value={item.old_price ?? ''}
                      onChange={(e) => updateItem(item.code, 'old_price', e.target.value ? Number(e.target.value) : null)}
                      placeholder="—"
                      className="w-24 bg-transparent border-b border-neutral-700 focus:border-neutral-400 outline-none text-neutral-400 text-right py-0.5 px-1"
                    />
                  </td>

                  {/* Editable: stock */}
                  <td className="px-3 py-2 whitespace-nowrap">
                    <input
                      type="number"
                      value={item.stock ?? ''}
                      onChange={(e) => updateItem(item.code, 'stock', e.target.value ? Number(e.target.value) : null)}
                      placeholder="—"
                      className="w-16 bg-transparent border-b border-neutral-700 focus:border-neutral-400 outline-none text-white text-right py-0.5 px-1"
                    />
                  </td>

                  {/* Editable: availability toggle */}
                  <td className="px-3 py-2">
                    <button
                      onClick={() => updateItem(item.code, 'availability', !item.availability)}
                      className={`text-xs px-2 py-0.5 rounded-full border transition-colors
                        ${item.availability
                          ? 'bg-green-500/10 border-green-500/20 text-green-400'
                          : 'bg-red-500/10 border-red-500/20 text-red-400'}`}
                    >
                      {item.availability ? 'Є' : 'Нема'}
                    </button>
                  </td>

                  {/* Source indicator */}
                  <td className="px-3 py-2 text-center">
                    {item.source === 'excel' && (
                      <span className="text-xs text-yellow-500" title="З Excel">E</span>
                    )}
                    {item.source === 'manual' && (
                      <span className="text-xs text-blue-400" title="Ручне редагування">M</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Footer */}
      <div className="mt-4 text-xs text-neutral-700 text-center">
        Дані зберігаються локально в браузері · Зміни публікуються через GitHub API
      </div>
    </div>
  )
}
