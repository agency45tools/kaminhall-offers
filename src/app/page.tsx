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
  ExcelRow,
  ALL_DELIVERY_METHODS,
  DEFAULT_DELIVERY_METHODS,
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
    <div className="min-h-screen flex items-center justify-center bg-[#f5f5f7]">
      <div className="w-80 space-y-4">
        <div className="text-center mb-10">
          <div className="inline-flex items-center justify-center w-14 h-14 rounded-2xl bg-violet-100 mb-4">
            <span className="text-2xl">🔥</span>
          </div>
          <div className="text-2xl font-bold text-gray-900 tracking-tight">КАМІНХОЛ</div>
          <div className="text-sm text-gray-400 mt-1">Управління пропозиціями</div>
        </div>
        <input
          type="password"
          placeholder="Пароль"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && submit()}
          className={`w-full bg-white border rounded-xl px-4 py-3 text-gray-900 outline-none transition-all shadow-sm
            ${error ? 'border-red-400' : 'border-gray-200 focus:border-violet-400 focus:ring-2 focus:ring-violet-100'}`}
          autoFocus
        />
        {error && <p className="text-red-500 text-sm text-center">Невірний пароль</p>}
        <button
          onClick={submit}
          className="w-full bg-violet-600 hover:bg-violet-700 text-white rounded-xl py-3 font-medium transition-colors shadow-sm"
        >
          Увійти
        </button>
      </div>
    </div>
  )
}

// ─── Upload Badge ─────────────────────────────────────────────────────────────

function UploadBadge({ matched, total, vendor }: { matched: number; total: number; vendor: string }) {
  return (
    <span className="inline-flex items-center gap-1.5 text-xs px-2.5 py-1 rounded-full bg-amber-50 text-amber-600 border border-amber-200 font-medium">
      <span className="w-1.5 h-1.5 rounded-full bg-amber-400 inline-block" />
      {vendor}: {matched}/{total}
    </span>
  )
}

// ─── Delivery Popup ───────────────────────────────────────────────────────────

function DeliveryPopup({ methods, onUpdate }: {
  methods: string[]
  onUpdate: (updated: string[]) => void
}) {
  const [open, setOpen] = useState(false)
  const ref = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false)
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [])

  return (
    <div ref={ref} className="relative">
      <button
        onClick={() => setOpen((v) => !v)}
        className={`text-xs px-2.5 py-1 rounded-lg border font-medium transition-colors
          ${methods.length === 4
            ? 'border-gray-200 text-gray-500 bg-gray-50 hover:border-gray-300'
            : 'border-amber-300 text-amber-600 bg-amber-50'}`}>
        {methods.length}/4
      </button>
      {open && (
        <div className="absolute left-0 top-8 z-50 bg-white border border-gray-200 rounded-xl p-3 shadow-xl min-w-max">
          {ALL_DELIVERY_METHODS.map((dm) => (
            <label key={dm.method} className="flex items-center gap-2.5 text-xs text-gray-600 py-1.5 cursor-pointer hover:text-gray-900 transition-colors">
              <input
                type="checkbox"
                checked={methods.includes(dm.method)}
                onChange={(e) => {
                  const updated = e.target.checked
                    ? [...methods, dm.method]
                    : methods.filter((m) => m !== dm.method)
                  onUpdate(updated)
                }}
                className="accent-violet-600"
              />
              {dm.label}
            </label>
          ))}
        </div>
      )}
    </div>
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

  useEffect(() => {
    if (sessionStorage.getItem('kh_auth') === '1') setAuthed(true)
  }, [])

  useEffect(() => {
    if (!authed) return
    const saved = localStorage.getItem(STORAGE_KEY)
    if (saved) {
      try { setCatalog(JSON.parse(saved)); return } catch { /* fall through */ }
    }
    fetchXml()
  }, [authed])

  useEffect(() => {
    if (catalog.length > 0) localStorage.setItem(STORAGE_KEY, JSON.stringify(catalog))
  }, [catalog])

  const fetchXml = useCallback(async () => {
    setLoading(true)
    setLoadError('')
    try {
      const res = await fetch(XML_FEED_URL)
      if (!res.ok) throw new Error(`HTTP ${res.status}`)
      const text = await res.text()
      const items = parseXmlFeed(text)

      try {
        const offersRes = await fetch('https://kaminhall-feed.vercel.app/offers.json?t=' + Date.now())
        if (offersRes.ok) {
          const offersData = await offersRes.json()
          type OfferEntry = { code: string; price: number; old_price: number | null; stock: number; availability: boolean; days_to_dispatch?: number; max_pay_in_parts?: number; delivery_methods?: { method: string }[] }
          const offersMap = new Map<string, OfferEntry>(offersData.data.map((o: OfferEntry) => [o.code, o]))
          items.forEach((item) => {
            const offer = offersMap.get(item.code)
            if (offer) {
              item.price = offer.price
              item.old_price = offer.old_price
              item.stock = offer.stock
              item.availability = offer.availability
              item.days_to_dispatch = offer.days_to_dispatch ?? 1
              item.max_pay_in_parts = offer.max_pay_in_parts ?? 12
              item.delivery_methods = offer.delivery_methods ? offer.delivery_methods.map((d) => d.method) : DEFAULT_DELIVERY_METHODS
              item.source = 'excel'
            }
          })
        }
      } catch { /* skip */ }

      setCatalog(items)
    } catch (e) {
      setLoadError(`Помилка завантаження XML: ${e}`)
    } finally {
      setLoading(false)
    }
  }, [])

  const handleExcelUpload = useCallback(async (file: File, vendor: 'weber' | 'broil-king') => {
    const buffer = await file.arrayBuffer()
    const wb = XLSX.read(buffer, { type: 'array' })
    const ws = wb.Sheets[wb.SheetNames[0]]
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(ws, { header: 1, defval: null })
    const excelRows: ExcelRow[] = vendor === 'weber'
      ? parseWeberExcel(rows as (string | number | null)[][])
      : parseBroilKingExcel(rows as (string | number | null)[][])
    const { items, matchedCount } = mergeExcelIntoCatalog(catalog, excelRows, vendor)
    setCatalog(items)
    setUploadStats((prev) => {
      const label = vendor === 'weber' ? 'Weber' : 'Broil King'
      return [...prev.filter((s) => s.vendor !== label), { vendor: label, matched: matchedCount, total: excelRows.length }]
    })
  }, [catalog])

  const updateItem = useCallback((code: string, field: keyof CatalogItem, value: unknown) => {
    setCatalog((prev) => prev.map((item) =>
      item.code === code ? { ...item, [field]: value, changed: true, source: 'manual' } : item
    ))
  }, [])

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

  const filtered = catalog.filter((item) => {
    if (filterBrand !== 'all' && item.brand !== filterBrand) return false
    if (filterChanged && !item.changed) return false
    if (search) {
      const q = search.toLowerCase()
      return item.title.toLowerCase().includes(q) || item.code.toLowerCase().includes(q) || item.vendor_code.toLowerCase().includes(q)
    }
    return true
  })

  const changedCount = catalog.filter((i) => i.changed).length
  const withPriceCount = catalog.filter((i) => i.price !== null && i.price > 0).length

  if (!authed) return <PasswordGate onAuth={() => setAuthed(true)} />

  return (
    <div className="min-h-screen bg-[#f5f5f7]">
      {/* Header */}
      <div className="bg-white border-b border-gray-200 sticky top-0 z-40 shadow-sm">
        <div className="px-6 py-4 flex items-center justify-between gap-4 flex-wrap">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-violet-100 flex items-center justify-center text-lg">🔥</div>
            <div>
              <div className="text-base font-bold text-gray-900 leading-tight tracking-tight">КАМІНХОЛ</div>
              <div className="text-xs text-gray-400 leading-tight">
                {catalog.length} товарів · <span className="text-violet-600 font-medium">{withPriceCount} з цінами</span>
                {changedCount > 0 && <> · <span className="text-amber-500 font-medium">{changedCount} змінено</span></>}
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            {uploadStats.map((s) => <UploadBadge key={s.vendor} {...s} />)}

            <input ref={bkInputRef} type="file" accept=".xls,.xlsx" className="hidden"
              onChange={(e) => { const f = e.target.files?.[0]; if (f) handleExcelUpload(f, 'broil-king'); e.target.value = '' }} />
            <input ref={weberInputRef} type="file" accept=".xls,.xlsx" className="hidden"
              onChange={(e) => { const f = e.target.files?.[0]; if (f) handleExcelUpload(f, 'weber'); e.target.value = '' }} />

            <button onClick={() => bkInputRef.current?.click()}
              className="text-xs px-3 py-2 rounded-lg border border-gray-200 text-gray-600 bg-white hover:border-orange-300 hover:text-orange-600 hover:bg-orange-50 transition-colors shadow-sm font-medium flex items-center gap-1.5">
              <span className="text-orange-500">↑</span> Broil King
            </button>
            <button onClick={() => weberInputRef.current?.click()}
              className="text-xs px-3 py-2 rounded-lg border border-gray-200 text-gray-600 bg-white hover:border-blue-300 hover:text-blue-600 hover:bg-blue-50 transition-colors shadow-sm font-medium flex items-center gap-1.5">
              <span className="text-blue-500">↑</span> Weber
            </button>
            <button
              onClick={() => { if (window.confirm('Ви впевнені що треба скинути XML файл?\n\nВсі ціни та залишки будуть очищені.')) { fetchXml() } }}
              disabled={loading}
              className="text-xs px-3 py-2 rounded-lg border border-gray-200 text-gray-500 bg-white hover:border-gray-300 hover:text-gray-700 transition-colors shadow-sm font-medium">
              {loading ? '…' : '↻ XML'}
            </button>
            <button
              onClick={handlePublish}
              disabled={publishing || withPriceCount === 0}
              className="text-xs px-4 py-2 rounded-lg bg-violet-600 hover:bg-violet-700 text-white font-semibold disabled:opacity-40 disabled:cursor-not-allowed transition-colors shadow-sm">
              {publishing ? 'Публікую…' : `Опублікувати (${withPriceCount})`}
            </button>
          </div>
        </div>
      </div>

      <div className="px-6 py-5">
        {/* Alerts */}
        {loadError && (
          <div className="mb-4 px-4 py-3 rounded-xl bg-red-50 border border-red-200 text-red-600 text-sm">{loadError}</div>
        )}
        {publishResult && (
          <div className={`mb-4 px-4 py-3 rounded-xl text-sm border flex items-center justify-between
            ${publishResult.ok ? 'bg-emerald-50 border-emerald-200 text-emerald-700' : 'bg-red-50 border-red-200 text-red-600'}`}>
            <span className="font-medium">{publishResult.msg}</span>
            <button onClick={() => setPublishResult(null)} className="ml-4 opacity-50 hover:opacity-100 text-lg leading-none">×</button>
          </div>
        )}

        {/* Filters */}
        <div className="flex items-center gap-3 mb-4 flex-wrap">
          <div className="relative">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm">⌕</span>
            <input
              type="text"
              placeholder="Пошук за назвою, кодом, артикулом…"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="bg-white border border-gray-200 rounded-xl pl-8 pr-4 py-2 text-sm text-gray-900 outline-none focus:border-violet-400 focus:ring-2 focus:ring-violet-100 w-72 transition-all shadow-sm"
            />
          </div>
          <div className="flex rounded-xl border border-gray-200 overflow-hidden bg-white shadow-sm">
            {(['all', 'Broil King', 'Weber'] as const).map((b) => (
              <button key={b}
                onClick={() => setFilterBrand(b)}
                className={`text-xs px-4 py-2 font-medium transition-colors
                  ${filterBrand === b ? 'bg-violet-600 text-white' : 'text-gray-500 hover:text-gray-800 hover:bg-gray-50'}`}>
                {b === 'all' ? 'Всі' : b === 'Broil King' ? 'Broil King' : 'Weber'}
              </button>
            ))}
          </div>
          <label className="flex items-center gap-2 text-xs text-gray-500 cursor-pointer select-none hover:text-gray-700 transition-colors font-medium">
            <input type="checkbox" checked={filterChanged} onChange={(e) => setFilterChanged(e.target.checked)} className="accent-violet-600" />
            Тільки змінені
          </label>
          <span className="text-xs text-gray-400">{filtered.length} / {catalog.length}</span>
        </div>

        {/* Table */}
        {loading ? (
          <div className="text-center text-gray-400 py-24 text-sm">Завантаження каталогу…</div>
        ) : catalog.length === 0 ? (
          <div className="text-center text-gray-400 py-24">
            <p className="mb-3">Каталог порожній</p>
            <button onClick={fetchXml} className="text-sm text-violet-600 hover:text-violet-700 font-medium underline">Завантажити XML</button>
          </div>
        ) : (
          <div className="rounded-2xl border border-gray-200 overflow-hidden shadow-sm bg-white">
            <div className="overflow-x-auto">
              <table className="w-full text-sm border-collapse">
                <thead>
                  <tr className="border-b border-gray-100 bg-gray-50">
                    {['Код', 'Артикул', 'Бренд', 'Назва', 'Ціна', 'Стара ціна', 'Залишок', 'Дні', 'Частини', 'Доставка', 'Наявність', 'Дж.'].map((h) => (
                      <th key={h} className="px-4 py-3 text-left text-xs font-semibold text-gray-400 whitespace-nowrap uppercase tracking-wider">
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {filtered.map((item) => (
                    <tr key={item.code}
                      className={`transition-colors hover:bg-gray-50 ${item.changed ? 'bg-amber-50/60 border-l-2 border-l-amber-400' : ''}`}>

                      <td className="px-4 py-3 font-mono text-xs text-gray-400 whitespace-nowrap">{item.code}</td>
                      <td className="px-4 py-3 font-mono text-xs text-gray-300 whitespace-nowrap">{item.vendor_code}</td>

                      <td className="px-4 py-3 whitespace-nowrap">
                        <span className={`text-xs font-semibold px-2.5 py-1 rounded-lg
                          ${item.brand === 'Broil King'
                            ? 'bg-orange-100 text-orange-600'
                            : 'bg-blue-100 text-blue-600'}`}>
                          {item.brand === 'Broil King' ? 'BK' : 'W'}
                        </span>
                      </td>

                      <td className="px-4 py-3 text-gray-800 max-w-xs">
                        <span className="block truncate text-sm font-medium" title={item.title}>{item.title}</span>
                      </td>

                      <td className="px-4 py-3 whitespace-nowrap">
                        <input type="number" value={item.price ?? ''} placeholder="—"
                          onChange={(e) => updateItem(item.code, 'price', e.target.value ? Number(e.target.value) : null)}
                          className="w-24 bg-transparent border-b-2 border-gray-200 focus:border-violet-400 outline-none text-gray-900 text-right py-0.5 text-sm font-medium transition-colors" />
                      </td>

                      <td className="px-4 py-3 whitespace-nowrap">
                        <input type="number" value={item.old_price ?? ''} placeholder="—"
                          onChange={(e) => updateItem(item.code, 'old_price', e.target.value ? Number(e.target.value) : null)}
                          className="w-24 bg-transparent border-b-2 border-gray-200 focus:border-violet-400 outline-none text-gray-400 text-right py-0.5 text-sm transition-colors" />
                      </td>

                      <td className="px-4 py-3 whitespace-nowrap">
                        <input type="number" value={item.stock ?? ''} placeholder="—"
                          onChange={(e) => updateItem(item.code, 'stock', e.target.value ? Number(e.target.value) : null)}
                          className="w-16 bg-transparent border-b-2 border-gray-200 focus:border-violet-400 outline-none text-gray-900 text-right py-0.5 text-sm font-medium transition-colors" />
                      </td>

                      <td className="px-4 py-3 whitespace-nowrap">
                        <input type="number" min="0" max="30" value={item.days_to_dispatch ?? 1}
                          onChange={(e) => updateItem(item.code, 'days_to_dispatch', Number(e.target.value))}
                          className="w-12 bg-transparent border-b-2 border-gray-200 focus:border-violet-400 outline-none text-gray-900 text-right py-0.5 text-sm font-medium transition-colors" />
                      </td>

                      <td className="px-4 py-3 whitespace-nowrap">
                        <input type="number" min="1" max="36" value={item.max_pay_in_parts ?? 12}
                          onChange={(e) => updateItem(item.code, 'max_pay_in_parts', Number(e.target.value))}
                          className="w-12 bg-transparent border-b-2 border-gray-200 focus:border-violet-400 outline-none text-gray-900 text-right py-0.5 text-sm font-medium transition-colors" />
                      </td>

                      <td className="px-4 py-3 whitespace-nowrap">
                        <DeliveryPopup
                          methods={item.delivery_methods ?? DEFAULT_DELIVERY_METHODS}
                          onUpdate={(updated) => updateItem(item.code, 'delivery_methods', updated)}
                        />
                      </td>

                      <td className="px-4 py-3 whitespace-nowrap">
                        <span className={`text-xs font-semibold px-2.5 py-1 rounded-lg
                          ${(item.stock !== null && item.stock > 0)
                            ? 'bg-emerald-100 text-emerald-700'
                            : 'bg-red-100 text-red-600'}`}>
                          {(item.stock !== null && item.stock > 0) ? 'Є' : 'Нема'}
                        </span>
                      </td>

                      <td className="px-4 py-3 text-center">
                        {item.source === 'excel' && <span className="text-xs text-amber-500 font-bold" title="З Excel">E</span>}
                        {item.source === 'manual' && <span className="text-xs text-violet-500 font-bold" title="Ручне">M</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div className="mt-6 text-xs text-gray-300 text-center">
          Дані синхронізуються з offers.json · Публікація через GitHub API → Vercel
        </div>
      </div>
    </div>
  )
}
