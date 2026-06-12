// ─── Types ───────────────────────────────────────────────────────────────────

export type Vendor = 'broil-king' | 'weber'

export const ALL_DELIVERY_METHODS = [
  { method: 'nova-post:branch', label: 'НП відділення' },
  { method: 'nova-post:postomat', label: 'НП поштомат' },
  { method: 'nova-post:cargo_branch', label: 'НП вантажне' },
  { method: 'courier:nova-post', label: "Кур'єр НП" },
]

export const DEFAULT_DELIVERY_METHODS = ALL_DELIVERY_METHODS.map(m => m.method)

export interface CatalogItem {
  // From XML
  id: string
  code: string
  vendor_code: string
  title: string
  brand: string
  warranty_period: number
  country_code: string

  // From Excel / manual edit
  price: number | null
  old_price: number | null
  stock: number | null
  availability: boolean

  // Manually controlled
  days_to_dispatch: number
  max_pay_in_parts: number
  delivery_methods: string[]

  // UI state
  changed?: boolean
  source?: 'excel' | 'manual'
}

export interface OfferItem {
  code: string
  price: number
  old_price: number | null
  availability: boolean
  stock: number
  warranty_type: 'manufacturer' | 'merchant' | 'no'
  warranty_period: number
  days_to_dispatch: number
  max_pay_in_parts: number
  delivery_methods: { method: string; price: number }[]
  manufacture: { country_code: string; year: number | null } | null
  warehouses: { id: string; stock: number }[]
}

export interface OffersJson {
  total: number
  updatedAt: string
  data: OfferItem[]
}

// ─── XML Parser ──────────────────────────────────────────────────────────────

export function parseXmlFeed(xmlText: string): CatalogItem[] {
  const parser = new DOMParser()
  const doc = parser.parseFromString(xmlText, 'application/xml')
  const offers = doc.querySelectorAll('offer')
  const items: CatalogItem[] = []

  offers.forEach((offer) => {
    const get = (tag: string) => offer.querySelector(tag)?.textContent?.trim() ?? ''

    // Parse warranty from <param name="Гарантія">12 місяців</param>
    let warrantyPeriod = 24 // default
    const warrantyParam = Array.from(offer.querySelectorAll('param')).find(
      (p) => p.getAttribute('name') === 'Гарантія'
    )
    if (warrantyParam) {
      const match = warrantyParam.textContent?.match(/\d+/)
      if (match) warrantyPeriod = parseInt(match[0])
    }

    // Parse country from <param name="Країна виробництва">
    let countryCode = 'US' // default for BK/Weber
    const countryParam = Array.from(offer.querySelectorAll('param')).find(
      (p) => p.getAttribute('name') === 'Країна виробництва'
    )
    if (countryParam) {
      const countryMap: Record<string, string> = {
        'США': 'US', 'Канада': 'CA', 'Китай': 'CN',
        'Чехія': 'CZ', 'Німеччина': 'DE', 'Австрія': 'AT',
        'Великобританія': 'GB', 'Франція': 'FR', 'Польща': 'PL',
      }
      const raw = countryParam.textContent?.trim() ?? ''
      countryCode = countryMap[raw] ?? 'US'
    }

    const code = get('code')
    const vendor_code = get('vendor_code')
    if (!code || !vendor_code) return

    items.push({
      id: get('id'),
      code,
      vendor_code,
      title: get('title'),
      brand: get('brand'),
      warranty_period: warrantyPeriod,
      country_code: countryCode,
      price: null,
      old_price: null,
      stock: null,
      availability: true,
      days_to_dispatch: 1,
      max_pay_in_parts: 12,
      delivery_methods: DEFAULT_DELIVERY_METHODS,
    })
  })

  return items
}

// ─── Excel Parser ─────────────────────────────────────────────────────────────

export interface ExcelRow {
  vendor_code: string
  price: number
  old_price: number | null
  stock: number
}

/**
 * Parse Weber XLSX:
 * Row 0: group headers (Товар.Группа ... Залишок)
 * Row 1: column headers (Товар | Артикул | Новинка | Ціна | Акційна ціна | Залишок)
 * Rows 2+: data (skip rows where col[1] is empty — those are category headers)
 */
export function parseWeberExcel(rows: (string | number | null)[][]): ExcelRow[] {
  const result: ExcelRow[] = []

  for (let i = 2; i < rows.length; i++) {
    const row = rows[i]
    const vendorCode = row[1]
    if (!vendorCode || String(vendorCode).trim() === '') continue

    const priceRaw = row[3]
    const salePriceRaw = row[4]
    const stockRaw = row[5]

    const basePrice = typeof priceRaw === 'number' ? priceRaw : parseFloat(String(priceRaw ?? '0'))
    const salePrice = typeof salePriceRaw === 'number' ? salePriceRaw
      : salePriceRaw ? parseFloat(String(salePriceRaw)) : null

    if (!basePrice) continue

    result.push({
      vendor_code: String(vendorCode).trim(),
      price: salePrice ?? basePrice,
      old_price: salePrice ? basePrice : null,
      stock: typeof stockRaw === 'number' ? stockRaw : parseInt(String(stockRaw ?? '0')) || 0,
    })
  }

  return result
}

/**
 * Parse Broil King XLS:
 * Rows 0-1: empty
 * Row 2: headers (Артикул | Номенклатура | Группа верхняя | Единица хранения | Залишок | Ціна)
 * Row 3: empty divider
 * Rows 4+: data — filter by brand col[2] === 'Broil King'
 * NOTE: col[4] = залишок, col[5] = ціна (назва колонки оманлива)
 */
export function parseBroilKingExcel(rows: (string | number | null)[][]): ExcelRow[] {
  const result: ExcelRow[] = []

  for (let i = 4; i < rows.length; i++) {
    const row = rows[i]
    const vendorCode = row[0]
    const brand = String(row[2] ?? '').trim()

    // Only Broil King products (skip Big Green Egg, Cape Herb etc.)
    if (brand !== 'Broil King') continue
    if (!vendorCode || String(vendorCode).trim() === '') continue

    // col[4] = залишок, col[5] = ціна (колонки переплутані в заголовку)
    const stockRaw = row[4]
    const priceRaw = row[5]

    const price = typeof priceRaw === 'number' ? priceRaw : parseFloat(String(priceRaw ?? '0'))
    if (!price) continue

    result.push({
      vendor_code: String(vendorCode).trim(),
      price: Math.round(price),
      old_price: null,
      stock: typeof stockRaw === 'number' ? Math.round(stockRaw) : parseInt(String(stockRaw ?? '0')) || 0,
    })
  }

  return result
}

// ─── Catalog Merger ───────────────────────────────────────────────────────────

/**
 * Merge Excel rows into catalog items.
 * Returns updated items with `changed: true` on those that differ.
 */
export function mergeExcelIntoCatalog(
  catalog: CatalogItem[],
  excelRows: ExcelRow[]
): { items: CatalogItem[]; matchedCount: number; unmatchedCodes: string[] } {
  const excelMap = new Map<string, ExcelRow>()
  excelRows.forEach((r) => excelMap.set(String(r.vendor_code), r))

  const unmatchedCodes: string[] = []
  let matchedCount = 0

  const items = catalog.map((item) => {
    const match = excelMap.get(item.vendor_code)
    if (!match) return { ...item, changed: false }

    matchedCount++
    const changed =
      item.price !== match.price ||
      item.old_price !== match.old_price ||
      item.stock !== match.stock

    return {
      ...item,
      price: match.price,
      old_price: match.old_price,
      stock: match.stock,
      availability: match.stock > 0,
      changed,
      source: 'excel' as const,
    }
  })

  // Find unmatched Excel rows
  const catalogVendorCodes = new Set(catalog.map((i) => i.vendor_code))
  excelMap.forEach((_, vc) => {
    if (!catalogVendorCodes.has(vc)) unmatchedCodes.push(vc)
  })

  return { items, matchedCount, unmatchedCodes }
}

// ─── Offers JSON Generator ────────────────────────────────────────────────────

export function generateOffersJson(catalog: CatalogItem[]): OffersJson {
  const WAREHOUSE_ID = 'WH-01'

  // Only include items that have price set
  const activeItems = catalog.filter((item) => item.price !== null && item.price > 0)

  const data: OfferItem[] = activeItems.map((item) => ({
    code: item.code,
    price: Math.round(item.price!),
    old_price: item.old_price ? Math.round(item.old_price) : null,
    availability: item.stock !== null && item.stock > 0,
    stock: item.stock ?? 0,
    warranty_type: 'manufacturer',
    warranty_period: item.warranty_period,
    days_to_dispatch: item.days_to_dispatch ?? 1,
    max_pay_in_parts: item.max_pay_in_parts ?? 12,
    delivery_methods: (item.delivery_methods ?? DEFAULT_DELIVERY_METHODS).map(m => ({ method: m, price: 0 })),
    manufacture: item.country_code
      ? { country_code: item.country_code, year: null }
      : null,
    warehouses: [{ id: WAREHOUSE_ID, stock: item.stock ?? 0 }],
  }))

  return {
    total: data.length,
    updatedAt: new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'),
    data,
  }
}

// ─── Country name → ISO code ──────────────────────────────────────────────────

export function formatPrice(n: number | null): string {
  if (n === null) return '—'
  return new Intl.NumberFormat('uk-UA', { minimumFractionDigits: 0 }).format(n) + ' ₴'
}
