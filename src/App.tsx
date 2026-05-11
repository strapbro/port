import { useMemo, useState } from 'react'
import Papa from 'papaparse'
import * as XLSX from 'xlsx'
import {
  ArrowClockwise,
  ChartBar,
  Database,
  DownloadSimple,
  FileArrowUp,
  Funnel,
  MagnifyingGlass,
  PencilSimple,
  Pulse,
  Scales,
  ShieldWarning,
  Sparkle,
  StackSimple,
  Trash,
  UploadSimple,
  Warning,
} from '@phosphor-icons/react'
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  Treemap,
  XAxis,
  YAxis,
} from 'recharts'
import {
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getGroupedRowModel,
  getSortedRowModel,
  useReactTable,
} from '@tanstack/react-table'
import type { ColumnDef, GroupingState, SortingState } from '@tanstack/react-table'
import './index.css'

type AssetClass = 'Broad US equity' | 'AI buildout sleeve' | 'International equity' | 'Bonds/fixed income' | 'Cash' | 'Alternatives/other'
type Directness = 'direct' | 'indirect' | 'none'
type ClassificationSource = 'manual' | 'default' | 'imported' | 'unknown'
type WarningSeverity = 'high' | 'medium' | 'low'
type ActionType = 'Trim to target %' | 'Sell dollar amount' | 'Buy dollar amount' | 'Allocate cash to bucket'

type AIBucket =
  | 'Big tech / hyperscalers'
  | 'AI platforms / AI software'
  | 'Public indirect AI lab exposure'
  | 'Semiconductors'
  | 'GPUs / accelerators'
  | 'CPUs / general compute'
  | 'Memory / storage'
  | 'Semiconductor equipment'
  | 'Fabs / foundries'
  | 'Photonics / optical / interconnect'
  | 'Networking'
  | 'Data centers / colocation'
  | 'Power generation'
  | 'Grid / electrification'
  | 'Cooling / thermal management'
  | 'Materials / specialty chemicals'
  | 'Cybersecurity'
  | 'Enterprise software'
  | 'Robotics / automation'
  | 'Broad passive index exposure'

type BucketExposure = { bucket: AIBucket; weight: number }
type AIExposureClassification = {
  score: number
  buckets: BucketExposure[]
  directness: Directness
  confidence: 'high' | 'medium' | 'low'
  notes: string
  source: ClassificationSource
}
type Holding = {
  id: string
  accountOwner: string
  accountName: string
  accountType: string
  ticker: string
  securityName: string
  shares: number
  price?: number
  marketValue: number
  assetClass: AssetClass
  sector: string
  ai: AIExposureClassification
  costBasis?: number
  unrealizedGainLoss?: number
  notes?: string
}
type PortfolioSnapshot = { id: string; name: string; date: string; source: string; holdings: Holding[] }
type StrategyTemplate = {
  id: string
  name: string
  purpose: string
  targets: Record<AssetClass, number>
  aiRange: [number, number]
  maxSingle: number
  maxTop10: number
  maxBucket: number
  risk: string
  drawdownRange: string
  upside: number
  complexity: number
  pros: string[]
  cons: string[]
  bestFor: string
  warningLabel: string
  aiInternalSplit?: Partial<Record<AIBucket, number>>
}
type RecompCandidate = {
  id: string
  actionType: ActionType
  tickerOrBucket: string
  dollarAmount: number
  estimatedShares?: number
  beforeWeight: number
  afterWeight: number
  reason: string
  riskImpact: string
  alignmentImpact: string
  stressImpact: string
  warning?: string
}
type ManualSandboxAction = RecompCandidate & { note: string }
type StressScenario = { id: string; name: string; shocks: Partial<Record<AssetClass | AIBucket, number>> }
type AppWarning = { id: string; severity: WarningSeverity; title: string; detail: string }
type DecisionLogEntry = {
  id: string
  date: string
  snapshotName: string
  templateName: string
  currentAIExposure: number
  driftScore: number
  actions: RecompCandidate[]
  warnings: AppWarning[]
  notes: string
}
const assetClasses: AssetClass[] = ['Broad US equity', 'AI buildout sleeve', 'International equity', 'Bonds/fixed income', 'Cash', 'Alternatives/other']
const aiBuckets: AIBucket[] = [
  'Big tech / hyperscalers',
  'AI platforms / AI software',
  'Public indirect AI lab exposure',
  'Semiconductors',
  'GPUs / accelerators',
  'CPUs / general compute',
  'Memory / storage',
  'Semiconductor equipment',
  'Fabs / foundries',
  'Photonics / optical / interconnect',
  'Networking',
  'Data centers / colocation',
  'Power generation',
  'Grid / electrification',
  'Cooling / thermal management',
  'Materials / specialty chemicals',
  'Cybersecurity',
  'Enterprise software',
  'Robotics / automation',
  'Broad passive index exposure',
]
const colors = ['#34d399', '#60a5fa', '#a7f3d0', '#fbbf24', '#d1d5db', '#94a3b8', '#f87171', '#22d3ee']
const defaultClassifications: Record<string, Partial<Holding>> = {
  AAPL: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('AI platforms / AI software', 4, 'direct') },
  MSFT: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Big tech / hyperscalers', 5, 'direct', [{ bucket: 'Public indirect AI lab exposure', weight: 25 }]) },
  GOOGL: { assetClass: 'AI buildout sleeve', sector: 'Communication Services', ai: ai('Big tech / hyperscalers', 5, 'direct') },
  AMZN: { assetClass: 'AI buildout sleeve', sector: 'Consumer Discretionary', ai: ai('Big tech / hyperscalers', 4, 'direct') },
  NVDA: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('GPUs / accelerators', 5, 'direct', [{ bucket: 'Semiconductors', weight: 40 }]) },
  AMD: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('GPUs / accelerators', 4, 'direct') },
  AVGO: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Networking', 4, 'direct', [{ bucket: 'Semiconductors', weight: 40 }]) },
  TSM: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Fabs / foundries', 4, 'direct') },
  ASML: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Semiconductor equipment', 5, 'direct') },
  MU: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Memory / storage', 4, 'direct') },
  VRT: { assetClass: 'AI buildout sleeve', sector: 'Industrials', ai: ai('Cooling / thermal management', 4, 'direct') },
  ETN: { assetClass: 'AI buildout sleeve', sector: 'Industrials', ai: ai('Grid / electrification', 3, 'indirect') },
  CEG: { assetClass: 'AI buildout sleeve', sector: 'Utilities', ai: ai('Power generation', 3, 'indirect') },
  EQIX: { assetClass: 'AI buildout sleeve', sector: 'Real Estate', ai: ai('Data centers / colocation', 4, 'direct') },
  PLD: { assetClass: 'Alternatives/other', sector: 'Real Estate', ai: ai('Data centers / colocation', 1, 'indirect') },
  PANW: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Cybersecurity', 3, 'direct') },
  CRWD: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Cybersecurity', 3, 'direct') },
  VTI: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  VOO: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  VXUS: { assetClass: 'International equity', sector: 'International Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  BND: { assetClass: 'Bonds/fixed income', sector: 'Fixed Income', ai: ai('Broad passive index exposure', 0, 'none') },
  SGOV: { assetClass: 'Cash', sector: 'Cash & equivalents', ai: ai('Broad passive index exposure', 0, 'none') },
  SPAXX: { assetClass: 'Cash', sector: 'Cash & equivalents', ai: ai('Broad passive index exposure', 0, 'none') },
}

function ai(bucket: AIBucket, score: number, directness: Directness, extra: BucketExposure[] = []): AIExposureClassification {
  const primaryWeight = Math.max(0, 100 - extra.reduce((sum, item) => sum + item.weight, 0))
  return { score, directness, confidence: 'medium', source: 'default', notes: '', buckets: [{ bucket, weight: primaryWeight }, ...extra] }
}

const strategyTemplates: StrategyTemplate[] = [
  template('balanced-ai-growth', 'Balanced AI Growth', 'Meaningful AI upside while keeping a healthier diversified structure.', [40, 25, 10, 15, 5, 5], [20, 30], 7, 45, 35, 'Moderate growth', '-18% to -30%', 7, 5, ['Strong AI participation', 'Better diversified than a pure tech bet', 'Bonds and cash ballast'], ['Can still draw down in a tech bear market', 'May trail a pure AI rally'], 'A growth-oriented family portfolio that wants AI buildout exposure without going fully aggressive.', 'Watch single-name concentration'),
  template('aggressive-ai-infrastructure', 'Aggressive AI Infrastructure', 'Maximize participation in the AI infrastructure buildout while keeping basic guardrails.', [30, 40, 5, 10, 5, 10], [35, 45], 10, 60, 40, 'High growth', '-25% to -42%', 9, 7, ['Highest AI upside participation', 'Strong exposure across compute, chips, power, data centers, and infrastructure'], ['Higher volatility', 'More vulnerable to AI capex disappointment'], 'An intentionally aggressive AI infrastructure bet.', 'Higher volatility'),
  template('ai-barbell', 'AI Barbell', 'Pair aggressive AI upside with defensive ballast.', [25, 35, 5, 20, 10, 5], [30, 40], 8, 50, 35, 'Barbell growth', '-20% to -36%', 8, 6, ['Meaningful AI upside', 'More cash and bond protection', 'Dry powder after drawdowns'], ['Cash and bonds may drag in bull markets', 'Still exposed to AI theme volatility'], 'Someone bullish on AI but uncomfortable with full high-beta exposure.', 'Barbell volatility'),
  {
    ...template('diversified-ai-supply-chain', 'Diversified AI Supply Chain', 'Spread AI exposure across the full buildout chain, not only mega-cap tech and chips.', [35, 30, 10, 15, 5, 5], [25, 35], 7, 50, 30, 'Diversified growth', '-20% to -34%', 8, 7, ['Best fit for the broad AI infrastructure thesis', 'Captures second-order beneficiaries', 'Avoids relying only on mega-cap tech or chips'], ['More complex', 'Some infrastructure names may not move like classic AI stocks'], 'A portfolio built around the whole AI supply chain.', 'Balance the AI sleeve'),
    aiInternalSplit: {
      'Big tech / hyperscalers': 20,
      'Semiconductors': 25,
      'Power generation': 10,
      'Grid / electrification': 10,
      'Data centers / colocation': 10,
      'Cooling / thermal management': 5,
      'Networking': 5,
      'Photonics / optical / interconnect': 5,
      'AI platforms / AI software': 5,
      Cybersecurity: 5,
    },
  },
  template('conservative-ai-participation', 'Conservative AI Participation', 'Keep AI exposure meaningful while prioritizing drawdown control.', [40, 15, 10, 25, 7, 3], [10, 20], 5, 35, 30, 'Capital preservation', '-12% to -24%', 5, 4, ['Lower expected volatility', 'Still participates in AI buildout', 'Better fit for capital preservation'], ['Less upside if AI infrastructure stocks rally hard', 'May feel too conservative for high conviction'], 'A portfolio where preserving capital matters more than maximizing AI upside.', 'Lower AI sleeve'),
]

function template(id: string, name: string, purpose: string, targetValues: number[], aiRange: [number, number], maxSingle: number, maxTop10: number, maxBucket: number, risk: string, drawdownRange: string, upside: number, complexity: number, pros: string[], cons: string[], bestFor: string, warningLabel: string): StrategyTemplate {
  return { id, name, purpose, targets: Object.fromEntries(assetClasses.map((asset, index) => [asset, targetValues[index]])) as Record<AssetClass, number>, aiRange, maxSingle, maxTop10, maxBucket, risk, drawdownRange, upside, complexity, pros, cons, bestFor, warningLabel }
}

const stressScenarios: StressScenario[] = [
  { id: 'market-correction', name: 'Market correction', shocks: { 'Broad US equity': -10, 'AI buildout sleeve': -15, 'Bonds/fixed income': 0, Cash: 0, 'Alternatives/other': -5 } },
  { id: 'bear-market', name: 'Bear market', shocks: { 'Broad US equity': -20, 'AI buildout sleeve': -35, 'Bonds/fixed income': 2, Cash: 0, 'Alternatives/other': -10 } },
  { id: 'ai-disappointment', name: 'AI disappointment', shocks: { 'Broad US equity': -10, 'AI buildout sleeve': -45, 'Bonds/fixed income': 2, Cash: 0, 'Alternatives/other': -8 } },
  { id: 'ai-mania', name: 'AI mania continuation', shocks: { 'Broad US equity': 12, 'AI buildout sleeve': 35, 'Bonds/fixed income': 0, Cash: 0, 'Alternatives/other': 8 } },
  { id: 'rates-power', name: 'Rates up / power bottleneck', shocks: { 'Broad US equity': -8, 'AI buildout sleeve': -15, 'Bonds/fixed income': -8, Cash: 0, 'Power generation': 8, 'Grid / electrification': 8 } },
  { id: 'big-tech-compression', name: 'Big tech multiple compression', shocks: { 'Broad US equity': -12, 'Big tech / hyperscalers': -30, Semiconductors: -25, 'Bonds/fixed income': 1, Cash: 0 } },
  { id: 'semi-shock', name: 'Semiconductor shock', shocks: { 'Broad US equity': -8, Semiconductors: -35, 'GPUs / accelerators': -35, 'Semiconductor equipment': -35, 'Big tech / hyperscalers': -15, 'Power generation': -5, 'Grid / electrification': -5, 'Data centers / colocation': -5, 'Cooling / thermal management': -5, 'Bonds/fixed income': 1, Cash: 0 } },
]

function sampleSnapshots(): PortfolioSnapshot[] {
  const holdings: Holding[] = [
    holding('Family', 'Joint taxable', 'Taxable', 'VTI', 'Vanguard Total Stock Market ETF', 540, 247.8, 'Broad US equity', 'Broad Market'),
    holding('Family', 'Joint taxable', 'Taxable', 'MSFT', 'Microsoft Corp', 220, 506.4, 'AI buildout sleeve', 'Information Technology'),
    holding('Family', 'Joint taxable', 'Taxable', 'NVDA', 'NVIDIA Corp', 390, 187.6, 'AI buildout sleeve', 'Information Technology'),
    holding('Family', 'Joint taxable', 'Taxable', 'ETN', 'Eaton Corp', 105, 421.3, 'AI buildout sleeve', 'Industrials'),
    holding('Family', 'Joint taxable', 'Taxable', 'EQIX', 'Equinix Inc', 38, 812.4, 'AI buildout sleeve', 'Real Estate'),
    holding('Family', 'Roth IRA', 'Roth IRA', 'AAPL', 'Apple Inc', 160, 293.05, 'AI buildout sleeve', 'Information Technology'),
    holding('Family', 'Roth IRA', 'Roth IRA', 'AAPL', 'Apple Inc', 36, 293.05, 'AI buildout sleeve', 'Information Technology'),
    holding('Family', 'Roth IRA', 'Roth IRA', 'ASML', 'ASML Holding NV', 22, 725.9, 'AI buildout sleeve', 'Information Technology'),
    holding('Family', '401k', 'Retirement', 'BND', 'Vanguard Total Bond Market ETF', 710, 72.1, 'Bonds/fixed income', 'Fixed Income'),
    holding('Family', '401k', 'Retirement', 'VXUS', 'Vanguard Total International Stock ETF', 620, 65.4, 'International equity', 'International Broad Market'),
    holding('Family', '401k', 'Retirement', 'SGOV', 'iShares 0-3 Month Treasury Bond ETF', 260, 100.6, 'Cash', 'Cash & equivalents'),
    holding('Family', 'DB Plan', 'Tax-advantaged', 'VRT', 'Vertiv Holdings', 135, 176.7, 'AI buildout sleeve', 'Industrials'),
    holding('Family', 'DB Plan', 'Tax-advantaged', 'PANW', 'Palo Alto Networks', 55, 214.2, 'AI buildout sleeve', 'Information Technology'),
    holding('Family', 'DB Plan', 'Tax-advantaged', 'CEG', 'Constellation Energy', 80, 351.8, 'AI buildout sleeve', 'Utilities'),
    holding('Family', 'DB Plan', 'Tax-advantaged', 'AGNC', 'AGNC Investment Corp REIT', 300, 10.86, 'Alternatives/other', 'Real Estate'),
    holding('Family', 'DB Plan', 'Tax-advantaged', 'PSHZF', 'Pershing Square Holdings Rights', 157, 0.35, 'Alternatives/other', 'Financials'),
  ]
  const prior = holdings.map((item, index) => ({ ...item, id: `${item.id}-prior`, shares: index % 3 === 0 ? item.shares * 0.92 : item.shares, marketValue: index % 4 === 0 ? item.marketValue * 0.88 : item.marketValue * 0.96 }))
  return [
    { id: 'sample-prior', name: 'Sample prior snapshot', date: new Date(Date.now() - 1000 * 60 * 60 * 24 * 35).toISOString(), source: 'Synthetic sample', holdings: prior },
    { id: 'sample-current', name: 'Sample current snapshot', date: new Date().toISOString(), source: 'Synthetic sample', holdings },
  ]
}

function holding(accountOwner: string, accountName: string, accountType: string, ticker: string, securityName: string, shares: number, price: number, assetClass: AssetClass, sector: string): Holding {
  const defaults = defaultClassifications[ticker] ?? {}
  return {
    id: `${accountName}-${ticker}-${Math.random().toString(36).slice(2)}`,
    accountOwner,
    accountName,
    accountType,
    ticker,
    securityName,
    shares,
    price,
    marketValue: shares * price,
    assetClass: (defaults.assetClass as AssetClass) ?? assetClass,
    sector: defaults.sector ?? sector,
    ai: (defaults.ai as AIExposureClassification) ?? { score: 0, buckets: [], directness: 'none', confidence: 'low', source: 'unknown', notes: '', },
    costBasis: shares * price * 0.78,
    unrealizedGainLoss: shares * price * 0.22,
  }
}

const dollarFmt = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 })
const percentFmt = new Intl.NumberFormat('en-US', { maximumFractionDigits: 1 })
const storageKey = 'ai-buildout-portfolio-recomp-state-v1'

function App() {
  const persisted = loadState()
  const [snapshots, setSnapshots] = useState<PortfolioSnapshot[]>(persisted.snapshots)
  const [selectedSnapshotId, setSelectedSnapshotId] = useState(persisted.selectedSnapshotId)
  const [activeTab, setActiveTab] = useState('Overview')
  const [selectedTemplateId, setSelectedTemplateId] = useState(persisted.selectedTemplateId)
  const [selectedStressId, setSelectedStressId] = useState('ai-disappointment')
  const [recompCandidates, setRecompCandidates] = useState<RecompCandidate[]>(persisted.recompCandidates)
  const [manualActions, setManualActions] = useState<ManualSandboxAction[]>(persisted.manualActions)
  const [decisionLog, setDecisionLog] = useState<DecisionLogEntry[]>(persisted.decisionLog)
  const [search, setSearch] = useState('')
  const [grouping, setGrouping] = useState<GroupingState>([])
  const [sorting, setSorting] = useState<SortingState>([])
  const [minDollar, setMinDollar] = useState(1500)
  const [minWeight, setMinWeight] = useState(0.25)
  const [importRows, setImportRows] = useState<Record<string, unknown>[]>([])
  const [importName, setImportName] = useState('')
  const [importAccountOwner, setImportAccountOwner] = useState('Family')
  const [importAccountType, setImportAccountType] = useState('Tax-advantaged')
  const [importMessage, setImportMessage] = useState('')
  const [selectedHoldingId, setSelectedHoldingId] = useState<string | null>(null)

  const currentSnapshot = snapshots.find((snapshot) => snapshot.id === selectedSnapshotId) ?? snapshots.at(-1)!
  const template = strategyTemplates.find((item) => item.id === selectedTemplateId) ?? strategyTemplates[0]
  const stress = stressScenarios.find((item) => item.id === selectedStressId) ?? stressScenarios[2]
  const analytics = useMemo(() => analyze(currentSnapshot.holdings, template, stress, minDollar, minWeight), [currentSnapshot, template, stress, minDollar, minWeight])
  const simulated = useMemo(() => simulateCandidates(currentSnapshot.holdings, [...recompCandidates, ...manualActions]), [currentSnapshot, recompCandidates, manualActions])
  const simulatedAnalytics = useMemo(() => analyze(simulated, template, stress, minDollar, minWeight), [simulated, template, stress, minDollar, minWeight])
  const selectedHolding = currentSnapshot.holdings.find((item) => item.id === selectedHoldingId)

  persist({ snapshots, selectedSnapshotId: currentSnapshot.id, selectedTemplateId: template.id, recompCandidates, manualActions, decisionLog })

  const columns = useMemo<ColumnDef<Holding>[]>(() => [
    { accessorKey: 'ticker', header: 'Ticker' },
    { accessorKey: 'securityName', header: 'Security' },
    { accessorKey: 'accountName', header: 'Account' },
    { accessorKey: 'marketValue', header: 'Market value', cell: ({ row }) => dollarFmt.format(row.original.marketValue) },
    { id: 'portfolioWeight', header: 'Portfolio %', cell: ({ row }) => `${percentFmt.format(weight(row.original.marketValue, analytics.total))}%` },
    { accessorKey: 'assetClass', header: 'Asset class' },
    { accessorKey: 'sector', header: 'Sector' },
    { id: 'aiBucket', header: 'AI bucket', cell: ({ row }) => row.original.ai.buckets[0]?.bucket ?? 'Missing' },
    { id: 'aiScore', header: 'AI score', cell: ({ row }) => row.original.ai.score.toFixed(1) },
    { id: 'directness', header: 'Directness', cell: ({ row }) => row.original.ai.directness },
    { id: 'warning', header: 'Warning', cell: ({ row }) => row.original.marketValue < minDollar ? 'Nuisance' : weight(row.original.marketValue, analytics.total) > template.maxSingle ? 'High concentration' : row.original.ai.source === 'unknown' ? 'Missing AI data' : 'Clear' },
  ], [analytics.total, minDollar, template.maxSingle])
  // eslint-disable-next-line react-hooks/incompatible-library
  const table = useReactTable({ data: currentSnapshot.holdings, columns, state: { globalFilter: search, grouping, sorting }, onGlobalFilterChange: setSearch, onGroupingChange: setGrouping, onSortingChange: setSorting, getCoreRowModel: getCoreRowModel(), getFilteredRowModel: getFilteredRowModel(), getGroupedRowModel: getGroupedRowModel(), getSortedRowModel: getSortedRowModel() })

  function generateCandidates() {
    const candidates = buildRecompCandidates(analytics, template)
    setRecompCandidates(candidates)
    setDecisionLog((logs) => [{ id: crypto.randomUUID(), date: new Date().toISOString(), snapshotName: currentSnapshot.name, templateName: template.name, currentAIExposure: analytics.aiExposure, driftScore: analytics.driftScore, actions: candidates, warnings: analytics.warnings, notes: 'Auto-generated simulated recomp candidate run.' }, ...logs])
    setActiveTab('Recomp Sandbox')
  }

  function updateHolding(updated: Holding) {
    setSnapshots((items) => items.map((snapshot) => snapshot.id === currentSnapshot.id ? { ...snapshot, holdings: snapshot.holdings.map((item) => item.id === updated.id ? updated : item) } : snapshot))
  }

  async function parseUpload(file: File) {
    const buffer = await file.arrayBuffer()
    const text = new TextDecoder().decode(buffer)
    let rows: Record<string, unknown>[]
    let sourceName = file.name
    if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
      const workbook = XLSX.read(buffer)
      const sheet = workbook.Sheets[workbook.SheetNames[0]]
      rows = XLSX.utils.sheet_to_json(sheet, { defval: '' })
    } else {
      const lines = text.split(/\r?\n/)
      const accountLine = lines.find((line) => line.includes('Positions for account'))
      sourceName = accountLine?.replaceAll('"', '').trim() || file.name
      const headerIndex = lines.findIndex((line) => line.includes('Symbol') && line.includes('Description'))
      const parseText = headerIndex >= 0 ? lines.slice(headerIndex).join('\n') : text
      const parsed = Papa.parse<Record<string, unknown>>(parseText, { header: true, skipEmptyLines: true })
      rows = parsed.data
    }
    setImportRows(rows.filter((row) => row.Symbol || row.ticker || row.Ticker))
    setImportName(sourceName)
    setImportMessage(`${rows.length} rows detected. Review account fields, then save as a snapshot.`)
  }

  function saveImportSnapshot() {
    const holdings = importRows.map((row, index) => normalizeImportedRow(row, importAccountOwner, importName || 'Uploaded account', importAccountType, index)).filter((row): row is Holding => Boolean(row))
    if (!holdings.length) {
      setImportMessage('No usable holdings found. Confirm the file includes ticker and market value columns.')
      return
    }
    const snapshot: PortfolioSnapshot = { id: crypto.randomUUID(), name: `${importName || 'Uploaded portfolio'} snapshot`, date: new Date().toISOString(), source: importName || 'Uploaded file', holdings }
    setSnapshots((items) => [...items, snapshot])
    setSelectedSnapshotId(snapshot.id)
    setImportRows([])
    setImportMessage(`Saved ${holdings.length} holdings as a dated snapshot.`)
  }

  return (
    <main className="min-h-[100dvh] bg-zinc-950 text-zinc-100">
      <div className="mx-auto max-w-[1400px] px-4 py-5 md:px-6">
        <header className="mb-6 grid gap-5 border-b border-white/10 pb-5 lg:grid-cols-[1.5fr_1fr]">
          <div>
            <div className="mb-3 flex flex-wrap items-center gap-2 text-xs uppercase tracking-[0.18em] text-emerald-300/80">
              <span className="rounded-full border border-emerald-300/20 bg-emerald-300/10 px-3 py-1">Local-first</span>
              <span>Offline portfolio planning</span>
            </div>
            <h1 className="max-w-4xl text-3xl font-semibold tracking-tight text-zinc-50 md:text-5xl">AI Buildout Portfolio Recomp Cockpit</h1>
            <p className="mt-3 max-w-3xl text-sm leading-6 text-zinc-400 md:text-base">
              A private browser workspace for x-raying family holdings, measuring AI infrastructure exposure, stress testing simple scenarios, and drafting simulated recomp candidates.
            </p>
          </div>
          <div className="rounded-2xl border border-white/10 bg-zinc-900/70 p-4 shadow-[inset_0_1px_0_rgba(255,255,255,0.06)]">
            <div className="flex items-center justify-between gap-4">
              <span className="text-sm text-zinc-400">Current snapshot</span>
              <select className="control" value={currentSnapshot.id} onChange={(event) => setSelectedSnapshotId(event.target.value)}>
                {snapshots.map((snapshot) => <option key={snapshot.id} value={snapshot.id}>{snapshot.name}</option>)}
              </select>
            </div>
            <div className="mt-4 grid grid-cols-2 gap-3 text-sm">
              <Metric label="Value" value={dollarFmt.format(analytics.total)} />
              <Metric label="AI exposure" value={`${percentFmt.format(analytics.aiExposure)}%`} />
              <Metric label="Top 10" value={`${percentFmt.format(analytics.top10Weight)}%`} />
              <Metric label="Stress" value={`${percentFmt.format(analytics.stressImpact)}%`} />
            </div>
          </div>
        </header>

        <nav className="mb-6 flex gap-2 overflow-x-auto rounded-2xl border border-white/10 bg-zinc-900/80 p-2">
          {['Overview', 'Import & Snapshots', 'X-Ray & Concentration', 'AI Buildout', 'Recomp Sandbox'].map((tab) => (
            <button key={tab} className={`tab ${activeTab === tab ? 'tab-active' : ''}`} onClick={() => setActiveTab(tab)}>{tab}</button>
          ))}
        </nav>

        {activeTab === 'Overview' && <Overview analytics={analytics} template={template} stress={stress} setTemplate={setSelectedTemplateId} setStress={setSelectedStressId} generateCandidates={generateCandidates} />}
        {activeTab === 'Import & Snapshots' && <ImportSnapshots snapshots={snapshots} current={currentSnapshot} rows={importRows} message={importMessage} accountOwner={importAccountOwner} accountType={importAccountType} setAccountOwner={setImportAccountOwner} setAccountType={setImportAccountType} parseUpload={parseUpload} saveImportSnapshot={saveImportSnapshot} setSnapshots={setSnapshots} setSelectedSnapshotId={setSelectedSnapshotId} generateCandidates={generateCandidates} />}
        {activeTab === 'X-Ray & Concentration' && <Xray analytics={analytics} table={table} search={search} setSearch={setSearch} grouping={grouping} setGrouping={setGrouping} minDollar={minDollar} minWeight={minWeight} setMinDollar={setMinDollar} setMinWeight={setMinWeight} setSelectedHoldingId={setSelectedHoldingId} />}
        {activeTab === 'AI Buildout' && <AIBuildout analytics={analytics} current={currentSnapshot} selectedHolding={selectedHolding} setSelectedHoldingId={setSelectedHoldingId} updateHolding={updateHolding} />}
        {activeTab === 'Recomp Sandbox' && <Sandbox analytics={analytics} simulatedAnalytics={simulatedAnalytics} holdings={currentSnapshot.holdings} candidates={recompCandidates} manualActions={manualActions} setManualActions={setManualActions} clearCandidates={() => setRecompCandidates([])} generateCandidates={generateCandidates} decisionLog={decisionLog} />}

        <footer className="mt-8 border-t border-white/10 pt-4 text-xs leading-5 text-zinc-500">
          This app is for portfolio analysis and planning only. It does not provide financial advice, tax advice, or execute trades. All outputs are simulations based on uploaded data and simplified assumptions.
        </footer>
      </div>
    </main>
  )
}

function Overview({ analytics, template, stress, setTemplate, setStress, generateCandidates }: { analytics: ReturnType<typeof analyze>; template: StrategyTemplate; stress: StressScenario; setTemplate: (id: string) => void; setStress: (id: string) => void; generateCandidates: () => void }) {
  return <section className="grid gap-5">
    <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-5">
      <MetricCard icon={<Database size={20} />} label="Total portfolio value" value={dollarFmt.format(analytics.total)} />
      <MetricCard icon={<StackSimple size={20} />} label="Holdings" value={String(analytics.holdings.length)} />
      <MetricCard icon={<Sparkle size={20} />} label="AI buildout exposure" value={`${percentFmt.format(analytics.aiExposure)}%`} />
      <MetricCard icon={<Funnel size={20} />} label="Largest holding" value={`${percentFmt.format(analytics.largestWeight)}%`} />
      <MetricCard icon={<Pulse size={20} />} label="Risk-budget score" value={analytics.riskScore.toFixed(2)} />
    </div>
    <div className="grid gap-5 xl:grid-cols-[1.2fr_0.8fr]">
      <Panel title="30-second read" action={<button className="primary" onClick={generateCandidates}><Sparkle size={16} /> Generate Auto-Recomp Candidates</button>}>
        <p className="text-balance text-lg leading-8 text-zinc-200">{executiveSummary(analytics, template)}</p>
        <div className="mt-5 grid gap-3 md:grid-cols-2">
          <label className="field"><span>Selected strategy template</span><select className="control" value={template.id} onChange={(event) => setTemplate(event.target.value)}>{strategyTemplates.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select></label>
          <label className="field"><span>Stress scenario</span><select className="control" value={stress.id} onChange={(event) => setStress(event.target.value)}>{stressScenarios.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select></label>
        </div>
      </Panel>
      <WarningPanel warnings={analytics.warnings} />
    </div>
    <div className="grid gap-5 xl:grid-cols-2">
      <ChartPanel title="Asset class"><Donut data={analytics.assetClassData} /></ChartPanel>
      <ChartPanel title="Sector"><Donut data={analytics.sectorData.slice(0, 8)} /></ChartPanel>
      <ChartPanel title="Current vs target allocation"><TargetBars analytics={analytics} template={template} /></ChartPanel>
      <ChartPanel title="AI buildout exposure by bucket"><BarList data={analytics.aiBucketData.slice(0, 10)} /></ChartPanel>
      <ChartPanel title="Top 10 holdings"><BarList data={analytics.topHoldings.slice(0, 10).map((h) => ({ name: h.ticker, value: weight(h.marketValue, analytics.total) }))} /></ChartPanel>
    </div>
  </section>
}

function ImportSnapshots(props: { snapshots: PortfolioSnapshot[]; current: PortfolioSnapshot; rows: Record<string, unknown>[]; message: string; accountOwner: string; accountType: string; setAccountOwner: (v: string) => void; setAccountType: (v: string) => void; parseUpload: (file: File) => void; saveImportSnapshot: () => void; setSnapshots: React.Dispatch<React.SetStateAction<PortfolioSnapshot[]>>; setSelectedSnapshotId: (id: string) => void; generateCandidates: () => void }) {
  const latest = props.snapshots.at(-1)
  const previous = props.snapshots.at(-2)
  const comparison = latest && previous ? compareSnapshots(previous, latest) : []
  return <section className="grid gap-5 xl:grid-cols-[0.9fr_1.1fr]">
    <Panel title="Upload holdings" action={<button className="primary" onClick={props.generateCandidates}><Sparkle size={16} /> Generate Auto-Recomp Candidates</button>}>
      <label className="upload">
        <FileArrowUp size={28} />
        <span>Drop in a CSV or XLSX brokerage positions file</span>
        <input type="file" accept=".csv,.xlsx,.xls" onChange={(event) => event.target.files?.[0] && props.parseUpload(event.target.files[0])} />
      </label>
      {props.message && <p className="mt-3 rounded-xl border border-emerald-300/20 bg-emerald-300/10 p-3 text-sm text-emerald-100">{props.message}</p>}
      {props.rows.length > 0 && <div className="mt-4 grid gap-3 md:grid-cols-2">
        <label className="field"><span>Account owner</span><input className="control" value={props.accountOwner} onChange={(event) => props.setAccountOwner(event.target.value)} /></label>
        <label className="field"><span>Account type</span><input className="control" value={props.accountType} onChange={(event) => props.setAccountType(event.target.value)} /></label>
        <button className="primary md:col-span-2" onClick={props.saveImportSnapshot}><UploadSimple size={16} /> Save as dated snapshot</button>
        <PreviewRows rows={props.rows} />
      </div>}
    </Panel>
    <Panel title="Snapshots and changes">
      <div className="space-y-3">
        {props.snapshots.map((snapshot) => <div key={snapshot.id} className="row">
          <div><p className="font-medium text-zinc-100">{snapshot.name}</p><p className="text-xs text-zinc-500">{new Date(snapshot.date).toLocaleString()} · {snapshot.holdings.length} holdings · {snapshot.source}</p></div>
          <div className="flex gap-2">
            <button className="ghost" onClick={() => props.setSelectedSnapshotId(snapshot.id)}>View</button>
            {props.snapshots.length > 1 && <button className="icon-button" onClick={() => props.setSnapshots((items) => items.filter((item) => item.id !== snapshot.id))}><Trash size={16} /></button>}
          </div>
        </div>)}
      </div>
      <div className="mt-5">
        <h3 className="section-label">Latest vs previous</h3>
        <div className="max-h-80 overflow-auto rounded-xl border border-white/10">
          {comparison.map((item) => <div key={item.name} className="grid grid-cols-[1fr_auto_auto] gap-3 border-b border-white/5 p-3 text-sm last:border-0"><span>{item.name}</span><span className="font-mono text-zinc-300">{dollarFmt.format(item.change)}</span><span className={item.change >= 0 ? 'text-emerald-300' : 'text-red-300'}>{item.status}</span></div>)}
        </div>
      </div>
    </Panel>
  </section>
}

function Xray({ analytics, table, search, setSearch, grouping, setGrouping, minDollar, minWeight, setMinDollar, setMinWeight, setSelectedHoldingId }: { analytics: ReturnType<typeof analyze>; table: ReturnType<typeof useReactTable<Holding>>; search: string; setSearch: (v: string) => void; grouping: GroupingState; setGrouping: (v: GroupingState) => void; minDollar: number; minWeight: number; setMinDollar: (v: number) => void; setMinWeight: (v: number) => void; setSelectedHoldingId: (id: string) => void }) {
  return <section className="grid gap-5">
    <div className="grid gap-5 xl:grid-cols-2">
      <ChartPanel title="Top holdings ranked"><BarList data={analytics.topHoldings.slice(0, 15).map((h) => ({ name: h.ticker, value: weight(h.marketValue, analytics.total) }))} /></ChartPanel>
      <ChartPanel title="Holdings treemap"><Treemap width={500} height={300} data={analytics.topHoldings.map((h) => ({ name: h.ticker, size: h.marketValue }))} dataKey="size" aspectRatio={4 / 3} stroke="#18181b" fill="#34d399" /></ChartPanel>
    </div>
    <Panel title="Holdings table" action={<button className="ghost" onClick={() => exportCsv('current-holdings.csv', analytics.holdings)}><DownloadSimple size={16} /> Export</button>}>
      <div className="mb-4 grid gap-3 lg:grid-cols-[1fr_auto_auto_auto]">
        <label className="search"><MagnifyingGlass size={16} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search holdings" /></label>
        <select className="control" value={grouping[0] ?? ''} onChange={(event) => setGrouping(event.target.value ? [event.target.value] : [])}><option value="">No grouping</option><option value="accountName">Group account</option><option value="assetClass">Group asset class</option><option value="sector">Group sector</option></select>
        <label className="compact-field">Min $<input type="number" className="control" value={minDollar} onChange={(event) => setMinDollar(Number(event.target.value))} /></label>
        <label className="compact-field">Min %<input type="number" className="control" value={minWeight} onChange={(event) => setMinWeight(Number(event.target.value))} /></label>
      </div>
      <div className="overflow-auto rounded-xl border border-white/10">
        <table className="data-table">
          <thead>{table.getHeaderGroups().map((group) => <tr key={group.id}>{group.headers.map((header) => <th key={header.id} onClick={header.column.getToggleSortingHandler()}>{flexRender(header.column.columnDef.header, header.getContext())}</th>)}</tr>)}</thead>
          <tbody>{table.getRowModel().rows.map((row) => <tr key={row.id} onClick={() => setSelectedHoldingId(row.original.id)}>{row.getVisibleCells().map((cell) => <td key={cell.id}>{flexRender(cell.column.columnDef.cell, cell.getContext())}</td>)}</tr>)}</tbody>
        </table>
      </div>
    </Panel>
  </section>
}

function AIBuildout({ analytics, current, selectedHolding, setSelectedHoldingId, updateHolding }: { analytics: ReturnType<typeof analyze>; current: PortfolioSnapshot; selectedHolding?: Holding; setSelectedHoldingId: (id: string | null) => void; updateHolding: (holding: Holding) => void }) {
  return <section className="grid gap-5 xl:grid-cols-[1fr_380px]">
    <div className="grid gap-5">
      <ChartPanel title="AI exposure by bucket"><BarList data={analytics.aiBucketData} /></ChartPanel>
      <ChartPanel title="Direct vs indirect AI exposure"><Donut data={analytics.directnessData} /></ChartPanel>
      <Panel title="Top AI exposure contributors">
        <div className="space-y-2">
          {analytics.holdings.filter((h) => h.ai.score > 0).sort((a, b) => b.marketValue * b.ai.score - a.marketValue * a.ai.score).slice(0, 12).map((holding) => <button key={holding.id} className="row w-full text-left" onClick={() => setSelectedHoldingId(holding.id)}><span><strong>{holding.ticker}</strong> <span className="text-zinc-500">{holding.securityName}</span></span><span className="font-mono text-emerald-300">{holding.ai.score.toFixed(1)}</span></button>)}
        </div>
      </Panel>
    </div>
    <Panel title="Editable classification" action={selectedHolding && <button className="ghost" onClick={() => setSelectedHoldingId(null)}>Close</button>}>
      {!selectedHolding ? <div className="empty"><PencilSimple size={28} /><p>Select a holding from the contributor list or x-ray table to edit its AI classification.</p></div> : <ClassificationEditor holding={selectedHolding} current={current} updateHolding={updateHolding} />}
    </Panel>
  </section>
}

function Sandbox({ analytics, simulatedAnalytics, holdings, candidates, manualActions, setManualActions, clearCandidates, generateCandidates, decisionLog }: { analytics: ReturnType<typeof analyze>; simulatedAnalytics: ReturnType<typeof analyze>; holdings: Holding[]; candidates: RecompCandidate[]; manualActions: ManualSandboxAction[]; setManualActions: React.Dispatch<React.SetStateAction<ManualSandboxAction[]>>; clearCandidates: () => void; generateCandidates: () => void; decisionLog: DecisionLogEntry[] }) {
  const [ticker, setTicker] = useState(holdings[0]?.ticker ?? '')
  const [amount, setAmount] = useState(5000)
  const allActions = [...candidates, ...manualActions]
  function addManual(actionType: ActionType) {
    const holding = holdings.find((item) => item.ticker === ticker)
    setManualActions((items) => [...items, { id: crypto.randomUUID(), actionType, tickerOrBucket: ticker, dollarAmount: amount, estimatedShares: holding?.price ? amount / holding.price : undefined, beforeWeight: holding ? weight(holding.marketValue, analytics.total) : 0, afterWeight: holding ? weight(holding.marketValue + (actionType.includes('Buy') ? amount : -amount), analytics.total) : 0, reason: 'Manual sandbox action.', riskImpact: 'User-defined impact.', alignmentImpact: 'Included in simulated before/after view.', stressImpact: 'Recomputed in simulated portfolio.', note: '', }])
  }
  return <section className="grid gap-5">
    <div className="grid gap-5 xl:grid-cols-[0.8fr_1.2fr]">
      <Panel title="Manual sandbox">
        <div className="grid gap-3 md:grid-cols-3">
          <label className="field"><span>Ticker or bucket</span><input className="control" value={ticker} onChange={(event) => setTicker(event.target.value.toUpperCase())} /></label>
          <label className="field"><span>Dollar amount</span><input className="control" type="number" value={amount} onChange={(event) => setAmount(Number(event.target.value))} /></label>
          <div className="flex items-end gap-2"><button className="ghost" onClick={() => addManual('Buy dollar amount')}>Buy</button><button className="ghost" onClick={() => addManual('Sell dollar amount')}>Sell</button></div>
        </div>
      </Panel>
      <Panel title="Auto-Recomp" action={<div className="flex gap-2"><button className="primary" onClick={generateCandidates}><ArrowClockwise size={16} /> Generate</button><button className="ghost" onClick={clearCandidates}>Clear</button></div>}>
        <p className="text-sm leading-6 text-zinc-400">Auto-Recomp uses cash first, trims oversized concentration, reduces overexposed sleeves, and creates bucket-level allocate-here candidates when it should not name a new security.</p>
      </Panel>
    </div>
    <div className="grid gap-5 xl:grid-cols-3">
      <MetricCard icon={<Scales size={20} />} label="Current stress estimate" value={`${percentFmt.format(analytics.stressImpact)}%`} />
      <MetricCard icon={<Scales size={20} />} label="Simulated stress estimate" value={`${percentFmt.format(simulatedAnalytics.stressImpact)}%`} />
      <MetricCard icon={<ChartBar size={20} />} label="Cash remaining" value={dollarFmt.format(simulatedAnalytics.assetTotals.Cash ?? 0)} />
    </div>
    <Panel title="Simulated trade candidate table" action={<button className="ghost" onClick={() => exportCsv('simulated-recomp-candidates.csv', allActions)}><DownloadSimple size={16} /> Export CSV</button>}>
      <ActionTable actions={allActions} />
    </Panel>
    <div className="grid gap-5 xl:grid-cols-2">
      <ChartPanel title="Current vs simulated allocation"><BeforeAfter before={analytics.assetClassData} after={simulatedAnalytics.assetClassData} /></ChartPanel>
      <ChartPanel title="Current vs simulated AI exposure"><BeforeAfter before={analytics.aiBucketData.slice(0, 8)} after={simulatedAnalytics.aiBucketData.slice(0, 8)} /></ChartPanel>
    </div>
    <Panel title="Decision log" action={<button className="ghost" onClick={() => exportJson('decision-log.json', decisionLog)}><DownloadSimple size={16} /> Export JSON</button>}>
      <div className="space-y-2">{decisionLog.map((log) => <div key={log.id} className="row"><span>{new Date(log.date).toLocaleString()} · {log.templateName}</span><span className="font-mono text-zinc-400">{log.actions.length} actions</span></div>)}</div>
    </Panel>
  </section>
}

function Metric({ label, value }: { label: string; value: string }) {
  return <div><p className="text-xs uppercase tracking-[0.16em] text-zinc-500">{label}</p><p className="mt-1 font-mono text-lg text-zinc-100">{value}</p></div>
}
function MetricCard({ icon, label, value }: { icon: React.ReactNode; label: string; value: string }) {
  return <div className="metric-card"><div className="text-emerald-300">{icon}</div><div><p>{label}</p><strong>{value}</strong></div></div>
}
function Panel({ title, action, children }: { title: string; action?: React.ReactNode; children: React.ReactNode }) {
  return <section className="panel"><div className="panel-header"><h2>{title}</h2>{action}</div>{children}</section>
}
function ChartPanel({ title, children }: { title: string; children: React.ReactNode }) {
  return <Panel title={title}><div className="h-80">{children}</div></Panel>
}
function WarningPanel({ warnings }: { warnings: AppWarning[] }) {
  return <Panel title="Warnings"><div className="space-y-3">{warnings.map((warning) => <div key={warning.id} className={`warning warning-${warning.severity}`}><ShieldWarning size={18} /><div><p>{warning.title}</p><span>{warning.detail}</span></div></div>)}</div></Panel>
}
function Donut({ data }: { data: { name: string; value: number }[] }) {
  return <ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={data} dataKey="value" nameKey="name" innerRadius={58} outerRadius={96} paddingAngle={2}>{data.map((_, i) => <Cell key={i} fill={colors[i % colors.length]} />)}</Pie><Tooltip formatter={(v) => `${percentFmt.format(Number(v))}%`} contentStyle={tooltipStyle} /><Legend /></PieChart></ResponsiveContainer>
}
function BarList({ data }: { data: { name: string; value: number }[] }) {
  return <ResponsiveContainer width="100%" height="100%"><BarChart data={data} layout="vertical" margin={{ left: 16, right: 20 }}><CartesianGrid strokeDasharray="3 3" stroke="#27272a" /><XAxis type="number" stroke="#71717a" /><YAxis type="category" dataKey="name" width={150} stroke="#a1a1aa" tick={{ fontSize: 12 }} /><Tooltip formatter={(v) => `${percentFmt.format(Number(v))}%`} contentStyle={tooltipStyle} /><Bar dataKey="value" fill="#34d399" radius={[0, 6, 6, 0]} /></BarChart></ResponsiveContainer>
}
function TargetBars({ analytics, template }: { analytics: ReturnType<typeof analyze>; template: StrategyTemplate }) {
  const data = assetClasses.map((asset) => ({ name: asset, current: weight(analytics.assetTotals[asset] ?? 0, analytics.total), target: template.targets[asset] }))
  return <ResponsiveContainer width="100%" height="100%"><BarChart data={data} layout="vertical"><CartesianGrid strokeDasharray="3 3" stroke="#27272a" /><XAxis type="number" stroke="#71717a" /><YAxis dataKey="name" type="category" width={150} stroke="#a1a1aa" tick={{ fontSize: 12 }} /><Tooltip formatter={(v) => `${percentFmt.format(Number(v))}%`} contentStyle={tooltipStyle} /><Legend /><Bar dataKey="current" fill="#34d399" radius={[0, 6, 6, 0]} /><Bar dataKey="target" fill="#52525b" radius={[0, 6, 6, 0]} /></BarChart></ResponsiveContainer>
}
function BeforeAfter({ before, after }: { before: { name: string; value: number }[]; after: { name: string; value: number }[] }) {
  const names = Array.from(new Set([...before.map((i) => i.name), ...after.map((i) => i.name)]))
  const data = names.map((name) => ({ name, current: before.find((i) => i.name === name)?.value ?? 0, simulated: after.find((i) => i.name === name)?.value ?? 0 }))
  return <ResponsiveContainer width="100%" height="100%"><BarChart data={data} layout="vertical"><XAxis type="number" stroke="#71717a" /><YAxis dataKey="name" type="category" width={150} stroke="#a1a1aa" tick={{ fontSize: 12 }} /><Tooltip contentStyle={tooltipStyle} /><Legend /><Bar dataKey="current" fill="#71717a" /><Bar dataKey="simulated" fill="#34d399" /></BarChart></ResponsiveContainer>
}
function PreviewRows({ rows }: { rows: Record<string, unknown>[] }) {
  const keys = Object.keys(rows[0] ?? {}).slice(0, 6)
  return <div className="md:col-span-2 overflow-auto rounded-xl border border-white/10"><table className="data-table"><thead><tr>{keys.map((key) => <th key={key}>{key}</th>)}</tr></thead><tbody>{rows.slice(0, 5).map((row, i) => <tr key={i}>{keys.map((key) => <td key={key}>{String(row[key] ?? '')}</td>)}</tr>)}</tbody></table></div>
}
function ClassificationEditor({ holding, updateHolding }: { holding: Holding; current: PortfolioSnapshot; updateHolding: (holding: Holding) => void }) {
  const [draft, setDraft] = useState(holding)
  return <div className="space-y-3">
    <div><p className="text-lg font-semibold">{holding.ticker}</p><p className="text-sm text-zinc-500">{holding.securityName}</p></div>
    <label className="field"><span>AI exposure score 0-5</span><input className="control" type="number" min={0} max={5} step={0.5} value={draft.ai.score} onChange={(e) => setDraft({ ...draft, ai: { ...draft.ai, score: Number(e.target.value), source: 'manual' } })} /></label>
    <label className="field"><span>Primary AI bucket</span><select className="control" value={draft.ai.buckets[0]?.bucket ?? aiBuckets[0]} onChange={(e) => setDraft({ ...draft, ai: { ...draft.ai, source: 'manual', buckets: [{ bucket: e.target.value as AIBucket, weight: 100 }] } })}>{aiBuckets.map((bucket) => <option key={bucket}>{bucket}</option>)}</select></label>
    <label className="field"><span>Direct or indirect</span><select className="control" value={draft.ai.directness} onChange={(e) => setDraft({ ...draft, ai: { ...draft.ai, directness: e.target.value as Directness, source: 'manual' } })}><option>direct</option><option>indirect</option><option>none</option></select></label>
    <label className="field"><span>Notes</span><textarea className="control min-h-24" value={draft.ai.notes} onChange={(e) => setDraft({ ...draft, ai: { ...draft.ai, notes: e.target.value, source: 'manual' } })} /></label>
    <button className="primary w-full justify-center" onClick={() => updateHolding(draft)}><PencilSimple size={16} /> Save classification</button>
  </div>
}
function ActionTable({ actions }: { actions: RecompCandidate[] }) {
  if (!actions.length) return <div className="empty"><Warning size={28} /><p>No simulated actions yet. Generate auto-recomp candidates or add manual sandbox actions.</p></div>
  return <div className="overflow-auto rounded-xl border border-white/10"><table className="data-table"><thead><tr><th>Action</th><th>Ticker / bucket</th><th>Amount</th><th>Before</th><th>After</th><th>Reason</th><th>Impact</th></tr></thead><tbody>{actions.map((action) => <tr key={action.id}><td>{action.actionType}</td><td>{action.tickerOrBucket}</td><td>{dollarFmt.format(action.dollarAmount)}</td><td>{percentFmt.format(action.beforeWeight)}%</td><td>{percentFmt.format(action.afterWeight)}%</td><td>{action.reason}</td><td>{action.alignmentImpact}</td></tr>)}</tbody></table></div>
}

const tooltipStyle = { background: '#18181b', border: '1px solid rgba(255,255,255,.1)', color: '#f4f4f5', borderRadius: 12 }

function analyze(holdings: Holding[], template: StrategyTemplate, stress: StressScenario, minDollar: number, minWeight: number) {
  const total = holdings.reduce((sum, h) => sum + h.marketValue, 0)
  const assetTotals = sumBy(holdings, (h) => h.assetClass)
  const sectorTotals = sumBy(holdings, (h) => h.sector)
  const aiBucketTotals: Record<string, number> = {}
  const directnessTotals: Record<string, number> = {}
  holdings.forEach((holding) => {
    if (holding.ai.score > 0) {
      const aiValue = holding.marketValue * (holding.ai.score / 5)
      directnessTotals[holding.ai.directness] = (directnessTotals[holding.ai.directness] ?? 0) + aiValue
      holding.ai.buckets.forEach((bucket) => {
        aiBucketTotals[bucket.bucket] = (aiBucketTotals[bucket.bucket] ?? 0) + aiValue * (bucket.weight / 100)
      })
    }
  })
  const aiValue = Object.values(aiBucketTotals).reduce((sum, value) => sum + value, 0)
  const topHoldings = [...holdings].sort((a, b) => b.marketValue - a.marketValue)
  const top10Weight = weight(topHoldings.slice(0, 10).reduce((sum, h) => sum + h.marketValue, 0), total)
  const largestWeight = weight(topHoldings[0]?.marketValue ?? 0, total)
  const riskScore = holdings.reduce((sum, h) => sum + weight(h.marketValue, total) / 100 * riskMultiplier(h), 0)
  const stressImpact = holdings.reduce((sum, h) => sum + weight(h.marketValue, total) / 100 * shockFor(h, stress), 0) * 100
  const assetClassData = toPercentData(assetTotals, total)
  const sectorData = toPercentData(sectorTotals, total)
  const aiBucketData = toPercentData(aiBucketTotals, aiValue || total)
  const directnessData = toPercentData(directnessTotals, aiValue || total)
  const driftScore = assetClasses.reduce((sum, asset) => sum + Math.abs(weight(assetTotals[asset] ?? 0, total) - template.targets[asset]), 0)
  const warnings = buildWarnings({ holdings, total, top10Weight, largestWeight, aiExposure: weight(aiValue, total), aiBucketData, template, assetTotals, minDollar, minWeight })
  return { holdings, total, assetTotals, sectorTotals, aiBucketTotals, aiValue, aiExposure: weight(aiValue, total), topHoldings, top10Weight, largestWeight, riskScore, stressImpact, assetClassData, sectorData, aiBucketData, directnessData, driftScore, warnings }
}
function buildWarnings(input: { holdings: Holding[]; total: number; top10Weight: number; largestWeight: number; aiExposure: number; aiBucketData: { name: string; value: number }[]; template: StrategyTemplate; assetTotals: Record<string, number>; minDollar: number; minWeight: number }): AppWarning[] {
  const warnings: AppWarning[] = []
  if (input.largestWeight > input.template.maxSingle) warnings.push({ id: 'single', severity: 'high', title: 'Single holding above template max', detail: `${percentFmt.format(input.largestWeight)}% is above the ${input.template.maxSingle}% guardrail.` })
  if (input.top10Weight > input.template.maxTop10) warnings.push({ id: 'top10', severity: 'medium', title: 'Top 10 holdings concentrated', detail: `Top 10 holdings are ${percentFmt.format(input.top10Weight)}% vs ${input.template.maxTop10}% target limit.` })
  if (input.aiExposure > input.template.aiRange[1]) warnings.push({ id: 'ai-high', severity: 'medium', title: 'AI sleeve above target', detail: `${percentFmt.format(input.aiExposure)}% AI exposure is above the ${input.template.aiRange[1]}% range high.` })
  if (input.aiExposure < input.template.aiRange[0]) warnings.push({ id: 'ai-low', severity: 'low', title: 'AI sleeve below target', detail: `${percentFmt.format(input.aiExposure)}% AI exposure is below the ${input.template.aiRange[0]}% range low.` })
  const biggestBucket = input.aiBucketData[0]
  if (biggestBucket && biggestBucket.value > input.template.maxBucket) warnings.push({ id: 'bucket', severity: 'medium', title: 'One AI bucket dominates', detail: `${biggestBucket.name} is ${percentFmt.format(biggestBucket.value)}% of AI exposure.` })
  const cash = weight(input.assetTotals.Cash ?? 0, input.total)
  if (cash > 12) warnings.push({ id: 'cash-high', severity: 'low', title: 'High cash', detail: `Cash is ${percentFmt.format(cash)}%.` })
  if (cash < 2) warnings.push({ id: 'cash-low', severity: 'low', title: 'Low cash', detail: `Cash is ${percentFmt.format(cash)}%.` })
  const missing = input.holdings.filter((h) => h.ai.source === 'unknown' || !h.sector || !h.assetClass).length
  if (missing) warnings.push({ id: 'missing', severity: 'low', title: 'Missing classification data', detail: `${missing} holdings need AI, sector, or asset class cleanup.` })
  const nuisance = input.holdings.filter((h) => h.marketValue < input.minDollar || weight(h.marketValue, input.total) < input.minWeight).length
  if (nuisance) warnings.push({ id: 'nuisance', severity: 'low', title: 'Tiny nuisance positions', detail: `${nuisance} positions are below the consolidation threshold.` })
  return warnings
}
function buildRecompCandidates(analytics: ReturnType<typeof analyze>, template: StrategyTemplate): RecompCandidate[] {
  const actions: RecompCandidate[] = []
  const minTrade = Math.max(1000, analytics.total * 0.003)
  analytics.topHoldings.forEach((holding) => {
    const before = weight(holding.marketValue, analytics.total)
    if (before > template.maxSingle) {
      const targetValue = analytics.total * (template.maxSingle / 100)
      const amount = holding.marketValue - targetValue
      if (amount > minTrade) actions.push(candidate('Trim to target %', holding.ticker, amount, before, template.maxSingle, `Trim ${holding.ticker} toward the ${template.maxSingle}% max single-position guardrail.`, holding.price))
    }
  })
  analytics.aiBucketData.slice(0, 3).forEach((bucket) => {
    if (bucket.value > template.maxBucket) actions.push(candidate('Allocate cash to bucket', `Reduce ${bucket.name}`, analytics.total * 0.015, bucket.value, template.maxBucket, `${bucket.name} dominates the AI sleeve; redirect incremental exposure elsewhere.`))
  })
  assetClasses.forEach((asset) => {
    const current = weight(analytics.assetTotals[asset] ?? 0, analytics.total)
    const target = template.targets[asset]
    if (target - current > 3) actions.push(candidate('Allocate cash to bucket', asset, Math.min(analytics.total * ((target - current) / 100), analytics.total * 0.04), current, target, `${asset} is underweight versus the selected template.`))
  })
  analytics.holdings.filter((h) => h.marketValue < 1500).slice(0, 4).forEach((holding) => actions.push(candidate('Sell dollar amount', holding.ticker, holding.marketValue, weight(holding.marketValue, analytics.total), 0, `Consolidate tiny nuisance position to reduce portfolio clutter.`, holding.price)))
  return actions.slice(0, 12)
}
function candidate(actionType: ActionType, tickerOrBucket: string, dollarAmount: number, beforeWeight: number, afterWeight: number, reason: string, price?: number): RecompCandidate {
  return { id: crypto.randomUUID(), actionType, tickerOrBucket, dollarAmount, estimatedShares: price ? dollarAmount / price : undefined, beforeWeight, afterWeight, reason, riskImpact: 'Estimated risk-budget contribution recalculates in the simulated view.', alignmentImpact: 'Moves closer to selected template guardrails.', stressImpact: 'Stress-test estimate updates after applying this simulated action.', warning: actionType.includes('Buy') && afterWeight > beforeWeight ? 'Check concentration before acting.' : undefined }
}
function simulateCandidates(holdings: Holding[], actions: RecompCandidate[]) {
  const simulated = holdings.map((h) => ({ ...h }))
  actions.forEach((action) => {
    const item = simulated.find((h) => h.ticker === action.tickerOrBucket)
    if (!item) return
    const direction = action.actionType.includes('Buy') ? 1 : -1
    item.marketValue = Math.max(0, item.marketValue + direction * action.dollarAmount)
    if (item.price) item.shares = item.marketValue / item.price
  })
  return simulated.filter((h) => h.marketValue > 1)
}
function normalizeImportedRow(row: Record<string, unknown>, accountOwner: string, accountName: string, accountType: string, index: number): Holding | null {
  const ticker = clean(row.Symbol ?? row.Ticker ?? row.ticker)
  const marketValue = parseMoney(row['Mkt Val (Market Value)'] ?? row.market_value ?? row['Market Value'])
  if (!ticker || !Number.isFinite(marketValue) || marketValue <= 0) return null
  const shares = parseMoney(row['Qty (Quantity)'] ?? row.shares ?? row.Quantity) || 0
  const price = parseMoney(row.Price)
  const defaults = defaultClassifications[ticker] ?? {}
  return { id: `${accountName}-${ticker}-${index}-${crypto.randomUUID()}`, accountOwner, accountName: accountName.replace(/^Positions for account\s*/i, ''), accountType, ticker, securityName: clean(row.Description ?? row.security_name ?? row.Name) || ticker, shares, price, marketValue, assetClass: (defaults.assetClass as AssetClass) ?? assetFromRaw(clean(row['Asset Type'])), sector: defaults.sector ?? 'Unclassified', ai: (defaults.ai as AIExposureClassification) ?? { score: 0, buckets: [], directness: 'none', confidence: 'low', source: 'unknown', notes: '' }, costBasis: parseMoney(row['Cost Basis']), unrealizedGainLoss: parseMoney(row['Gain $ (Gain/Loss $)']) }
}
function compareSnapshots(a: PortfolioSnapshot, b: PortfolioSnapshot) {
  const mapA = sumBy(a.holdings, (h) => h.ticker)
  const mapB = sumBy(b.holdings, (h) => h.ticker)
  const tickers = Array.from(new Set([...Object.keys(mapA), ...Object.keys(mapB)]))
  return tickers.map((ticker) => {
    const change = (mapB[ticker] ?? 0) - (mapA[ticker] ?? 0)
    return { name: ticker, change, status: !mapA[ticker] ? 'new' : !mapB[ticker] ? 'removed' : change >= 0 ? 'increased' : 'decreased' }
  }).sort((x, y) => Math.abs(y.change) - Math.abs(x.change)).slice(0, 20)
}
function executiveSummary(analytics: ReturnType<typeof analyze>, template: StrategyTemplate) {
  const driver = analytics.aiBucketData[0]?.name ?? 'broad equity'
  return `This portfolio is growth-oriented with an AI buildout exposure estimate of ${percentFmt.format(analytics.aiExposure)}%, compared with the ${template.name} target range of ${template.aiRange[0]}-${template.aiRange[1]}%. Top 10 holdings represent ${percentFmt.format(analytics.top10Weight)}% of the portfolio. Risk is mostly driven by ${driver}. The cleanest simulated next action is to reduce single-name concentration, resolve missing classifications, and redirect drift toward underweight template sleeves.`
}
function sumBy(holdings: Holding[], key: (h: Holding) => string): Record<string, number> {
  return holdings.reduce((acc, h) => ({ ...acc, [key(h)]: (acc[key(h)] ?? 0) + h.marketValue }), {} as Record<string, number>)
}
function toPercentData(totals: Record<string, number>, total: number) {
  return Object.entries(totals).map(([name, value]) => ({ name, value: weight(value, total) })).sort((a, b) => b.value - a.value)
}
function weight(value: number, total: number) { return total ? (value / total) * 100 : 0 }
function parseMoney(value: unknown) { const cleanValue = String(value ?? '').replace(/[$,%"]/g, '').replace(/,/g, '').trim(); const parsed = Number(cleanValue); return Number.isFinite(parsed) ? parsed : 0 }
function clean(value: unknown) { return String(value ?? '').replaceAll('"', '').trim() }
function assetFromRaw(raw: string): AssetClass { return raw.toLowerCase().includes('cash') ? 'Cash' : raw.toLowerCase().includes('bond') ? 'Bonds/fixed income' : 'Broad US equity' }
function riskMultiplier(h: Holding) {
  if (h.assetClass === 'Cash') return 0
  if (h.assetClass === 'Bonds/fixed income') return 0.3
  if (h.assetClass === 'Alternatives/other') return 0.8
  const bucket = h.ai.buckets[0]?.bucket
  if (bucket && ['Semiconductors', 'GPUs / accelerators', 'Photonics / optical / interconnect', 'Networking', 'Semiconductor equipment'].includes(bucket)) return 1.5
  if (bucket && ['Power generation', 'Grid / electrification', 'Data centers / colocation', 'Cooling / thermal management'].includes(bucket)) return 1.3
  if (bucket && ['Big tech / hyperscalers', 'AI platforms / AI software', 'Cybersecurity'].includes(bucket)) return 1.2
  return 1
}
function shockFor(h: Holding, scenario: StressScenario) {
  const bucketShock = h.ai.buckets.map((bucket) => (scenario.shocks[bucket.bucket] ?? 0) * (bucket.weight / 100)).reduce((sum, value) => sum + value, 0)
  return bucketShock || scenario.shocks[h.assetClass] || 0
}
function exportCsv(name: string, rows: unknown[]) {
  const blob = new Blob([Papa.unparse(rows as Record<string, unknown>[])], { type: 'text/csv;charset=utf-8' })
  download(name, blob)
}
function exportJson(name: string, data: unknown) { download(name, new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' })) }
function download(name: string, blob: Blob) { const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = name; a.click(); URL.revokeObjectURL(url) }
function loadState() {
  try {
    const parsed = JSON.parse(localStorage.getItem(storageKey) || '{}')
    return { snapshots: parsed.snapshots?.length ? parsed.snapshots : sampleSnapshots(), selectedSnapshotId: parsed.selectedSnapshotId ?? 'sample-current', selectedTemplateId: parsed.selectedTemplateId ?? 'diversified-ai-supply-chain', recompCandidates: parsed.recompCandidates ?? [], manualActions: parsed.manualActions ?? [], decisionLog: parsed.decisionLog ?? [] }
  } catch {
    return { snapshots: sampleSnapshots(), selectedSnapshotId: 'sample-current', selectedTemplateId: 'diversified-ai-supply-chain', recompCandidates: [], manualActions: [], decisionLog: [] }
  }
}
function persist(state: { snapshots: PortfolioSnapshot[]; selectedSnapshotId: string; selectedTemplateId: string; recompCandidates: RecompCandidate[]; manualActions: ManualSandboxAction[]; decisionLog: DecisionLogEntry[] }) {
  localStorage.setItem(storageKey, JSON.stringify(state))
}

export default App
