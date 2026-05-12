import { useCallback, useEffect, useMemo, useState } from 'react'
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
  LabelList,
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
import {
  ai,
  aiBuckets,
  assetClasses,
  assetClassColors,
  semanticColors,
  strategyTemplates,
  stressScenarios,
} from './domain'
import { chartColor, holdingColor, holdingColorLegend } from './chartSemantics'
import type {
  AIBucket,
  AIExposureClassification,
  AppWarning,
  AssetClass,
  ActionType,
  DecisionLogEntry,
  Directness,
  Holding,
  LedgerMode,
  ManualSandboxAction,
  PersistedState,
  PositionPlan,
  PositionPlanStyle,
  PortfolioSnapshot,
  RecompCandidate,
  SecurityType,
  StrategyTemplate,
  StressScenario,
  ThemePreference,
} from './domain'
import './index.css'

const chartAxis = 'var(--chart-axis)'
const chartGrid = 'var(--chart-grid)'
const chartTick = { fill: chartAxis, fontSize: 11, fontFamily: 'Geist, "Geist Sans", Aptos, "Segoe UI", system-ui, sans-serif' }
const chartTooltip = {
  background: 'var(--tooltip-bg)',
  border: '1px solid var(--border)',
  color: 'var(--app-text)',
  borderRadius: 12,
}
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
  ALAB: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Networking', 4.5, 'direct', [{ bucket: 'Data centers / colocation', weight: 20 }], 'Connectivity silicon for AI/cloud infrastructure.') },
  AMSC: { assetClass: 'AI buildout sleeve', sector: 'Industrials', ai: ai('Grid / electrification', 3, 'indirect', [], 'Grid and power infrastructure beneficiary.') },
  FLNC: { assetClass: 'AI buildout sleeve', sector: 'Industrials', ai: ai('Grid / electrification', 3, 'indirect', [{ bucket: 'Power generation', weight: 20 }], 'Energy storage/grid reliability exposure.') },
  GEV: { assetClass: 'AI buildout sleeve', sector: 'Industrials', ai: ai('Power generation', 3, 'indirect', [{ bucket: 'Grid / electrification', weight: 35 }]) },
  NBIS: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Data centers / colocation', 4.5, 'direct', [{ bucket: 'GPUs / accelerators', weight: 30 }], 'AI cloud and GPU infrastructure exposure.') },
  ORCL: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Big tech / hyperscalers', 4, 'direct', [{ bucket: 'Enterprise software', weight: 25 }]) },
  PLTR: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('AI platforms / AI software', 4, 'direct', [{ bucket: 'Enterprise software', weight: 30 }]) },
  META: { assetClass: 'AI buildout sleeve', sector: 'Communication Services', ai: ai('Big tech / hyperscalers', 4, 'direct') },
  TSLA: { assetClass: 'AI buildout sleeve', sector: 'Consumer Discretionary', ai: ai('Robotics / automation', 3, 'indirect', [{ bucket: 'AI platforms / AI software', weight: 25 }]) },
  TEM: { assetClass: 'AI buildout sleeve', sector: 'Health Care', ai: ai('AI platforms / AI software', 3, 'direct') },
  SOUN: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('AI platforms / AI software', 3, 'direct') },
  IQEPF: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Photonics / optical / interconnect', 4, 'direct', [{ bucket: 'Semiconductors', weight: 30 }], 'Compound semiconductor wafer and photonics supply-chain exposure.') },
  SIVEF: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Photonics / optical / interconnect', 4, 'direct', [{ bucket: 'Semiconductors', weight: 30 }], 'Photonics/wireless semiconductor supply-chain exposure.') },
  LPKFF: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Semiconductor equipment', 3, 'indirect', [{ bucket: 'Photonics / optical / interconnect', weight: 25 }]) },
  POET: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Photonics / optical / interconnect', 4, 'direct', [{ bucket: 'Data centers / colocation', weight: 25 }]) },
  LASR: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Photonics / optical / interconnect', 3, 'direct') },
  NVEC: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Memory / storage', 2, 'indirect') },
  NOK: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Networking', 2.5, 'indirect') },
  JBL: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Data centers / colocation', 2.5, 'indirect', [{ bucket: 'Networking', weight: 25 }]) },
  PH: { assetClass: 'AI buildout sleeve', sector: 'Industrials', ai: ai('Cooling / thermal management', 2.5, 'indirect') },
  WCC: { assetClass: 'AI buildout sleeve', sector: 'Industrials', ai: ai('Grid / electrification', 2.5, 'indirect', [{ bucket: 'Data centers / colocation', weight: 25 }]) },
  MP: { assetClass: 'AI buildout sleeve', sector: 'Materials', ai: ai('Materials / specialty chemicals', 2.5, 'indirect') },
  ARAFF: { assetClass: 'AI buildout sleeve', sector: 'Materials', ai: ai('Materials / specialty chemicals', 2, 'indirect') },
  SOLS: { assetClass: 'AI buildout sleeve', sector: 'Materials', ai: ai('Materials / specialty chemicals', 2, 'indirect') },
  DRAM: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('Memory / storage', 4, 'indirect', [{ bucket: 'Semiconductors', weight: 30 }]) },
  NVDL: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('GPUs / accelerators', 5, 'indirect', [{ bucket: 'Semiconductors', weight: 40 }]) },
  NVDY: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('GPUs / accelerators', 4, 'indirect', [{ bucket: 'Semiconductors', weight: 35 }]) },
  PLTY: { assetClass: 'AI buildout sleeve', sector: 'Information Technology', ai: ai('AI platforms / AI software', 3, 'indirect') },
  'BRK/B': { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  QQQ: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1.5, 'indirect') },
  SPY: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  DIA: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 0.5, 'indirect') },
  RSP: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 0.8, 'indirect') },
  VUG: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1.5, 'indirect') },
  JEPQ: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1.5, 'indirect') },
  GPIQ: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1.5, 'indirect') },
  QQQI: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1.5, 'indirect') },
  QYLD: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1.5, 'indirect') },
  XYLD: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  DJIA: { assetClass: 'Broad US equity', securityType: 'ETF / closed-end fund', sector: 'Broad Market', ai: ai('Broad passive index exposure', 0.5, 'indirect') },
  NOBL: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 0.5, 'indirect') },
  VYM: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 0.5, 'indirect') },
  VHYAX: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 0.5, 'indirect') },
  XLU: { assetClass: 'AI buildout sleeve', sector: 'Utilities', ai: ai('Power generation', 1.5, 'indirect', [{ bucket: 'Grid / electrification', weight: 35 }]) },
  URA: { assetClass: 'AI buildout sleeve', sector: 'Utilities', ai: ai('Power generation', 2, 'indirect') },
  COPX: { assetClass: 'AI buildout sleeve', sector: 'Materials', ai: ai('Grid / electrification', 1.5, 'indirect', [{ bucket: 'Materials / specialty chemicals', weight: 35 }]) },
  ICOP: { assetClass: 'AI buildout sleeve', sector: 'Materials', ai: ai('Grid / electrification', 1.5, 'indirect', [{ bucket: 'Materials / specialty chemicals', weight: 35 }]) },
  XLE: { assetClass: 'Alternatives/other', sector: 'Energy', ai: ai('Power generation', 0.5, 'indirect') },
  GDX: { assetClass: 'Alternatives/other', securityType: 'ETF / closed-end fund', sector: 'Materials', ai: ai('Broad passive index exposure', 0, 'none') },
  GDXJ: { assetClass: 'Alternatives/other', securityType: 'ETF / closed-end fund', sector: 'Materials', ai: ai('Broad passive index exposure', 0, 'none') },
  PPA: { assetClass: 'Alternatives/other', securityType: 'ETF / closed-end fund', sector: 'Industrials', ai: ai('Robotics / automation', 1, 'indirect', [], 'Defense/aerospace ETF with indirect automation and autonomous systems exposure.') },
  TSLY: { assetClass: 'Broad US equity', securityType: 'ETF / closed-end fund', sector: 'Consumer Discretionary', ai: ai('Robotics / automation', 1, 'indirect', [], 'Option-income ETF tied to TSLA exposure; not direct AI infrastructure ownership.') },
  DXYZ: { assetClass: 'Alternatives/other', securityType: 'ETF / closed-end fund', sector: 'Private technology', ai: ai('Public indirect AI lab exposure', 1.5, 'indirect', [], 'Closed-end public vehicle with private technology exposure; classify as indirect only.') },
  ETHA: { assetClass: 'Alternatives/other', sector: 'Digital assets', ai: ai('Broad passive index exposure', 0, 'none') },
  IBIT: { assetClass: 'Alternatives/other', sector: 'Digital assets', ai: ai('Broad passive index exposure', 0, 'none') },
  IREN: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Digital assets', ai: ai('Data centers / colocation', 1.5, 'indirect') },
  DGXX: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Digital assets', ai: ai('Data centers / colocation', 1, 'indirect') },
  BMNR: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Digital assets', ai: ai('Broad passive index exposure', 0, 'none') },
  AZO: { assetClass: 'Broad US equity', sector: 'Consumer Discretionary', ai: ai('Broad passive index exposure', 0, 'none') },
  COST: { assetClass: 'Broad US equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  KR: { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  SFD: { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  UNFI: { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  WMT: { assetClass: 'Broad US equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  KO: { assetClass: 'Broad US equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  HSY: { assetClass: 'Broad US equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  KHC: { assetClass: 'Broad US equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  MO: { assetClass: 'Broad US equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  BTI: { assetClass: 'International equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  BUD: { assetClass: 'International equity', sector: 'Consumer Staples', ai: ai('Broad passive index exposure', 0, 'none') },
  JPM: { assetClass: 'Broad US equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  AB: { assetClass: 'Broad US equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  BEN: { assetClass: 'Broad US equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  COIN: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  EWBC: { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  FRCB: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  IVZ: { assetClass: 'Broad US equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  HOOD: { assetClass: 'Broad US equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  SOFI: { assetClass: 'Broad US equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  MA: { assetClass: 'Broad US equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  V: { assetClass: 'Broad US equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  AGNC: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Real Estate', ai: ai('Broad passive index exposure', 0, 'none') },
  O: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Real Estate', ai: ai('Data centers / colocation', 0.5, 'indirect') },
  MITT: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Real Estate', ai: ai('Broad passive index exposure', 0, 'none') },
  SAR: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  CVX: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Energy', ai: ai('Power generation', 0.5, 'indirect') },
  APA: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Energy', ai: ai('Power generation', 0.5, 'indirect') },
  EPD: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Energy', ai: ai('Power generation', 0.5, 'indirect') },
  ET: { assetClass: 'Alternatives/other', securityType: 'Individual equity', sector: 'Energy', ai: ai('Power generation', 0.5, 'indirect') },
  BHP: { assetClass: 'International equity', sector: 'Materials', ai: ai('Materials / specialty chemicals', 1, 'indirect') },
  CRS: { assetClass: 'Broad US equity', sector: 'Materials', ai: ai('Materials / specialty chemicals', 1, 'indirect') },
  CAT: { assetClass: 'Broad US equity', sector: 'Industrials', ai: ai('Grid / electrification', 1, 'indirect') },
  DE: { assetClass: 'Broad US equity', sector: 'Industrials', ai: ai('Robotics / automation', 1, 'indirect') },
  GD: { assetClass: 'Broad US equity', sector: 'Industrials', ai: ai('Broad passive index exposure', 0, 'none') },
  LMT: { assetClass: 'Broad US equity', sector: 'Industrials', ai: ai('Broad passive index exposure', 0, 'none') },
  MMM: { assetClass: 'Broad US equity', sector: 'Industrials', ai: ai('Broad passive index exposure', 0, 'none') },
  F: { assetClass: 'Broad US equity', sector: 'Consumer Discretionary', ai: ai('Robotics / automation', 0.5, 'indirect') },
  GM: { assetClass: 'Broad US equity', sector: 'Consumer Discretionary', ai: ai('Robotics / automation', 0.5, 'indirect') },
  WYNN: { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Consumer Discretionary', ai: ai('Broad passive index exposure', 0, 'none') },
  RL: { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Consumer Discretionary', ai: ai('Broad passive index exposure', 0, 'none') },
  WSM: { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Consumer Discretionary', ai: ai('Broad passive index exposure', 0, 'none') },
  CVNA: { assetClass: 'Broad US equity', sector: 'Consumer Discretionary', ai: ai('Broad passive index exposure', 0, 'none') },
  EBAY: { assetClass: 'Broad US equity', sector: 'Consumer Discretionary', ai: ai('Enterprise software', 0.5, 'indirect') },
  MELI: { assetClass: 'International equity', sector: 'Consumer Discretionary', ai: ai('Enterprise software', 0.5, 'indirect') },
  NFLX: { assetClass: 'Broad US equity', sector: 'Communication Services', ai: ai('AI platforms / AI software', 1, 'indirect') },
  LLY: { assetClass: 'Broad US equity', sector: 'Health Care', ai: ai('AI platforms / AI software', 0.5, 'indirect') },
  CLOV: { assetClass: 'Broad US equity', sector: 'Health Care', ai: ai('Enterprise software', 0.5, 'indirect') },
  RKLB: { assetClass: 'Broad US equity', sector: 'Industrials', ai: ai('Robotics / automation', 1, 'indirect') },
  QS: { assetClass: 'Broad US equity', sector: 'Consumer Discretionary', ai: ai('Grid / electrification', 1, 'indirect') },
  ONDS: { assetClass: 'AI buildout sleeve', securityType: 'Individual equity', sector: 'Information Technology', ai: ai('Robotics / automation', 2.5, 'direct', [{ bucket: 'Networking', weight: 25 }], 'Autonomous systems and private industrial wireless exposure; speculative/small-cap profile.') },
  OPEN: { assetClass: 'Broad US equity', securityType: 'Individual equity', sector: 'Real Estate', ai: ai('Enterprise software', 0.5, 'indirect') },
  OPENL: { assetClass: 'Alternatives/other', securityType: 'Option / warrant / right', sector: 'Real Estate', ai: ai('Enterprise software', 0.5, 'indirect') },
  OPENW: { assetClass: 'Alternatives/other', securityType: 'Option / warrant / right', sector: 'Real Estate', ai: ai('Enterprise software', 0.5, 'indirect') },
  OPENZ: { assetClass: 'Alternatives/other', securityType: 'Option / warrant / right', sector: 'Real Estate', ai: ai('Enterprise software', 0.5, 'indirect') },
  CGCTW: { assetClass: 'Alternatives/other', securityType: 'Option / warrant / right', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  VTI: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  VOO: { assetClass: 'Broad US equity', sector: 'Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  VXUS: { assetClass: 'International equity', sector: 'International Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  DXJ: { assetClass: 'International equity', securityType: 'ETF / closed-end fund', sector: 'International Broad Market', ai: ai('Broad passive index exposure', 0.5, 'indirect') },
  EWY: { assetClass: 'International equity', securityType: 'ETF / closed-end fund', sector: 'International Broad Market', ai: ai('Broad passive index exposure', 1, 'indirect') },
  SCHE: { assetClass: 'International equity', securityType: 'ETF / closed-end fund', sector: 'International Broad Market', ai: ai('Broad passive index exposure', 0.5, 'indirect') },
  SMFG: { assetClass: 'International equity', securityType: 'Individual equity', sector: 'Financials', ai: ai('Broad passive index exposure', 0, 'none') },
  TX: { assetClass: 'International equity', securityType: 'Individual equity', sector: 'Materials', ai: ai('Materials / specialty chemicals', 0.5, 'indirect') },
  BND: { assetClass: 'Bonds/fixed income', sector: 'Fixed Income', ai: ai('Broad passive index exposure', 0, 'none') },
  SGOV: { assetClass: 'Cash', sector: 'Cash & equivalents', ai: ai('Broad passive index exposure', 0, 'none') },
  SPAXX: { assetClass: 'Cash', sector: 'Cash & equivalents', ai: ai('Broad passive index exposure', 0, 'none') },
  SNVXX: { assetClass: 'Cash', sector: 'Cash & equivalents', ai: ai('Broad passive index exposure', 0, 'none') },
  'SEAL/PRA': { assetClass: 'Alternatives/other', securityType: 'Other', sector: 'Energy', ai: ai('Broad passive index exposure', 0, 'none') },
  '89236TJK2': { assetClass: 'Bonds/fixed income', securityType: 'Bond / fixed income', sector: 'Fixed Income', ai: ai('Broad passive index exposure', 0, 'none') },
  'CASH & CASH INVESTMENTS': { assetClass: 'Cash', securityType: 'Cash / money market', sector: 'Cash & equivalents', ai: ai('Broad passive index exposure', 0, 'none') },
  'Cash & Cash Investments': { assetClass: 'Cash', sector: 'Cash & equivalents', ai: ai('Broad passive index exposure', 0, 'none') },
}

function sampleSnapshots(): PortfolioSnapshot[] {
  const holdings: Holding[] = [
    holding('Portfolio', 'Joint taxable', 'Taxable', 'VTI', 'Vanguard Total Stock Market ETF', 540, 247.8, 'Broad US equity', 'Broad Market'),
    holding('Portfolio', 'Joint taxable', 'Taxable', 'MSFT', 'Microsoft Corp', 220, 506.4, 'AI buildout sleeve', 'Information Technology'),
    holding('Portfolio', 'Joint taxable', 'Taxable', 'NVDA', 'NVIDIA Corp', 390, 187.6, 'AI buildout sleeve', 'Information Technology'),
    holding('Portfolio', 'Joint taxable', 'Taxable', 'ETN', 'Eaton Corp', 105, 421.3, 'AI buildout sleeve', 'Industrials'),
    holding('Portfolio', 'Joint taxable', 'Taxable', 'EQIX', 'Equinix Inc', 38, 812.4, 'AI buildout sleeve', 'Real Estate'),
    holding('Portfolio', 'Roth IRA', 'Roth IRA', 'AAPL', 'Apple Inc', 196, 293.05, 'AI buildout sleeve', 'Information Technology'),
    holding('Portfolio', 'Roth IRA', 'Roth IRA', 'ASML', 'ASML Holding NV', 22, 725.9, 'AI buildout sleeve', 'Information Technology'),
    holding('Portfolio', '401k', 'Retirement', 'BND', 'Vanguard Total Bond Market ETF', 710, 72.1, 'Bonds/fixed income', 'Fixed Income'),
    holding('Portfolio', '401k', 'Retirement', 'VXUS', 'Vanguard Total International Stock ETF', 620, 65.4, 'International equity', 'International Broad Market'),
    holding('Portfolio', '401k', 'Retirement', 'SGOV', 'iShares 0-3 Month Treasury Bond ETF', 260, 100.6, 'Cash', 'Cash & equivalents'),
    holding('Portfolio', 'DB Plan', 'Tax-advantaged', 'VRT', 'Vertiv Holdings', 135, 176.7, 'AI buildout sleeve', 'Industrials'),
    holding('Portfolio', 'DB Plan', 'Tax-advantaged', 'PANW', 'Palo Alto Networks', 55, 214.2, 'AI buildout sleeve', 'Information Technology'),
    holding('Portfolio', 'DB Plan', 'Tax-advantaged', 'CEG', 'Constellation Energy', 80, 351.8, 'AI buildout sleeve', 'Utilities'),
    holding('Portfolio', 'DB Plan', 'Tax-advantaged', 'AGNC', 'AGNC Investment Corp REIT', 300, 10.86, 'Alternatives/other', 'Real Estate'),
    holding('Portfolio', 'DB Plan', 'Tax-advantaged', 'PSHZF', 'Pershing Square Holdings Rights', 157, 0.35, 'Alternatives/other', 'Financials'),
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
    accountId: slug(accountName),
    accountOwner,
    accountName,
    accountType,
    accountCategory: detectAccountType(accountName),
    ticker,
    securityName,
    shares,
    price,
    marketValue: shares * price,
    assetClass: (defaults.assetClass as AssetClass) ?? assetClass,
    securityType: (defaults.securityType as SecurityType | undefined) ?? inferSecurityType(ticker, securityName, assetClass),
    sector: defaults.sector ?? sector,
    ai: (defaults.ai as AIExposureClassification) ?? { score: 0, buckets: [], directness: 'none', confidence: 'low', source: 'unknown', notes: '', },
    costBasis: shares * price * 0.78,
    unrealizedGainLoss: shares * price * 0.22,
  }
}

const dollarFmt = new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 })
const percentFmt = new Intl.NumberFormat('en-US', { maximumFractionDigits: 1 })
const shareFmt = new Intl.NumberFormat('en-US', { maximumFractionDigits: 2 })
const storageKey = 'ai-buildout-portfolio-recomp-state-v1'

function App() {
  const persisted = loadState()
  const [snapshots, setSnapshots] = useState<PortfolioSnapshot[]>(persisted.snapshots)
  const [selectedSnapshotId, setSelectedSnapshotId] = useState(persisted.selectedSnapshotId)
  const [selectedAccountScope, setSelectedAccountScope] = useState(persisted.selectedAccountScope)
  const [activeTab, setActiveTab] = useState('Overview')
  const [selectedTemplateId, setSelectedTemplateId] = useState(persisted.selectedTemplateId)
  const [customTemplates, setCustomTemplates] = useState<StrategyTemplate[]>(persisted.customTemplates)
  const [themePreference, setThemePreference] = useState<ThemePreference>(persisted.themePreference)
  const [ledgerMode, setLedgerMode] = useState<LedgerMode>(persisted.ledgerMode)
  const [holdingEditOverlay, setHoldingEditOverlay] = useState<Record<string, Partial<Holding>>>(persisted.holdingEditOverlay)
  const [selectedStressId, setSelectedStressId] = useState('ai-disappointment')
  const [recompCandidates, setRecompCandidates] = useState<RecompCandidate[]>(persisted.recompCandidates)
  const [manualActions, setManualActions] = useState<ManualSandboxAction[]>(persisted.manualActions)
  const [positionPlans, setPositionPlans] = useState<Record<string, PositionPlan>>(persisted.positionPlans)
  const [hiddenHoldingIds, setHiddenHoldingIds] = useState<string[]>(persisted.hiddenHoldingIds)
  const [decisionLog, setDecisionLog] = useState<DecisionLogEntry[]>(persisted.decisionLog)
  const [search, setSearch] = useState('')
  const [grouping, setGrouping] = useState<GroupingState>([])
  const [sorting, setSorting] = useState<SortingState>([])
  const [assetFilter, setAssetFilter] = useState('All')
  const [sectorFilter, setSectorFilter] = useState('All')
  const [aiBucketFilter, setAiBucketFilter] = useState('All')
  const [sizeFilter, setSizeFilter] = useState('All sizes')
  const [minDollar, setMinDollar] = useState(1500)
  const [minWeight, setMinWeight] = useState(0.25)
  const [importRows, setImportRows] = useState<Record<string, unknown>[]>([])
  const [importName, setImportName] = useState('')
  const [importSnapshotDate, setImportSnapshotDate] = useState('')
  const [importAccountOwner, setImportAccountOwner] = useState('Portfolio')
  const [importAccountType, setImportAccountType] = useState('Tax-advantaged')
  const [importMessage, setImportMessage] = useState('')
  const [selectedHoldingId, setSelectedHoldingId] = useState<string | null>(null)

  const currentSnapshot = snapshots.find((snapshot) => snapshot.id === selectedSnapshotId) ?? snapshots.at(-1)!
  const accountOptions = useMemo(() => accountScopes(currentSnapshot.holdings), [currentSnapshot.holdings])
  const currentEffectiveHoldings = useMemo(() => applyHoldingEditOverlay(currentSnapshot.holdings, holdingEditOverlay), [currentSnapshot.holdings, holdingEditOverlay])
  const scopedSourceHoldings = useMemo(() => {
    const source = ledgerMode === 'sandbox' ? currentEffectiveHoldings : currentSnapshot.holdings
    return selectedAccountScope === 'combined' ? source : source.filter((holding) => holding.accountId === selectedAccountScope)
  }, [currentEffectiveHoldings, currentSnapshot.holdings, ledgerMode, selectedAccountScope])
  const includedSourceHoldings = useMemo(() => scopedSourceHoldings.filter((holding) => !hiddenHoldingIds.includes(holding.id)), [hiddenHoldingIds, scopedSourceHoldings])
  const scopedHoldings = useMemo(() => aggregateHoldingsForView(includedSourceHoldings), [includedSourceHoldings])
  const tableScopedHoldings = useMemo(() => aggregateHoldingsForView(scopedSourceHoldings), [scopedSourceHoldings])
  const hiddenHoldings = currentSnapshot.holdings.filter((holding) => hiddenHoldingIds.includes(holding.id))
  const editedCount = includedSourceHoldings.filter((holding) => isHoldingEdited(holding.id, holdingEditOverlay)).length
  const allTemplates = useMemo(() => [...strategyTemplates, ...customTemplates], [customTemplates])
  const template = allTemplates.find((item) => item.id === selectedTemplateId) ?? strategyTemplates[0]
  const stress = stressScenarios.find((item) => item.id === selectedStressId) ?? stressScenarios[2]
  const analytics = useMemo(() => analyze(scopedHoldings, template, stress, minDollar, minWeight), [scopedHoldings, template, stress, minDollar, minWeight])
  const simulated = useMemo(() => simulateCandidates(scopedHoldings, [...recompCandidates, ...manualActions]), [scopedHoldings, recompCandidates, manualActions])
  const simulatedAnalytics = useMemo(() => analyze(simulated, template, stress, minDollar, minWeight), [simulated, template, stress, minDollar, minWeight])
  const selectedHolding = scopedHoldings.find((item) => item.id === selectedHoldingId) ?? currentEffectiveHoldings.find((item) => item.id === selectedHoldingId)
  const cumulativeWeightByHoldingId = useMemo(() => cumulativeWeights(analytics.topHoldings, analytics.total), [analytics.topHoldings, analytics.total])
  const tableTotal = useMemo(() => tableScopedHoldings.reduce((sum, holding) => sum + holding.marketValue, 0), [tableScopedHoldings])
  const tableHoldings = useMemo(() => filterHoldingsForTable(tableScopedHoldings, assetFilter, sectorFilter, aiBucketFilter, sizeFilter, minDollar, minWeight, tableTotal), [aiBucketFilter, assetFilter, minDollar, minWeight, sectorFilter, sizeFilter, tableScopedHoldings, tableTotal])
  const tableFilterOptions = useMemo(() => tableFilters(tableScopedHoldings), [tableScopedHoldings])

  useEffect(() => {
    document.documentElement.dataset.theme = resolveTheme(themePreference)
  }, [themePreference])

  useEffect(() => {
    persist({ snapshots, selectedSnapshotId: currentSnapshot.id, selectedAccountScope, selectedTemplateId: template.id, customTemplates, themePreference, ledgerMode, holdingEditOverlay, hiddenHoldingIds, recompCandidates, manualActions, positionPlans, decisionLog })
  }, [currentSnapshot.id, customTemplates, decisionLog, hiddenHoldingIds, holdingEditOverlay, ledgerMode, manualActions, positionPlans, recompCandidates, selectedAccountScope, snapshots, template.id, themePreference])

  const commitHoldingEdit = useCallback((updated: Holding) => {
    setLedgerMode('sandbox')
    const sourceIds = updated.sourceHoldingIds?.length ? updated.sourceHoldingIds : [updated.id]
    const sourceRows = currentEffectiveHoldings.filter((holding) => sourceIds.includes(holding.id))
    const totalValue = sourceRows.reduce((sum, holding) => sum + holding.marketValue, 0)
    const totalShares = sourceRows.reduce((sum, holding) => sum + holding.shares, 0)
    setHoldingEditOverlay((overlay) => {
      const next = { ...overlay }
      sourceRows.forEach((source, index) => {
        const valueRatio = totalValue ? source.marketValue / totalValue : 1 / sourceRows.length
        const shareRatio = totalShares ? source.shares / totalShares : valueRatio
        const marketValue = updated.marketValue * valueRatio
        const shares = updated.shares * shareRatio
        next[source.id] = {
          ...source,
          assetClass: updated.assetClass,
          sector: updated.sector,
          ai: { ...updated.ai, buckets: [...updated.ai.buckets] },
          notes: updated.notes,
          marketValue,
          shares,
          price: updated.price ?? (shares ? marketValue / shares : source.price),
          securityName: index === 0 ? updated.securityName : source.securityName,
        }
      })
      return next
    })
  }, [currentEffectiveHoldings])

  const toggleHoldingIncluded = useCallback((holding: Holding, included: boolean) => {
    const ids = holding.sourceHoldingIds?.length ? holding.sourceHoldingIds : [holding.id]
    setHiddenHoldingIds((items) => included ? items.filter((id) => !ids.includes(id)) : [...new Set([...items, ...ids])])
  }, [])

  const columns = useMemo<ColumnDef<Holding>[]>(() => [
    { id: 'included', header: 'Included', enableSorting: false, cell: ({ row }) => {
      const included = !holdingExcludedInView(row.original, hiddenHoldingIds)
      return <input className="include-checkbox" type="checkbox" checked={included} title={included ? 'Included in analytics' : 'Excluded from analytics'} onClick={(event) => event.stopPropagation()} onChange={(event) => toggleHoldingIncluded(row.original, event.target.checked)} />
    } },
    { accessorKey: 'ticker', header: 'Ticker', cell: ({ row }) => <div className="flex items-center gap-2"><span>{row.original.ticker}</span>{holdingEditedInView(row.original, holdingEditOverlay) && <span className="edited-pill">Edited</span>}</div> },
    { accessorKey: 'securityName', header: 'Security' },
    { accessorKey: 'accountName', header: 'Account' },
    { accessorKey: 'shares', header: 'Shares', cell: ({ row }) => ledgerMode === 'sandbox' ? <NumberCell value={row.original.shares} step={0.0001} onChange={(value) => commitHoldingEdit(repriceHolding({ ...row.original, shares: value }))} /> : shareFmt.format(row.original.shares) },
    { accessorKey: 'price', header: 'Price', cell: ({ row }) => ledgerMode === 'sandbox' ? <NumberCell value={row.original.price ?? 0} step={0.01} onChange={(value) => commitHoldingEdit(repriceHolding({ ...row.original, price: value }))} /> : row.original.price ? dollarFmt.format(row.original.price) : '-' },
    { accessorKey: 'marketValue', header: 'Market value', cell: ({ row }) => ledgerMode === 'sandbox' ? <NumberCell value={row.original.marketValue} step={1} onChange={(value) => commitHoldingEdit({ ...row.original, marketValue: value })} /> : dollarFmt.format(row.original.marketValue) },
    { id: 'portfolioWeight', header: 'Portfolio %', accessorFn: (row) => weight(row.marketValue, analytics.total), cell: ({ getValue }) => `${percentFmt.format(Number(getValue()))}%` },
    { id: 'cumulativeWeight', header: 'Cumulative %', accessorFn: (row) => cumulativeWeightByHoldingId.get(row.id) ?? 0, cell: ({ getValue }) => `${percentFmt.format(Number(getValue()))}%` },
    { accessorKey: 'securityType', header: 'Type' },
    { accessorKey: 'assetClass', header: 'Asset class', cell: ({ row }) => ledgerMode === 'sandbox' ? <SelectCell value={row.original.assetClass} options={assetClasses} onChange={(value) => commitHoldingEdit({ ...row.original, assetClass: value as AssetClass })} /> : row.original.assetClass },
    { accessorKey: 'sector', header: 'Sector', cell: ({ row }) => ledgerMode === 'sandbox' ? <TextCell value={row.original.sector} onChange={(value) => commitHoldingEdit({ ...row.original, sector: value })} /> : row.original.sector },
    { id: 'aiBucket', header: 'AI bucket', accessorFn: (row) => row.ai.buckets[0]?.bucket ?? 'Missing', cell: ({ row, getValue }) => ledgerMode === 'sandbox' ? <SelectCell value={row.original.ai.buckets[0]?.bucket ?? aiBuckets[0]} options={aiBuckets} onChange={(value) => commitHoldingEdit({ ...row.original, ai: { ...row.original.ai, source: 'manual', buckets: [{ bucket: value as AIBucket, weight: 100 }] } })} /> : String(getValue()) },
    { id: 'aiScore', header: 'AI score', accessorFn: (row) => row.ai.score, cell: ({ row, getValue }) => ledgerMode === 'sandbox' ? <NumberCell value={row.original.ai.score} min={0} max={5} step={0.5} onChange={(value) => commitHoldingEdit({ ...row.original, ai: { ...row.original.ai, score: value, source: 'manual' } })} /> : Number(getValue()).toFixed(1) },
    { id: 'directness', header: 'Directness', accessorFn: (row) => row.ai.directness, cell: ({ row, getValue }) => ledgerMode === 'sandbox' ? <SelectCell value={row.original.ai.directness} options={['direct', 'indirect', 'none']} onChange={(value) => commitHoldingEdit({ ...row.original, ai: { ...row.original.ai, directness: value as Directness, source: 'manual' } })} /> : String(getValue()) },
    { accessorKey: 'notes', header: 'Notes', cell: ({ row }) => ledgerMode === 'sandbox' ? <TextCell value={row.original.notes ?? ''} onChange={(value) => commitHoldingEdit({ ...row.original, notes: value })} /> : row.original.notes || '-' },
    { id: 'warning', header: 'Warning', accessorFn: (row) => warningStatus(row, analytics.total, minDollar, minWeight, template.maxSingle), cell: ({ getValue }) => String(getValue()) },
  ], [analytics.total, commitHoldingEdit, cumulativeWeightByHoldingId, hiddenHoldingIds, holdingEditOverlay, ledgerMode, minDollar, minWeight, template.maxSingle, toggleHoldingIncluded])
  // eslint-disable-next-line react-hooks/incompatible-library
  const table = useReactTable({ data: tableHoldings, columns, state: { globalFilter: search, grouping, sorting }, onGlobalFilterChange: setSearch, onGroupingChange: setGrouping, onSortingChange: setSorting, getCoreRowModel: getCoreRowModel(), getFilteredRowModel: getFilteredRowModel(), getGroupedRowModel: getGroupedRowModel(), getSortedRowModel: getSortedRowModel() })

  function generateCandidates() {
    const candidates = buildRecompCandidates(analytics, template)
    setRecompCandidates(candidates)
    setDecisionLog((logs) => [{ id: crypto.randomUUID(), date: new Date().toISOString(), snapshotName: currentSnapshot.name, templateName: template.name, currentAIExposure: analytics.aiExposure, driftScore: analytics.driftScore, actions: candidates, warnings: analytics.warnings, notes: 'Auto-generated simulated recomp candidate run.' }, ...logs])
    setActiveTab('Recomp Sandbox')
  }

  function updateHolding(updated: Holding) {
    commitHoldingEdit(updated)
    setLedgerMode('sandbox')
    setActiveTab('Concentration & Holdings')
  }

  function discardHoldingEdits() {
    setHoldingEditOverlay({})
    setLedgerMode('uploaded')
  }

  function saveEditedSnapshot() {
    const editedHoldings = applyHoldingEditOverlay(currentSnapshot.holdings, holdingEditOverlay).map((holding) => ({
      ...holding,
      id: `${holding.id}-edited-${crypto.randomUUID()}`,
      ai: { ...holding.ai, buckets: [...holding.ai.buckets] },
    }))
    const snapshot: PortfolioSnapshot = {
      id: crypto.randomUUID(),
      name: `${currentSnapshot.name} edited`,
      date: new Date().toISOString(),
      source: `Edited sandbox from ${currentSnapshot.name}`,
      holdings: editedHoldings,
    }
    setSnapshots((items) => [...items, snapshot])
    setSelectedSnapshotId(snapshot.id)
    setHoldingEditOverlay({})
    setLedgerMode('uploaded')
    setSelectedHoldingId(null)
  }

  async function parseUpload(files: FileList | File) {
    const fileList = files instanceof File ? [files] : Array.from(files)
    const stagedRows: Record<string, unknown>[] = []
    for (const file of fileList) {
      const buffer = await file.arrayBuffer()
      const text = new TextDecoder().decode(buffer)
      let rows: Record<string, unknown>[]
      let sourceName = file.name
      let detected = detectAccountFromText(file.name, file.name)
      if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
        const workbook = XLSX.read(buffer)
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        rows = XLSX.utils.sheet_to_json(sheet, { defval: '' })
      } else {
        const lines = text.split(/\r?\n/)
        const accountLine = lines.find((line) => line.includes('Positions for account'))
        sourceName = accountLine?.replaceAll('"', '').trim() || file.name
        detected = detectAccountFromText(sourceName, file.name)
        const asOf = detectAsOfDate(sourceName)
        const headerIndex = lines.findIndex((line) => line.includes('Symbol') && line.includes('Description'))
        const parseText = headerIndex >= 0 ? lines.slice(headerIndex).join('\n') : text
        const parsed = Papa.parse<Record<string, unknown>>(parseText, { header: true, skipEmptyLines: true })
        rows = parsed.data
        if (asOf && !importSnapshotDate) setImportSnapshotDate(asOf)
      }
      rows.filter((row) => row.Symbol || row.ticker || row.Ticker).forEach((row) => stagedRows.push({ ...row, __accountName: detected.name, __accountType: detected.type, __accountId: detected.id, __source: sourceName, __asOfDate: detectAsOfDate(sourceName) }))
    }
    setImportRows(stagedRows)
    setImportName(fileList.length > 1 ? `Combined ${fileList.length}-account upload` : String(stagedRows[0]?.__source ?? fileList[0]?.name ?? 'Uploaded file'))
    setImportAccountType(String(stagedRows[0]?.__accountType ?? 'Tax-advantaged'))
    setImportMessage(`${stagedRows.length} rows detected across ${new Set(stagedRows.map((row) => row.__accountId)).size} account(s). Review, then save as a snapshot.`)
  }

  function saveImportSnapshot() {
    const holdings = importRows.map((row, index) => normalizeImportedRow(row, importAccountOwner, String(row.__accountName ?? importName ?? 'Uploaded account'), String(row.__accountType ?? importAccountType), index)).filter((row): row is Holding => Boolean(row))
    if (!holdings.length) {
      setImportMessage('No usable holdings found. Confirm the file includes ticker and market value columns.')
      return
    }
    const snapshotDate = importSnapshotDate ? new Date(importSnapshotDate).toISOString() : new Date().toISOString()
    const snapshot: PortfolioSnapshot = { id: crypto.randomUUID(), name: `${importName || 'Uploaded portfolio'} snapshot`, date: snapshotDate, source: importName || 'Uploaded file', holdings }
    setSnapshots((items) => [...items, snapshot])
    setSelectedSnapshotId(snapshot.id)
    setImportRows([])
    setImportSnapshotDate('')
    setImportMessage(`Saved ${holdings.length} holdings as a dated snapshot.`)
  }

  return (
    <main className="app-root min-h-[100dvh] bg-app text-app">
      <AmbientBackdrop />
      <div className="app-shell mx-auto max-w-[1400px] px-4 py-5 md:px-6">
        <header className="mb-6 grid gap-5 border-b border-white/10 pb-5 lg:grid-cols-[1.5fr_1fr]">
          <div>
            <div className="mb-3 flex flex-wrap items-center gap-2 text-xs uppercase tracking-[0.18em] text-emerald-300/80">
              <span className="rounded-full border border-emerald-300/20 bg-emerald-300/10 px-3 py-1">Local-first</span>
              <span>Offline portfolio planning</span>
            </div>
            <h1 className="max-w-4xl text-3xl font-semibold tracking-tight text-zinc-50 md:text-5xl">AI Buildout Portfolio Recomp Cockpit</h1>
            <p className="mt-3 max-w-3xl text-sm leading-6 text-muted md:text-base">
              A private browser workspace for x-raying portfolio holdings, measuring AI infrastructure exposure, stress testing simple scenarios, and drafting simulated recomp candidates.
            </p>
          </div>
          <div className="rounded-2xl border border-app bg-panel p-4 shadow-[inset_0_1px_0_rgba(255,255,255,0.06)]">
            <div className="flex items-center justify-between gap-4">
              <span className="text-sm text-muted">Current snapshot</span>
              <select className="control" value={currentSnapshot.id} onChange={(event) => setSelectedSnapshotId(event.target.value)}>
                {snapshots.map((snapshot) => <option key={snapshot.id} value={snapshot.id}>{snapshot.name}</option>)}
              </select>
            </div>
            <div className="mt-3 grid gap-3 sm:grid-cols-2">
              <label className="field"><span>Account scope</span><select className="control" value={selectedAccountScope} onChange={(event) => setSelectedAccountScope(event.target.value)}>{accountOptions.map((account) => <option key={account.id} value={account.id}>{account.label}</option>)}</select></label>
              <label className="field"><span>Theme</span><select className="control" value={themePreference} onChange={(event) => setThemePreference(event.target.value as ThemePreference)}><option value="system">System</option><option value="dark">Dark</option><option value="light">Light</option></select></label>
            </div>
            <p className="mt-3 text-xs text-muted">As of {new Date(currentSnapshot.date).toLocaleDateString()}</p>
            <div className="mt-4 grid grid-cols-2 gap-3 text-sm">
              <Metric label="Value" value={dollarFmt.format(analytics.total)} />
              <Metric label="AI exposure" value={`${percentFmt.format(analytics.aiExposure)}%`} />
              <Metric label="Top 10" value={`${percentFmt.format(analytics.top10Weight)}%`} />
              <Metric label="Scenario" value={`${percentFmt.format(analytics.stressImpact)}% (${dollarFmt.format(scenarioImpactDollars(analytics))})`} />
            </div>
          </div>
        </header>

        <nav className="mb-6 flex gap-2 overflow-x-auto rounded-2xl border border-app bg-panel p-2">
          {['Overview', 'Import & Snapshots', 'Concentration & Holdings', 'AI Buildout', 'Recomp Sandbox'].map((tab) => (
            <button key={tab} className={`tab ${activeTab === tab ? 'tab-active' : ''}`} onClick={() => setActiveTab(tab)}>{tab}</button>
          ))}
        </nav>

        {activeTab === 'Overview' && <Overview analytics={analytics} template={template} templates={allTemplates} customTemplates={customTemplates} setCustomTemplates={setCustomTemplates} stress={stress} setTemplate={setSelectedTemplateId} setStress={setSelectedStressId} generateCandidates={generateCandidates} />}
        {activeTab === 'Import & Snapshots' && <ImportSnapshots snapshots={snapshots} current={currentSnapshot} rows={importRows} setRows={setImportRows} message={importMessage} snapshotDate={importSnapshotDate} setSnapshotDate={setImportSnapshotDate} accountOwner={importAccountOwner} accountType={importAccountType} setAccountOwner={setImportAccountOwner} setAccountType={setImportAccountType} parseUpload={parseUpload} saveImportSnapshot={saveImportSnapshot} setSnapshots={setSnapshots} setSelectedSnapshotId={setSelectedSnapshotId} generateCandidates={generateCandidates} />}
        {activeTab === 'Concentration & Holdings' && <Xray analytics={analytics} table={table} search={search} setSearch={setSearch} grouping={grouping} setGrouping={setGrouping} minDollar={minDollar} minWeight={minWeight} setMinDollar={setMinDollar} setMinWeight={setMinWeight} setSelectedHoldingId={setSelectedHoldingId} excludedHoldingIds={hiddenHoldingIds} excludedHoldings={hiddenHoldings} includeHolding={(id) => setHiddenHoldingIds((items) => items.filter((item) => item !== id))} includeAll={() => setHiddenHoldingIds([])} ledgerMode={ledgerMode} setLedgerMode={setLedgerMode} editedCount={editedCount} discardEdits={discardHoldingEdits} saveEditedSnapshot={saveEditedSnapshot} filterOptions={tableFilterOptions} filters={{ asset: assetFilter, sector: sectorFilter, aiBucket: aiBucketFilter, size: sizeFilter }} setFilters={{ asset: setAssetFilter, sector: setSectorFilter, aiBucket: setAiBucketFilter, size: setSizeFilter }} />}
        {activeTab === 'AI Buildout' && <AIBuildout analytics={analytics} current={currentSnapshot} selectedHolding={selectedHolding} setSelectedHoldingId={setSelectedHoldingId} updateHolding={updateHolding} />}
        {activeTab === 'Recomp Sandbox' && <Sandbox analytics={analytics} simulatedAnalytics={simulatedAnalytics} simulatedHoldings={simulated} holdings={scopedHoldings} candidates={recompCandidates} setCandidates={setRecompCandidates} manualActions={manualActions} setManualActions={setManualActions} positionPlans={positionPlans} setPositionPlans={setPositionPlans} clearCandidates={() => setRecompCandidates([])} generateCandidates={generateCandidates} decisionLog={decisionLog} />}

        <footer className="mt-8 border-t border-white/10 pt-4 text-xs leading-5 text-zinc-500">
          This app is for portfolio analysis and planning only. It does not provide financial advice, tax advice, or execute trades. All outputs are simulations based on uploaded data and simplified assumptions.
        </footer>
      </div>
    </main>
  )
}

function AmbientBackdrop() {
  return <div className="ambient-backdrop" aria-hidden="true" />
}

function Overview({ analytics, template, templates, customTemplates, setCustomTemplates, stress, setTemplate, setStress, generateCandidates }: { analytics: ReturnType<typeof analyze>; template: StrategyTemplate; templates: StrategyTemplate[]; customTemplates: StrategyTemplate[]; setCustomTemplates: React.Dispatch<React.SetStateAction<StrategyTemplate[]>>; stress: StressScenario; setTemplate: (id: string) => void; setStress: (id: string) => void; generateCandidates: () => void }) {
  function duplicateTemplate() {
    const copy = { ...template, id: `custom-${crypto.randomUUID()}`, name: `${template.name} custom`, targets: { ...template.targets }, pros: [...template.pros], cons: [...template.cons], warningLabel: 'Custom planning profile' }
    setCustomTemplates((items) => [...items, copy])
    setTemplate(copy.id)
  }
  return <section className="grid gap-5">
    <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-5">
      <MetricCard icon={<Database size={20} />} label="Total portfolio value" value={dollarFmt.format(analytics.total)} />
      <MetricCard icon={<StackSimple size={20} />} label="Holdings" value={String(analytics.holdings.length)} />
      <MetricCard icon={<Sparkle size={20} />} label="AI buildout exposure" value={`${percentFmt.format(analytics.aiExposure)}%`} />
      <MetricCard icon={<Funnel size={20} />} label={`Largest holding · ${analytics.topHoldings[0]?.ticker ?? 'N/A'}`} value={`${percentFmt.format(analytics.largestWeight)}%`} />
      <MetricCard icon={<Pulse size={20} />} label="Risk-budget estimate" value={analytics.riskScore.toFixed(2)} />
    </div>
    <p className="metric-row-note"><strong>Risk-budget estimate:</strong> 1.00 is broad-stock-like. Higher means more high-beta AI or single-name tilt; lower means more cash/bonds.</p>
    <div className="grid gap-5 xl:grid-cols-[1.2fr_0.8fr]">
      <Panel title={`Current snapshot read vs ${template.name}`} action={<button className="primary" onClick={generateCandidates}><Sparkle size={16} /> Generate Auto-Recomp Candidates</button>}>
        <p className="mb-3 text-sm leading-6 text-muted">This is the portfolio you selected in the snapshot/account controls. The template is only the comparison yardstick, not a replacement portfolio.</p>
        <p className="text-balance text-lg leading-8 text-app">{executiveSummary(analytics, template)}</p>
        <OverviewGuide />
        <div className="mt-5 rounded-xl border border-app bg-soft p-3">
          <label className="field"><span>Scenario sensitivity</span><select className="control" value={stress.id} onChange={(event) => setStress(event.target.value)}>{stressScenarios.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select></label>
          <p className="mt-3 text-sm leading-6 text-muted">{stressExplanation(stress)}</p>
          <StressShockList scenario={stress} />
        </div>
      </Panel>
      <WarningPanel warnings={analytics.warnings} />
    </div>
    <div className="grid gap-5 xl:grid-cols-2">
      <ChartPanel title="Asset class"><Donut data={analytics.assetClassData} /></ChartPanel>
      <ChartPanel title="Sector"><Donut data={analytics.sectorData.slice(0, 8)} /></ChartPanel>
      <ChartPanel title="AI buildout exposure by bucket"><BarList data={analytics.aiBucketData.slice(0, 10)} /></ChartPanel>
      <ChartPanel title="Top 10 holdings"><BarList data={holdingBarData(analytics.topHoldings.slice(0, 10), analytics.total)} /></ChartPanel>
    </div>
    <Panel title="Template Fit">
      <p className="mb-4 max-w-4xl text-sm leading-6 text-muted">Pick the portfolio posture first, then use the recomp worksheet to see the approximate actions implied by that posture. The charts here update immediately so you do not have to bounce between template cards and the allocation chart.</p>
      <div className="template-fit-grid">
        <div className="template-grid template-grid-compact">
          {templates.map((item) => <TemplateCard key={item.id} template={item} analytics={analytics} selected={item.id === template.id} onSelect={() => setTemplate(item.id)} />)}
        </div>
        <div className="template-fit-visuals">
          <div className="grid gap-4 md:grid-cols-2">
            <div>
              <h3 className="mini-heading">Current allocation</h3>
              <div className="h-64"><Donut data={analytics.assetClassData} /></div>
            </div>
            <div>
              <h3 className="mini-heading">Target allocation</h3>
              <div className="h-64"><AllocationDonut data={targetAllocationData(template)} target /></div>
            </div>
          </div>
          <div className="mt-4 h-80"><TargetBars analytics={analytics} template={template} /></div>
        </div>
      </div>
      <TemplateEditor template={template} isCustom={template.id.startsWith('custom-')} customTemplates={customTemplates} setCustomTemplates={setCustomTemplates} setTemplate={setTemplate} duplicateTemplate={duplicateTemplate} />
    </Panel>
  </section>
}

function StressShockList({ scenario }: { scenario: StressScenario }) {
  return <div className="stress-list">{Object.entries(scenario.shocks).map(([name, value]) => <span key={name}>{name}: <strong>{value && value > 0 ? '+' : ''}{value}%</strong></span>)}</div>
}

function OverviewGuide() {
  const items = [
    ['Update data', 'Import & Snapshots', 'Upload files, set dates, rename snapshots, compare what changed.'],
    ['Audit holdings', 'Concentration & Holdings', 'Sort, filter, exclude noise, and edit sandbox fields.'],
    ['Check AI thesis', 'AI Buildout', 'See whether the AI sleeve is balanced or too crowded.'],
    ['Test a path', 'Recomp Sandbox', 'Convert a chosen template into simulated action math.'],
  ]
  return <div className="overview-guide">
    {items.map(([goal, title, body]) => <div key={title}><em>{goal}</em><strong>{title}</strong><span>{body}</span></div>)}
  </div>
}

function TemplateEditor({ template, isCustom, customTemplates, setCustomTemplates, setTemplate, duplicateTemplate }: { template: StrategyTemplate; isCustom: boolean; customTemplates: StrategyTemplate[]; setCustomTemplates: React.Dispatch<React.SetStateAction<StrategyTemplate[]>>; setTemplate: (id: string) => void; duplicateTemplate: () => void }) {
  function updateCustom(next: StrategyTemplate) {
    setCustomTemplates((items) => items.map((item) => item.id === next.id ? next : item))
  }
  return <div className="template-editor">
    <div>
      <h3>{isCustom ? 'Edit custom template' : 'Make your own template'}</h3>
      <p>{isCustom ? 'Adjust targets and guardrails. Changes are saved locally in this browser.' : 'Default profiles are locked. Duplicate one to create a local editable template.'}</p>
    </div>
    {!isCustom ? <button className="ghost" onClick={duplicateTemplate}>Duplicate selected template</button> : <div className="grid gap-3">
      <label className="field"><span>Template name</span><input className="control" value={template.name} onChange={(event) => updateCustom({ ...template, name: event.target.value })} /></label>
      <div className="custom-target-grid">
        {assetClasses.map((asset) => <label key={asset} className="field"><span>{assetShortLabel(asset)} target %</span><input className="control" type="number" min={0} max={100} value={template.targets[asset]} onChange={(event) => updateCustom({ ...template, targets: { ...template.targets, [asset]: Number(event.target.value) } })} /></label>)}
      </div>
      <div className="custom-target-grid">
        <label className="field"><span>Max single %</span><input className="control" type="number" value={template.maxSingle} onChange={(event) => updateCustom({ ...template, maxSingle: Number(event.target.value) })} /></label>
        <label className="field"><span>Max top 10 %</span><input className="control" type="number" value={template.maxTop10} onChange={(event) => updateCustom({ ...template, maxTop10: Number(event.target.value) })} /></label>
        <label className="field"><span>Max AI bucket %</span><input className="control" type="number" value={template.maxBucket} onChange={(event) => updateCustom({ ...template, maxBucket: Number(event.target.value) })} /></label>
      </div>
      <button className="ghost justify-self-start" onClick={() => { setCustomTemplates(customTemplates.filter((item) => item.id !== template.id)); setTemplate(strategyTemplates[0].id) }}>Delete custom template</button>
    </div>}
  </div>
}

function TemplateCard({ template, analytics, selected, onSelect }: { template: StrategyTemplate; analytics: ReturnType<typeof analyze>; selected: boolean; onSelect: () => void }) {
  const drift = assetClasses.reduce((sum, asset) => sum + Math.abs(weight(analytics.assetTotals[asset] ?? 0, analytics.total) - template.targets[asset]), 0)
  return <button className={`template-card ${selected ? 'template-card-selected' : ''}`} onClick={onSelect}>
    <div className="flex items-start justify-between gap-3">
      <div>
        <h3>{template.name}</h3>
        <p>{template.purpose}</p>
      </div>
      <span className="template-pill">{template.targets['AI buildout sleeve']}% AI</span>
    </div>
    <div className="allocation-strip" aria-label={`${template.name} allocation mix`}>
      {assetClasses.map((asset) => <span key={asset} title={`${asset}: ${template.targets[asset]}%`} style={{ width: `${template.targets[asset]}%`, backgroundColor: assetClassColors[asset] }} />)}
    </div>
    <div className="allocation-legend">
      {assetClasses.map((asset) => <span key={asset}><i style={{ backgroundColor: assetClassColors[asset] }} />{assetShortLabel(asset)} {template.targets[asset]}%</span>)}
    </div>
    <div className="template-stats">
      <span>Drawdown <strong>{template.drawdownRange}</strong></span>
      <span>Max single <strong>{template.maxSingle}%</strong></span>
      <span>Top 10 max <strong>{template.maxTop10}%</strong></span>
      <span>Current drift <strong>{percentFmt.format(drift)} pts</strong></span>
    </div>
    <div className="template-pros-cons">
      <div><strong>Pros</strong>{template.pros.slice(0, 2).map((item) => <span key={item}>{item}</span>)}</div>
      <div><strong>Cons</strong>{template.cons.slice(0, 2).map((item) => <span key={item}>{item}</span>)}</div>
    </div>
  </button>
}

function ImportSnapshots(props: { snapshots: PortfolioSnapshot[]; current: PortfolioSnapshot; rows: Record<string, unknown>[]; setRows: React.Dispatch<React.SetStateAction<Record<string, unknown>[]>>; message: string; snapshotDate: string; setSnapshotDate: (value: string) => void; accountOwner: string; accountType: string; setAccountOwner: (v: string) => void; setAccountType: (v: string) => void; parseUpload: (files: FileList | File) => void; saveImportSnapshot: () => void; setSnapshots: React.Dispatch<React.SetStateAction<PortfolioSnapshot[]>>; setSelectedSnapshotId: (id: string) => void; generateCandidates: () => void }) {
  const latest = props.snapshots.at(-1)
  const previous = props.snapshots.at(-2)
  const comparison = latest && previous ? compareSnapshots(previous, latest) : []
  const stagedAccounts = summarizeStagedAccounts(props.rows)
  const [renamingId, setRenamingId] = useState<string | null>(null)
  const [renameDraft, setRenameDraft] = useState('')
  function beginRename(snapshot: PortfolioSnapshot) {
    setRenamingId(snapshot.id)
    setRenameDraft(snapshot.name)
  }
  function saveRename() {
    const name = renameDraft.trim()
    if (!renamingId || !name) return
    props.setSnapshots((items) => items.map((item) => item.id === renamingId ? { ...item, name } : item))
    setRenamingId(null)
    setRenameDraft('')
  }
  return <section className="grid gap-5 xl:grid-cols-[0.9fr_1.1fr]">
    <Panel title="Upload holdings" action={<button className="primary" onClick={props.generateCandidates}><Sparkle size={16} /> Generate Auto-Recomp Candidates</button>}>
      <label className="upload">
        <FileArrowUp size={28} />
        <span>Drop in one or more CSV/XLSX brokerage position files</span>
        <input type="file" multiple accept=".csv,.xlsx,.xls" onChange={(event) => event.target.files && props.parseUpload(event.target.files)} />
      </label>
      {props.message && <p className="mt-3 rounded-xl border border-emerald-300/20 bg-emerald-300/10 p-3 text-sm text-emerald-100">{props.message}</p>}
      {props.rows.length > 0 && <div className="mt-4 grid gap-3 md:grid-cols-2">
        <label className="field"><span>Account owner</span><input className="control" value={props.accountOwner} onChange={(event) => props.setAccountOwner(event.target.value)} /></label>
        <label className="field"><span>Account type</span><input className="control" value={props.accountType} onChange={(event) => props.setAccountType(event.target.value)} /></label>
        <label className="field md:col-span-2"><span>Snapshot as-of date</span><input className="control" type="date" value={props.snapshotDate} onChange={(event) => props.setSnapshotDate(event.target.value)} /></label>
        <div className="staged-account-list md:col-span-2">
          {stagedAccounts.map((account) => <div key={account.id} className="staged-account">
            <div><strong>{account.name}</strong><span>{account.type} · {account.rows} rows · {dollarFmt.format(account.value)}</span></div>
            <button className="ghost" onClick={() => props.setRows((rows) => rows.filter((row) => row.__accountId !== account.id))}>Remove account</button>
          </div>)}
        </div>
        <div className="flex flex-wrap gap-2 md:col-span-2">
          <button className="primary" onClick={props.saveImportSnapshot}><UploadSimple size={16} /> Save as dated snapshot</button>
          <button className="ghost" onClick={() => props.setRows([])}>Clear staged files</button>
        </div>
        <PreviewRows rows={props.rows} />
      </div>}
    </Panel>
    <Panel title="Snapshots and changes">
      <div className="space-y-3">
        {props.snapshots.map((snapshot) => <div key={snapshot.id} className="row">
          <div className="min-w-0 flex-1">
            {renamingId === snapshot.id ? <div className="rename-inline">
              <input className="control" value={renameDraft} autoFocus onChange={(event) => setRenameDraft(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') saveRename(); if (event.key === 'Escape') setRenamingId(null) }} />
              <button className="primary" onClick={saveRename}>Save</button>
              <button className="ghost" onClick={() => setRenamingId(null)}>Cancel</button>
            </div> : <p className="font-medium text-zinc-100">{snapshot.name}</p>}
            <p className="text-xs text-zinc-500">{new Date(snapshot.date).toLocaleString()} · {snapshot.holdings.length} holdings · {snapshot.source}</p>
          </div>
          <div className="flex gap-2">
            <button className="ghost" onClick={() => props.setSelectedSnapshotId(snapshot.id)}>View</button>
            <button className="ghost" onClick={() => beginRename(snapshot)}><PencilSimple size={16} /> Rename</button>
            {props.snapshots.length > 1 && <button className="ghost" onClick={() => props.setSnapshots((items) => items.filter((item) => item.id !== snapshot.id))}><Trash size={16} /> Delete</button>}
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

function Xray({ analytics, table, search, setSearch, grouping, setGrouping, minDollar, minWeight, setMinDollar, setMinWeight, setSelectedHoldingId, excludedHoldingIds, excludedHoldings, includeHolding, includeAll, ledgerMode, setLedgerMode, editedCount, discardEdits, saveEditedSnapshot, filterOptions, filters, setFilters }: { analytics: ReturnType<typeof analyze>; table: ReturnType<typeof useReactTable<Holding>>; search: string; setSearch: (v: string) => void; grouping: GroupingState; setGrouping: (v: GroupingState) => void; minDollar: number; minWeight: number; setMinDollar: (v: number) => void; setMinWeight: (v: number) => void; setSelectedHoldingId: (id: string) => void; excludedHoldingIds: string[]; excludedHoldings: Holding[]; includeHolding: (id: string) => void; includeAll: () => void; ledgerMode: LedgerMode; setLedgerMode: (mode: LedgerMode) => void; editedCount: number; discardEdits: () => void; saveEditedSnapshot: () => void; filterOptions: ReturnType<typeof tableFilters>; filters: { asset: string; sector: string; aiBucket: string; size: string }; setFilters: { asset: (value: string) => void; sector: (value: string) => void; aiBucket: (value: string) => void; size: (value: string) => void } }) {
  const [treemapOpen, setTreemapOpen] = useState(false)
  const nuisanceCount = analytics.holdings.filter((holding) => isNuisanceHolding(holding, analytics.total, minDollar, minWeight)).length
  return <section className="grid gap-5">
    <ChartPanel title="Top holdings ranked" tall><ChartLegend items={holdingColorLegend} /><BarList data={holdingBarData(analytics.topHoldings.slice(0, 15), analytics.total)} /></ChartPanel>
    <ChartPanel title="Holdings treemap" action={<button className="ghost" onClick={() => setTreemapOpen(true)}>Enlarge</button>} tall><ChartLegend items={holdingColorLegend} /><HoldingsTreemap holdings={analytics.topHoldings} /></ChartPanel>
    {treemapOpen && <TreemapDialog holdings={analytics.topHoldings} onClose={() => setTreemapOpen(false)} />}
    <ConcentrationCutoffs analytics={analytics} />
    <div className="threshold-strip">
      <div>
        <strong>Nuisance threshold</strong>
        <span>Used for tiny-position warnings and the table size filter. It does not hide anything unless you choose a size filter below.</span>
      </div>
      <label className="field"><span>Below dollars</span><input type="number" className="control" value={minDollar} onChange={(event) => setMinDollar(Number(event.target.value))} /></label>
      <label className="field"><span>Below portfolio %</span><input type="number" className="control" value={minWeight} onChange={(event) => setMinWeight(Number(event.target.value))} /></label>
      <span className="threshold-count">{nuisanceCount} below threshold</span>
    </div>
    <Panel title="Holdings table" action={<button className="ghost" onClick={() => exportCsv('current-view-holdings.csv', analytics.holdings)}><DownloadSimple size={16} /> Export current view</button>}>
      <div className="ledger-toolbar">
        <div>
          <div className="segmented">
            <button className={ledgerMode === 'uploaded' ? 'selected' : ''} onClick={() => setLedgerMode('uploaded')}>Uploaded snapshot</button>
            <button className={ledgerMode === 'sandbox' ? 'selected' : ''} onClick={() => setLedgerMode('sandbox')}>Edit sandbox</button>
          </div>
          <p>Sandbox edits update every chart and warning without changing the uploaded snapshot. Save them as a new edited snapshot when you want a permanent copy.</p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <span className="edited-count">{editedCount} edited</span>
          <button className="ghost" onClick={discardEdits}>Discard edits</button>
          <button className="primary" onClick={saveEditedSnapshot} disabled={!editedCount}>Save as edited snapshot</button>
        </div>
      </div>
      <div className="table-filter-grid">
        <label className="search"><MagnifyingGlass size={16} /><input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Search holdings" /></label>
        <label className="field"><span>Asset class</span><select className="control" value={filters.asset} onChange={(event) => setFilters.asset(event.target.value)}><option>All</option>{filterOptions.assets.map((asset) => <option key={asset}>{asset}</option>)}</select></label>
        <label className="field"><span>Sector</span><select className="control" value={filters.sector} onChange={(event) => setFilters.sector(event.target.value)}><option>All</option>{filterOptions.sectors.map((sector) => <option key={sector}>{sector}</option>)}</select></label>
        <label className="field"><span>AI bucket</span><select className="control" value={filters.aiBucket} onChange={(event) => setFilters.aiBucket(event.target.value)}><option>All</option>{filterOptions.aiBuckets.map((bucket) => <option key={bucket}>{bucket}</option>)}</select></label>
        <label className="field"><span>Size</span><select className="control" value={filters.size} onChange={(event) => setFilters.size(event.target.value)}><option>All sizes</option><option>Below nuisance threshold</option><option>Above nuisance threshold</option></select></label>
        <label className="field"><span>Group rows</span><select className="control" value={grouping[0] ?? ''} onChange={(event) => setGrouping(event.target.value ? [event.target.value] : [])}><option value="">No grouping</option><option value="accountName">Group account</option><option value="assetClass">Group asset class</option><option value="sector">Group sector</option></select></label>
      </div>
      <div className="table-shell">
        <table className="data-table">
          <thead>{table.getHeaderGroups().map((group) => <tr key={group.id}>{group.headers.map((header) => <th key={header.id} className={header.column.getCanSort() ? 'sortable-th' : ''} onClick={header.column.getToggleSortingHandler()}><span>{flexRender(header.column.columnDef.header, header.getContext())}<SortIndicator value={header.column.getIsSorted()} /></span></th>)}</tr>)}</thead>
          <tbody>{table.getRowModel().rows.map((row) => <tr key={row.id} className={holdingExcludedInView(row.original, excludedHoldingIds) ? 'excluded-row' : ''} onClick={() => setSelectedHoldingId(row.original.id)}>{row.getVisibleCells().map((cell) => <td key={cell.id}>{flexRender(cell.column.columnDef.cell, cell.getContext())}</td>)}</tr>)}</tbody>
        </table>
      </div>
      {excludedHoldings.length > 0 && <div className="mt-4 rounded-xl border border-app bg-soft p-3">
        <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
          <h3 className="section-label !m-0">Excluded from analytics</h3>
          <button className="ghost" onClick={includeAll}>Include all</button>
        </div>
        <div className="flex flex-wrap gap-2">{excludedHoldings.map((holding) => <button key={holding.id} className="ghost" onClick={() => includeHolding(holding.id)}>{holding.ticker} · Include</button>)}</div>
      </div>}
    </Panel>
  </section>
}

function AIBuildout({ analytics, current, selectedHolding, setSelectedHoldingId, updateHolding }: { analytics: ReturnType<typeof analyze>; current: PortfolioSnapshot; selectedHolding?: Holding; setSelectedHoldingId: (id: string | null) => void; updateHolding: (holding: Holding) => void }) {
  return <section className="grid gap-5 xl:grid-cols-[1fr_380px]">
    <div className="grid gap-5">
      <ChartPanel title="AI exposure by bucket" tall><BarList data={analytics.aiBucketData} /></ChartPanel>
      <ChartPanel title="Direct vs indirect AI exposure"><Donut data={analytics.directnessData} /></ChartPanel>
      <Panel title="Top AI exposure contributors">
        <div className="space-y-2">
          {analytics.holdings.filter((h) => h.ai.score > 0).sort((a, b) => b.marketValue * b.ai.score - a.marketValue * a.ai.score).slice(0, 12).map((holding) => <button key={holding.id} className="row w-full text-left" onClick={() => setSelectedHoldingId(holding.id)}><span><strong>{holding.ticker}</strong> <span className="text-zinc-500">{holding.securityName}</span></span><span className="font-mono text-emerald-300">{holding.ai.score.toFixed(1)}</span></button>)}
        </div>
      </Panel>
    </div>
    <div className="grid gap-5">
      <AIScoreGuide />
      <Panel title="Editable classification" action={selectedHolding && <button className="ghost" onClick={() => setSelectedHoldingId(null)}>Close</button>}>
        {!selectedHolding ? <div className="empty"><PencilSimple size={28} /><p>Select a holding from the contributor list or holdings table to edit its AI classification.</p></div> : <ClassificationEditor holding={selectedHolding} current={current} updateHolding={updateHolding} />}
      </Panel>
    </div>
  </section>
}
function AIScoreGuide() {
  const rows = [
    ['0', 'No meaningful AI buildout exposure'],
    ['1', 'Broad, passive, or incidental exposure'],
    ['2', 'Tangential beneficiary'],
    ['3', 'Meaningful indirect supply-chain exposure'],
    ['4', 'Direct core AI infrastructure exposure'],
    ['5', 'Pure-play or mission-critical AI infrastructure'],
  ]
  return <Panel title="AI score guide">
    <p className="mb-3 text-sm leading-6 text-muted">The score is a local heuristic, not a market model. It mainly helps weight AI exposure across holdings; dollar size, bucket, directness, and confidence matter more than the score by itself.</p>
    <div className="score-guide">{rows.map(([score, label]) => <div key={score}><strong>{score}</strong><span>{label}</span></div>)}</div>
  </Panel>
}

function Sandbox({ analytics, simulatedAnalytics, simulatedHoldings, holdings, candidates, setCandidates, manualActions, setManualActions, positionPlans, setPositionPlans, clearCandidates, generateCandidates, decisionLog }: { analytics: ReturnType<typeof analyze>; simulatedAnalytics: ReturnType<typeof analyze>; simulatedHoldings: Holding[]; holdings: Holding[]; candidates: RecompCandidate[]; setCandidates: React.Dispatch<React.SetStateAction<RecompCandidate[]>>; manualActions: ManualSandboxAction[]; setManualActions: React.Dispatch<React.SetStateAction<ManualSandboxAction[]>>; positionPlans: Record<string, PositionPlan>; setPositionPlans: React.Dispatch<React.SetStateAction<Record<string, PositionPlan>>>; clearCandidates: () => void; generateCandidates: () => void; decisionLog: DecisionLogEntry[] }) {
  const [mode, setMode] = useState<'template' | 'manual' | 'plans'>('template')
  const [ticker, setTicker] = useState(holdings[0]?.ticker ?? '')
  const [amount, setAmount] = useState(5000)
  const allActions = [...candidates, ...manualActions]
  const netChange = simulatedAnalytics.total - analytics.total
  const currentScenario = scenarioImpactDollars(analytics)
  const simulatedScenario = scenarioImpactDollars(simulatedAnalytics)
  function updateAction(next: RecompCandidate) {
    setCandidates((items) => items.map((item) => item.id === next.id ? next : item))
    setManualActions((items) => items.map((item) => item.id === next.id ? { ...item, ...next } : item))
  }
  function addManual(actionType: ActionType) {
    const holding = holdings.find((item) => item.ticker === ticker)
    const price = priceForHolding(holding)
    setManualActions((items) => [...items, { id: crypto.randomUUID(), actionType, tickerOrBucket: ticker, dollarAmount: amount, estimatedShares: price ? amount / price : undefined, priceUsed: price, basis: 'manual', beforeWeight: holding ? weight(holding.marketValue, analytics.total) : 0, afterWeight: holding ? weight(holding.marketValue + (actionType.includes('Buy') ? amount : -amount), analytics.total) : 0, reason: 'Manual worksheet action.', riskImpact: 'User-defined impact.', alignmentImpact: 'Included in simulated before/after view.', stressImpact: 'Scenario sensitivity recomputes in the simulated portfolio.', note: '', }])
  }
  return <section className="grid gap-5">
    <Panel title="Recomp worksheet">
      <div className="ledger-toolbar">
        <div>
          <div className="segmented" aria-label="Sandbox mode">
            <button className={mode === 'template' ? 'selected' : ''} onClick={() => setMode('template')}>Template Gap</button>
            <button className={mode === 'manual' ? 'selected' : ''} onClick={() => setMode('manual')}>Manual Worksheet</button>
            <button className={mode === 'plans' ? 'selected' : ''} onClick={() => setMode('plans')}>Position Plans</button>
          </div>
          <p>{mode === 'plans' ? 'Holding notes and review triggers stay local. They flag review moments; they do not create orders.' : 'Simulated actions translate the selected template into editable dollars and share math. They are planning estimates, not brokerage orders.'}</p>
        </div>
        {mode === 'template' && <div className="flex gap-2"><button className="primary" onClick={generateCandidates}><ArrowClockwise size={16} /> Generate</button><button className="ghost" onClick={clearCandidates}>Clear</button></div>}
      </div>
      {mode === 'manual' && <div className="grid gap-3 md:grid-cols-[1fr_1fr_auto]">
        <label className="field"><span>Ticker</span><input className="control" list="holding-tickers" value={ticker} onChange={(event) => setTicker(event.target.value.toUpperCase())} /><datalist id="holding-tickers">{holdings.map((holding) => <option key={holding.id} value={holding.ticker} />)}</datalist></label>
        <label className="field"><span>Dollar amount</span><input className="control" type="number" value={amount} onChange={(event) => setAmount(Number(event.target.value))} /></label>
        <div className="flex items-end gap-2"><button className="ghost" onClick={() => addManual('Buy dollar amount')}>Buy</button><button className="ghost" onClick={() => addManual('Sell dollar amount')}>Sell</button></div>
      </div>}
      {mode === 'template' && <p className="text-sm leading-6 text-muted">Auto-Recomp trims oversized concentration, reduces overexposed sleeves, and creates bucket-level allocate-here candidates when the app should not invent a new security.</p>}
      {mode === 'plans' && <PositionPlans holdings={holdings} analytics={analytics} plans={positionPlans} setPlans={setPositionPlans} />}
    </Panel>
    {mode !== 'plans' && <>
    <Panel title="Simulation output">
      <p className="mb-4 text-sm leading-6 text-muted">Recomp simulations are reallocations, so this view does not call the result profit or loss. Scenario sensitivity is a secondary what-if: each holding gets the scenario shock for its AI bucket first, then asset class fallback, and the weighted impacts are summed.</p>
      <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-3">
        <MetricCard icon={<Database size={20} />} label="Simulated portfolio value" value={dollarFmt.format(simulatedAnalytics.total)} hint={`Current: ${dollarFmt.format(analytics.total)}`} />
        <MetricCard icon={<ChartBar size={20} />} label="Simulated action net change" value={dollarFmt.format(netChange)} hint={`${percentFmt.format(weight(netChange, analytics.total))}% vs current`} />
        <MetricCard icon={<ChartBar size={20} />} label="Cash after simulation" value={dollarFmt.format(simulatedAnalytics.assetTotals.Cash ?? 0)} hint={`Before: ${dollarFmt.format(analytics.assetTotals.Cash ?? 0)}`} />
        <MetricCard icon={<Sparkle size={20} />} label="AI exposure after simulation" value={`${percentFmt.format(simulatedAnalytics.aiExposure)}%`} hint={`Before: ${percentFmt.format(analytics.aiExposure)}%`} />
        <MetricCard icon={<Funnel size={20} />} label="Top 10 after simulation" value={`${percentFmt.format(simulatedAnalytics.top10Weight)}%`} hint={`Before: ${percentFmt.format(analytics.top10Weight)}%`} />
        <MetricCard icon={<Scales size={20} />} label="Stress scenario impact" value={`${percentFmt.format(simulatedAnalytics.stressImpact)}% (${dollarFmt.format(simulatedScenario)})`} hint={`Current: ${percentFmt.format(analytics.stressImpact)}% (${dollarFmt.format(currentScenario)})`} />
      </div>
    </Panel>
    <Panel title="Simulated trade candidate table" action={<button className="ghost" onClick={() => exportCsv('simulated-recomp-candidates.csv', allActions)}><DownloadSimple size={16} /> Export CSV</button>}>
      <p className="mb-3 text-sm leading-6 text-muted">Estimated shares use CSV price when present, otherwise market value divided by shares. Editing shares updates dollars; editing dollars updates shares. These are not order tickets or advice about timing or limit price.</p>
      <ActionTable actions={allActions} holdings={holdings} onChange={updateAction} />
    </Panel>
    <Panel title="Before / after holdings breakdown">
      <HoldingDeltaTable before={holdings} after={simulatedHoldings} />
    </Panel>
    <div className="grid gap-5 xl:grid-cols-2">
      <ChartPanel title="Current vs simulated top holdings"><BeforeAfter before={holdingBarData(analytics.topHoldings.slice(0, 10), analytics.total)} after={holdingBarData(simulatedAnalytics.topHoldings.slice(0, 10), simulatedAnalytics.total)} /></ChartPanel>
      <ChartPanel title="Current vs simulated allocation"><BeforeAfter before={analytics.assetClassData} after={simulatedAnalytics.assetClassData} /></ChartPanel>
      <ChartPanel title="Current vs simulated AI exposure"><BeforeAfter before={compactChartData(analytics.aiBucketData, 8)} after={compactChartData(simulatedAnalytics.aiBucketData, 8)} /></ChartPanel>
    </div>
    </>}
    <Panel title="Decision log" action={<button className="ghost" onClick={() => exportJson('decision-log.json', decisionLog)}><DownloadSimple size={16} /> Export JSON</button>}>
      <div className="space-y-2">{decisionLog.map((log) => <div key={log.id} className="row"><span>{new Date(log.date).toLocaleString()} · {log.templateName}</span><span className="font-mono text-zinc-400">{log.actions.length} actions</span></div>)}</div>
    </Panel>
  </section>
}

function Metric({ label, value }: { label: string; value: string }) {
  return <div><p className="text-xs uppercase tracking-[0.16em] text-zinc-500">{label}</p><p className="mt-1 font-mono text-lg text-zinc-100">{value}</p></div>
}
function MetricCard({ icon, label, value, hint }: { icon: React.ReactNode; label: string; value: string; hint?: string }) {
  return <div className="metric-card"><div className="metric-icon">{icon}</div><div><p>{label}</p><strong>{value}</strong>{hint && <span className="metric-hint">{hint}</span>}</div></div>
}
function Panel({ title, action, children }: { title: string; action?: React.ReactNode; children: React.ReactNode }) {
  return <section className="panel"><div className="panel-header"><h2>{title}</h2>{action}</div>{children}</section>
}
function NumberCell({ value, onChange, min, max, step }: { value: number; onChange: (value: number) => void; min?: number; max?: number; step?: number }) {
  return <input className="table-input" type="number" value={Number.isFinite(value) ? value : 0} min={min} max={max} step={step} onClick={(event) => event.stopPropagation()} onChange={(event) => onChange(Number(event.target.value))} />
}
function TextCell({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  return <input className="table-input table-input-wide" value={value} onClick={(event) => event.stopPropagation()} onChange={(event) => onChange(event.target.value)} />
}
function SelectCell({ value, options, onChange }: { value: string; options: string[]; onChange: (value: string) => void }) {
  return <select className="table-input table-input-wide" value={value} onClick={(event) => event.stopPropagation()} onChange={(event) => onChange(event.target.value)}>{options.map((option) => <option key={option}>{option}</option>)}</select>
}
function SortIndicator({ value }: { value: false | 'asc' | 'desc' }) {
  return <span className="sort-indicator">{value === 'asc' ? 'Asc' : value === 'desc' ? 'Desc' : 'Sort'}</span>
}
function ConcentrationCutoffs({ analytics }: { analytics: ReturnType<typeof analyze> }) {
  const cutoffs = concentrationCutoffs(analytics.topHoldings, analytics.total)
  return <Panel title="Concentration cutoffs">
    <div className="cutoff-grid">
      <CutoffCard label="50% checkpoint" cutoff={cutoffs[50]} />
      <CutoffCard label="80% checkpoint" cutoff={cutoffs[80]} />
      <div className="cutoff-card">
        <p>Long tail</p>
        <strong>{cutoffs[80] ? analytics.holdings.length - cutoffs[80].rank : analytics.holdings.length} holdings</strong>
        <span>{cutoffs[80] ? `${percentFmt.format(Math.max(0, 100 - cutoffs[80].cumulative))}% after 80% checkpoint` : 'Portfolio has no holdings yet'}</span>
      </div>
    </div>
  </Panel>
}
function CutoffCard({ label, cutoff }: { label: string; cutoff?: { holding: Holding; rank: number; before: number; cumulative: number } }) {
  return <div className="cutoff-card">
    <p>{label}</p>
    {cutoff ? <>
      <strong>{cutoff.rank} holdings</strong>
      <span>Reached at {cutoff.holding.ticker}, crossing {percentFmt.format(cutoff.before)}% to {percentFmt.format(cutoff.cumulative)}%.</span>
    </> : <>
      <strong>-</strong>
      <span>No holdings to calculate.</span>
    </>}
  </div>
}
function ChartPanel({ title, action, tall = false, children }: { title: string; action?: React.ReactNode; tall?: boolean; children: React.ReactNode }) {
  const treemap = title.toLowerCase().includes('treemap')
  return <Panel title={title} action={action}><div className={`chart-frame ${tall ? 'chart-frame-tall' : ''} ${treemap ? 'chart-frame-treemap' : ''}`}>{children}</div></Panel>
}
function ChartLegend({ items }: { items: { label: string; color: string }[] }) {
  return <div className="chart-legend">{items.map((item) => <span key={item.label}><i style={{ backgroundColor: item.color }} />{item.label}</span>)}</div>
}
function WarningPanel({ warnings }: { warnings: AppWarning[] }) {
  return <Panel title="Warnings"><div className="space-y-3">{warnings.map((warning) => <div key={warning.id} className={`warning warning-${warning.severity}`}><ShieldWarning size={18} /><div><p>{warning.title}</p><span>{warning.detail}</span></div></div>)}</div></Panel>
}
type BarDatum = { name: string; value: number; label?: string; color?: string }
function Donut({ data }: { data: { name: string; value: number }[] }) {
  return <ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={data} dataKey="value" nameKey="name" innerRadius={58} outerRadius={96} paddingAngle={2}>{data.map((item, i) => <Cell key={i} fill={chartColor(item.name, i)} stroke="var(--app-bg)" />)}</Pie><Tooltip formatter={(v) => `${percentFmt.format(Number(v))}%`} contentStyle={chartTooltip} /><Legend wrapperStyle={{ color: 'var(--muted-text)', fontSize: 12 }} /></PieChart></ResponsiveContainer>
}
function AllocationDonut({ data, target = false }: { data: { name: string; value: number }[]; target?: boolean }) {
  return <ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={data} dataKey="value" nameKey="name" innerRadius={58} outerRadius={96} paddingAngle={2}>{data.map((item, i) => <Cell key={i} fill={chartColor(item.name, i)} fillOpacity={target ? 0.52 : 1} stroke="var(--app-bg)" />)}</Pie><Tooltip formatter={(v) => `${percentFmt.format(Number(v))}%`} contentStyle={chartTooltip} /><Legend wrapperStyle={{ color: 'var(--muted-text)', fontSize: 12 }} /></PieChart></ResponsiveContainer>
}
function BarList({ data }: { data: BarDatum[] }) {
  const hasLabels = data.some((item) => item.label)
  return <ResponsiveContainer width="100%" height="100%"><BarChart data={data} layout="vertical" margin={{ top: 4, right: hasLabels ? 120 : 24, bottom: 24, left: 8 }}><CartesianGrid strokeDasharray="3 3" stroke={chartGrid} /><XAxis type="number" stroke={chartAxis} tick={chartTick} tickLine={false} axisLine={{ stroke: chartAxis }} /><YAxis type="category" dataKey="name" width={132} stroke={chartAxis} tick={chartTick} tickFormatter={shortChartLabel} interval={0} tickLine={false} /><Tooltip formatter={(v) => `${percentFmt.format(Number(v))}%`} labelFormatter={(label) => String(label)} contentStyle={chartTooltip} /><Bar dataKey="value" radius={[0, 6, 6, 0]}>{data.map((item, i) => <Cell key={item.name} fill={item.color ?? chartColor(item.name, i)} />)}{hasLabels && <LabelList dataKey="label" position="right" className="bar-value-label" />}</Bar></BarChart></ResponsiveContainer>
}
function HoldingsTreemap({ holdings }: { holdings: Holding[] }) {
  const total = holdings.reduce((sum, holding) => sum + holding.marketValue, 0)
  const data = holdings.map((holding) => ({ name: holding.ticker, size: holding.marketValue, weight: weight(holding.marketValue, total), color: holdingColor(holding) }))
  return <div className="treemap-stage"><ResponsiveContainer width="100%" height="100%"><Treemap data={data} dataKey="size" aspectRatio={4 / 3} stroke="var(--app-bg)" content={<TreemapCell />} /></ResponsiveContainer></div>
}
function TreemapDialog({ holdings, onClose }: { holdings: Holding[]; onClose: () => void }) {
  return <div className="modal-backdrop" role="presentation" onMouseDown={onClose}>
    <section className="treemap-modal" role="dialog" aria-modal="true" aria-label="Enlarged holdings treemap" onMouseDown={(event) => event.stopPropagation()}>
      <div className="panel-header">
        <h2>Holdings treemap</h2>
        <button className="ghost" onClick={onClose}>Close</button>
      </div>
      <ChartLegend items={holdingColorLegend} />
      <div className="treemap-modal-frame"><HoldingsTreemap holdings={holdings} /></div>
    </section>
  </div>
}
function TreemapCell(props: { x?: number; y?: number; width?: number; height?: number; name?: string; color?: string; size?: number; weight?: number }) {
  const { x = 0, y = 0, width = 0, height = 0, name = '', color = semanticColors.other, size = 0, weight: pct = 0 } = props
  if (width <= 0 || height <= 0) return null
  const label = `${name} · ${percentFmt.format(pct)}% · ${compactDollar(size)}`
  return <g className="treemap-cell">
    <title>{label}</title>
    <rect x={x + 1} y={y + 1} width={Math.max(0, width - 2)} height={Math.max(0, height - 2)} fill={color} rx={4} ry={4} />
    {width > 54 && height > 26 && <text x={x + 8} y={y + 18} fill="#ffffff" fontSize={11} fontWeight={500}>{name}</text>}
  </g>
}
function TargetBars({ analytics, template }: { analytics: ReturnType<typeof analyze>; template: StrategyTemplate }) {
  const data = assetClasses.map((asset) => ({ name: asset, current: weight(analytics.assetTotals[asset] ?? 0, analytics.total), target: template.targets[asset] }))
  return <ResponsiveContainer width="100%" height="100%"><BarChart data={data} layout="vertical" margin={{ top: 4, right: 24, bottom: 24, left: 8 }}><CartesianGrid strokeDasharray="3 3" stroke={chartGrid} /><XAxis type="number" stroke={chartAxis} tick={chartTick} tickLine={false} /><YAxis dataKey="name" type="category" width={150} stroke={chartAxis} tick={{ ...chartTick, fontSize: 12 }} interval={0} tickLine={false} /><Tooltip formatter={(v) => `${percentFmt.format(Number(v))}%`} contentStyle={chartTooltip} /><Legend wrapperStyle={{ color: 'var(--muted-text)', fontSize: 12 }} /><Bar dataKey="current" fill="var(--accent)" radius={[0, 6, 6, 0]} /><Bar dataKey="target" fill="var(--compare)" radius={[0, 6, 6, 0]} /></BarChart></ResponsiveContainer>
}
function BeforeAfter({ before, after }: { before: { name: string; value: number }[]; after: { name: string; value: number }[] }) {
  const names = Array.from(new Set([...before.map((i) => i.name), ...after.map((i) => i.name)]))
  const data = names.map((name) => ({ name, current: before.find((i) => i.name === name)?.value ?? 0, simulated: after.find((i) => i.name === name)?.value ?? 0 }))
  return <ResponsiveContainer width="100%" height="100%"><BarChart data={data} layout="vertical" margin={{ top: 4, right: 24, bottom: 24, left: 8 }}><XAxis type="number" stroke={chartAxis} tick={chartTick} tickLine={false} /><YAxis dataKey="name" type="category" width={150} stroke={chartAxis} tick={{ ...chartTick, fontSize: 12 }} interval={0} tickLine={false} /><Tooltip contentStyle={chartTooltip} /><Legend wrapperStyle={{ color: 'var(--muted-text)', fontSize: 12 }} /><Bar dataKey="current" fill="var(--compare)" /><Bar dataKey="simulated" fill="var(--accent)" /></BarChart></ResponsiveContainer>
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
function ActionTable({ actions, holdings, onChange }: { actions: RecompCandidate[]; holdings: Holding[]; onChange: (action: RecompCandidate) => void }) {
  if (!actions.length) return <div className="empty"><Warning size={28} /><p>No simulated actions yet. Generate auto-recomp candidates or add manual sandbox actions.</p></div>
  return <div className="overflow-auto rounded-xl border border-white/10"><table className="data-table"><thead><tr><th>Action</th><th>Ticker / bucket</th><th>Basis</th><th>Amount</th><th>Est. shares</th><th>Price used</th><th>Before</th><th>After</th><th>Reason</th><th>Impact</th></tr></thead><tbody>{actions.map((action) => {
    const price = priceUsedForAction(action, holdings)
    const shares = action.estimatedShares ?? (price ? action.dollarAmount / price : undefined)
    return <tr key={action.id}><td>{action.actionType}</td><td>{action.tickerOrBucket}</td><td>{action.basis ?? 'template drift'}</td><td><NumberCell value={action.dollarAmount} step={100} onChange={(value) => onChange({ ...action, dollarAmount: value, estimatedShares: price ? value / price : undefined, priceUsed: price })} /></td><td>{price ? <NumberCell value={shares ?? 0} step={0.01} onChange={(value) => onChange({ ...action, estimatedShares: value, dollarAmount: value * price, priceUsed: price })} /> : formatShares(shares)}</td><td>{formatPriceUsed(action, holdings)}</td><td>{percentFmt.format(action.beforeWeight)}%</td><td>{percentFmt.format(action.afterWeight)}%</td><td>{action.reason}</td><td>{action.alignmentImpact}</td></tr>
  })}</tbody></table></div>
}

function HoldingDeltaTable({ before, after }: { before: Holding[]; after: Holding[] }) {
  const rows = holdingDeltaRows(before, after)
  if (!rows.length) return <div className="empty"><ChartBar size={28} /><p>No holding-level changes yet. Generate candidates or add manual worksheet actions.</p></div>
  return <div className="overflow-auto rounded-xl border border-white/10"><table className="data-table data-table-compact"><thead><tr><th>Ticker</th><th>Current value</th><th>Current %</th><th>Simulated value</th><th>Simulated %</th><th>Delta</th><th>Shares after</th></tr></thead><tbody>{rows.map((row) => <tr key={row.ticker}><td>{row.ticker}</td><td>{dollarFmt.format(row.beforeValue)}</td><td>{percentFmt.format(row.beforeWeight)}%</td><td>{dollarFmt.format(row.afterValue)}</td><td>{percentFmt.format(row.afterWeight)}%</td><td className={row.delta >= 0 ? 'positive-delta' : 'negative-delta'}>{row.delta >= 0 ? '+' : ''}{dollarFmt.format(row.delta)}</td><td>{formatShares(row.afterShares)}</td></tr>)}</tbody></table></div>
}

function PositionPlans({ holdings, analytics, plans, setPlans }: { holdings: Holding[]; analytics: ReturnType<typeof analyze>; plans: Record<string, PositionPlan>; setPlans: React.Dispatch<React.SetStateAction<Record<string, PositionPlan>>> }) {
  const [selectedTicker, setSelectedTicker] = useState(holdings[0]?.ticker ?? '')
  const holding = holdings.find((item) => item.ticker === selectedTicker) ?? holdings[0]
  const plan = holding ? plans[holding.ticker] ?? defaultPositionPlan(holding, analytics.total) : undefined
  function updatePlan(next: PositionPlan) {
    setPlans((items) => ({ ...items, [next.ticker]: next }))
  }
  if (!holding || !plan) return <div className="empty"><PencilSimple size={28} /><p>No holdings available for position planning.</p></div>
  const currentWeight = weight(holding.marketValue, analytics.total)
  const flags = positionPlanFlags(plan, currentWeight)
  return <div className="grid gap-4 xl:grid-cols-[320px_1fr]">
    <div className="space-y-2">
      {holdings.slice(0, 40).map((item) => {
        const saved = Boolean(plans[item.ticker])
        return <button key={item.id} className={`row w-full text-left ${item.ticker === holding.ticker ? 'row-selected' : ''}`} onClick={() => setSelectedTicker(item.ticker)}>
          <span><strong>{item.ticker}</strong><span className="block text-xs text-muted">{percentFmt.format(weight(item.marketValue, analytics.total))}% - {item.securityName}</span></span>
          {saved && <span className="edited-pill">Plan</span>}
        </button>
      })}
    </div>
    <div className="position-plan-card">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div><h3>{holding.ticker} position plan</h3><p>{holding.securityName}</p></div>
        <span className="template-pill">Current {percentFmt.format(currentWeight)}%</span>
      </div>
      <div className="custom-target-grid">
        <label className="field"><span>Management style</span><select className="control" value={plan.style} onChange={(event) => updatePlan({ ...plan, style: event.target.value as PositionPlanStyle })}>{positionPlanStyles.map((style) => <option key={style}>{style}</option>)}</select></label>
        <label className="field"><span>Next review date</span><input className="control" type="date" value={plan.reviewDate} onChange={(event) => updatePlan({ ...plan, reviewDate: event.target.value })} /></label>
        <label className="field"><span>Target min %</span><input className="control" type="number" step={0.1} value={plan.targetMin} onChange={(event) => updatePlan({ ...plan, targetMin: Number(event.target.value) })} /></label>
        <label className="field"><span>Target weight %</span><input className="control" type="number" step={0.1} value={plan.targetWeight} onChange={(event) => updatePlan({ ...plan, targetWeight: Number(event.target.value) })} /></label>
        <label className="field"><span>Target max %</span><input className="control" type="number" step={0.1} value={plan.targetMax} onChange={(event) => updatePlan({ ...plan, targetMax: Number(event.target.value) })} /></label>
        <label className="field"><span>Profit-harvest review at gain %</span><input className="control" type="number" step={1} value={plan.profitHarvestTrigger ?? ''} onChange={(event) => updatePlan({ ...plan, profitHarvestTrigger: optionalNumber(event.target.value) })} /></label>
        <label className="field"><span>Drawdown review at loss %</span><input className="control" type="number" step={1} value={plan.drawdownReviewTrigger ?? ''} onChange={(event) => updatePlan({ ...plan, drawdownReviewTrigger: optionalNumber(event.target.value) })} /></label>
      </div>
      <div className="grid gap-3 md:grid-cols-2">
        <label className="field"><span>Thesis note</span><textarea className="control min-h-28" value={plan.thesis} onChange={(event) => updatePlan({ ...plan, thesis: event.target.value })} /></label>
        <label className="field"><span>Action note</span><textarea className="control min-h-28" value={plan.actionNote} onChange={(event) => updatePlan({ ...plan, actionNote: event.target.value })} /></label>
      </div>
      <div className="review-flags">
        {flags.length ? flags.map((flag) => <span key={flag}>{flag}</span>) : <span>No review flags from this plan.</span>}
      </div>
      <p className="text-sm leading-6 text-muted">Position Plans are notes and review triggers. They intentionally avoid market, limit, stop, trailing stop, day, or GTC order fields.</p>
    </div>
  </div>
}

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
  const stressImpact = holdings.reduce((sum, h) => sum + weight(h.marketValue, total) / 100 * shockFor(h, stress), 0)
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
  const nuisance = input.holdings.filter((holding) => isNuisanceHolding(holding, input.total, input.minDollar, input.minWeight)).length
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
      if (amount > minTrade) actions.push(candidate('Trim to target %', holding.ticker, amount, before, template.maxSingle, `Trim ${holding.ticker} toward the ${template.maxSingle}% max single-position guardrail.`, priceForHolding(holding), 'single-position cap'))
    }
  })
  analytics.aiBucketData.slice(0, 3).forEach((bucket) => {
    if (bucket.value > template.maxBucket) actions.push(candidate('Allocate cash to bucket', `Reduce ${bucket.name}`, analytics.total * 0.015, bucket.value, template.maxBucket, `${bucket.name} dominates the AI sleeve; redirect incremental exposure elsewhere.`, undefined, 'AI bucket overweight'))
  })
  assetClasses.forEach((asset) => {
    const current = weight(analytics.assetTotals[asset] ?? 0, analytics.total)
    const target = template.targets[asset]
    if (target - current > 3) actions.push(candidate('Allocate cash to bucket', asset, Math.min(analytics.total * ((target - current) / 100), analytics.total * 0.04), current, target, `${asset} is underweight versus the selected template.`, undefined, 'template underweight'))
  })
  analytics.holdings.filter((h) => h.marketValue < 1500).slice(0, 4).forEach((holding) => actions.push(candidate('Sell dollar amount', holding.ticker, holding.marketValue, weight(holding.marketValue, analytics.total), 0, `Consolidate tiny nuisance position to reduce portfolio clutter.`, priceForHolding(holding), 'nuisance position')))
  return actions.slice(0, 12)
}
function candidate(actionType: ActionType, tickerOrBucket: string, dollarAmount: number, beforeWeight: number, afterWeight: number, reason: string, price?: number, basis?: string): RecompCandidate {
  return { id: crypto.randomUUID(), actionType, tickerOrBucket, dollarAmount, estimatedShares: price ? dollarAmount / price : undefined, priceUsed: price, basis, beforeWeight, afterWeight, reason, riskImpact: 'Estimated risk-budget contribution recalculates in the simulated view.', alignmentImpact: 'Moves closer to selected template guardrails.', stressImpact: 'Scenario sensitivity updates after applying this simulated action.', warning: actionType.includes('Buy') && afterWeight > beforeWeight ? 'Check concentration before acting.' : undefined }
}
function simulateCandidates(holdings: Holding[], actions: RecompCandidate[]) {
  const simulated = holdings.map(cloneHolding)
  actions.forEach((action) => {
    const item = simulated.find((h) => h.ticker === action.tickerOrBucket)
    if (!item) return
    const direction = action.actionType.includes('Buy') ? 1 : -1
    const amount = direction < 0 ? Math.min(action.dollarAmount, item.marketValue) : action.dollarAmount
    item.marketValue = Math.max(0, item.marketValue + direction * amount)
    const price = priceForHolding(item)
    if (price) {
      item.price = price
      item.shares = item.marketValue / price
    }
    if (item.assetClass !== 'Cash') adjustCash(simulated, direction < 0 ? amount : -amount)
  })
  return simulated.filter((h) => h.marketValue > 1)
}
function adjustCash(holdings: Holding[], amount: number) {
  const cash = holdings.find((holding) => holding.assetClass === 'Cash')
  if (!cash) {
    if (amount > 0) holdings.push({ id: `simulation-cash-${crypto.randomUUID()}`, accountId: 'simulation', accountOwner: 'Portfolio', accountName: 'Simulation cash', accountType: 'Cash', accountCategory: 'Cash', ticker: 'CASH', securityName: 'Simulation cash', shares: amount, price: 1, marketValue: amount, assetClass: 'Cash', securityType: 'Cash / money market', sector: 'Cash & equivalents', ai: ai('Broad passive index exposure', 0, 'none'), notes: 'Created by simulation offset.' })
    return
  }
  cash.marketValue = Math.max(0, cash.marketValue + amount)
  cash.shares = cash.price ? cash.marketValue / cash.price : cash.marketValue
}
function normalizeImportedRow(row: Record<string, unknown>, accountOwner: string, accountName: string, accountType: string, index: number): Holding | null {
  const ticker = clean(row.Symbol ?? row.Ticker ?? row.ticker).toUpperCase()
  const marketValue = parseMoney(row['Mkt Val (Market Value)'] ?? row.market_value ?? row['Market Value'])
  if (!ticker || ticker.toLowerCase().includes('positions total') || ticker.toLowerCase() === 'no number') return null
  if (!ticker || !Number.isFinite(marketValue) || marketValue <= 0) return null
  const shares = parseMoney(row['Qty (Quantity)'] ?? row.shares ?? row.Quantity) || 0
  const explicitPrice = parseMoney(row.Price ?? row.price ?? row['Last Price'] ?? row['Current Price'] ?? row['Price ($)'])
  const price = explicitPrice || (shares ? marketValue / shares : undefined)
  const defaults = defaultClassifications[ticker] ?? {}
  const normalizedAccountName = accountName.replace(/^Positions for account\s*/i, '')
  const securityName = clean(row.Description ?? row.security_name ?? row.Name) || ticker
  const rawAssetType = clean(row['Asset Type'])
  const assetClass = (defaults.assetClass as AssetClass) ?? assetFromRaw(rawAssetType)
  return { id: `${accountName}-${ticker}-${index}-${crypto.randomUUID()}`, accountId: slug(normalizedAccountName), accountOwner, accountName: normalizedAccountName, accountType, accountCategory: detectAccountType(normalizedAccountName), ticker, securityName, shares, price, marketValue, assetClass, securityType: (defaults.securityType as SecurityType | undefined) ?? inferSecurityType(ticker, securityName, assetClass, rawAssetType), sector: defaults.sector ?? 'Unclassified', ai: (defaults.ai as AIExposureClassification) ?? { score: 0, buckets: [], directness: 'none', confidence: 'low', source: 'unknown', notes: '' }, costBasis: parseMoney(row['Cost Basis']), unrealizedGainLoss: parseMoney(row['Gain $ (Gain/Loss $)']) }
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
function targetAllocationData(template: StrategyTemplate) {
  return assetClasses.map((asset) => ({ name: asset, value: template.targets[asset] })).filter((item) => item.value > 0)
}
function compactChartData(data: { name: string; value: number }[], limit: number) {
  if (data.length <= limit) return data
  const head = data.slice(0, limit - 1)
  const other = data.slice(limit - 1).reduce((sum, item) => sum + item.value, 0)
  return [...head, { name: 'Remaining AI buckets', value: other }]
}
function holdingBarData(holdings: Holding[], total: number): BarDatum[] {
  return holdings.map((holding) => {
    const pct = weight(holding.marketValue, total)
    return {
      name: holding.ticker,
      value: pct,
      label: `${percentFmt.format(pct)}% · ${compactDollar(holding.marketValue)}`,
      color: holdingColor(holding),
    }
  })
}
function cumulativeWeights(holdings: Holding[], total: number) {
  let cumulative = 0
  const values = new Map<string, number>()
  holdings.forEach((holding) => {
    cumulative += weight(holding.marketValue, total)
    values.set(holding.id, cumulative)
  })
  return values
}
function concentrationCutoffs(holdings: Holding[], total: number) {
  const result: Partial<Record<50 | 80, { holding: Holding; rank: number; before: number; cumulative: number }>> = {}
  let cumulative = 0
  holdings.forEach((holding, index) => {
    const before = cumulative
    cumulative += weight(holding.marketValue, total)
    if (!result[50] && cumulative >= 50) result[50] = { holding, rank: index + 1, before, cumulative }
    if (!result[80] && cumulative >= 80) result[80] = { holding, rank: index + 1, before, cumulative }
  })
  return result
}
function compactDollar(value: number) {
  if (Math.abs(value) >= 1_000_000) return `$${(value / 1_000_000).toFixed(1)}M`
  if (Math.abs(value) >= 1_000) return `$${Math.round(value / 1_000)}K`
  return dollarFmt.format(value)
}
function formatShares(value?: number) {
  if (!Number.isFinite(value)) return '-'
  return shareFmt.format(value ?? 0)
}
function priceForHolding(holding?: Holding) {
  return holding?.price ?? (holding?.shares ? holding.marketValue / holding.shares : undefined)
}
function priceUsedForAction(action: RecompCandidate, holdings: Holding[]) {
  return action.priceUsed ?? priceForHolding(holdings.find((holding) => holding.ticker === action.tickerOrBucket))
}
function formatPriceUsed(action: RecompCandidate, holdings: Holding[]) {
  const price = priceUsedForAction(action, holdings)
  if (price) return dollarFmt.format(price)
  if ([...assetClasses, ...aiBuckets].includes(action.tickerOrBucket as AssetClass | AIBucket)) return 'N/A'
  return 'Price missing'
}
function holdingDeltaRows(before: Holding[], after: Holding[]) {
  const beforeTotal = before.reduce((sum, holding) => sum + holding.marketValue, 0)
  const afterTotal = after.reduce((sum, holding) => sum + holding.marketValue, 0)
  const beforeMap = new Map(before.map((holding) => [holding.ticker, holding]))
  const afterMap = new Map(after.map((holding) => [holding.ticker, holding]))
  const tickers = Array.from(new Set([...beforeMap.keys(), ...afterMap.keys()]))
  return tickers.map((ticker) => {
    const beforeHolding = beforeMap.get(ticker)
    const afterHolding = afterMap.get(ticker)
    const beforeValue = beforeHolding?.marketValue ?? 0
    const afterValue = afterHolding?.marketValue ?? 0
    return {
      ticker,
      beforeValue,
      afterValue,
      beforeWeight: weight(beforeValue, beforeTotal),
      afterWeight: weight(afterValue, afterTotal),
      delta: afterValue - beforeValue,
      afterShares: afterHolding?.shares ?? 0,
    }
  }).filter((row) => Math.abs(row.delta) > 0.5).sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta)).slice(0, 20)
}
const positionPlanStyles: PositionPlanStyle[] = ['Passive', 'Core', 'Tactical', 'Speculative', 'Income / ballast', 'Review only']
function defaultPositionPlan(holding: Holding, total: number): PositionPlan {
  const current = weight(holding.marketValue, total)
  return {
    ticker: holding.ticker,
    style: holding.assetClass === 'Bonds/fixed income' || holding.assetClass === 'Cash' ? 'Income / ballast' : holding.ai.score >= 3 ? 'Tactical' : 'Core',
    targetMin: Math.max(0, Number((current * 0.75).toFixed(1))),
    targetWeight: Number(current.toFixed(1)),
    targetMax: Number(Math.max(current * 1.25, current + 0.5).toFixed(1)),
    reviewDate: '',
    thesis: '',
    actionNote: '',
  }
}
function positionPlanFlags(plan: PositionPlan, currentWeight: number) {
  const flags: string[] = []
  if (currentWeight > plan.targetMax) flags.push(`Above target max by ${percentFmt.format(currentWeight - plan.targetMax)} pts`)
  if (currentWeight < plan.targetMin) flags.push(`Below target min by ${percentFmt.format(plan.targetMin - currentWeight)} pts`)
  if (plan.reviewDate && new Date(plan.reviewDate) <= new Date()) flags.push('Review date due')
  if (plan.profitHarvestTrigger) flags.push(`Profit-harvest review note at +${plan.profitHarvestTrigger}% gain`)
  if (plan.drawdownReviewTrigger) flags.push(`Drawdown review note at -${plan.drawdownReviewTrigger}% loss`)
  return flags
}
function optionalNumber(value: string) {
  const parsed = Number(value)
  return Number.isFinite(parsed) && value !== '' ? parsed : undefined
}
function aggregateHoldingsForView(holdings: Holding[]): Holding[] {
  const groups = new Map<string, Holding[]>()
  holdings.forEach((holding) => {
    const key = holding.ticker.toUpperCase()
    groups.set(key, [...(groups.get(key) ?? []), holding])
  })
  return Array.from(groups.entries()).map(([ticker, group]) => {
    if (group.length === 1) return cloneHolding(group[0])
    const totalValue = group.reduce((sum, holding) => sum + holding.marketValue, 0)
    const shares = group.reduce((sum, holding) => sum + holding.shares, 0)
    const dominant = [...group].sort((a, b) => b.marketValue - a.marketValue)[0]
    return {
      ...cloneHolding(dominant),
      id: `view-${ticker}-${group.map((holding) => holding.id).sort().join('-')}`,
      sourceHoldingIds: group.map((holding) => holding.id),
      accountId: group.map((holding) => holding.accountId).join('|'),
      accountName: uniqueText(group.map((holding) => holding.accountName)).join(' + '),
      accountType: uniqueText(group.map((holding) => holding.accountType)).join(' + '),
      accountCategory: uniqueText(group.map((holding) => holding.accountCategory)).join(' + '),
      ticker,
      securityName: dominant.securityName,
      shares,
      price: shares ? totalValue / shares : dominant.price,
      marketValue: totalValue,
      assetClass: dominantByValue(group, (holding) => holding.assetClass) as AssetClass,
      securityType: dominantByValue(group, (holding) => holding.securityType) as SecurityType,
      sector: dominantByValue(group, (holding) => holding.sector),
      ai: aggregateAI(group),
      costBasis: group.reduce((sum, holding) => sum + (holding.costBasis ?? 0), 0),
      unrealizedGainLoss: group.reduce((sum, holding) => sum + (holding.unrealizedGainLoss ?? 0), 0),
      notes: uniqueText(group.map((holding) => holding.notes).filter(Boolean) as string[]).join(' | '),
    }
  }).sort((a, b) => b.marketValue - a.marketValue)
}
function tableFilters(holdings: Holding[]) {
  return {
    assets: uniqueText(holdings.map((holding) => holding.assetClass)).sort(),
    sectors: uniqueText(holdings.map((holding) => holding.sector)).sort(),
    aiBuckets: uniqueText(holdings.map((holding) => holding.ai.buckets[0]?.bucket ?? 'Missing')).sort(),
  }
}
function filterHoldingsForTable(holdings: Holding[], asset: string, sector: string, aiBucket: string, sizeFilter: string, minDollar: number, minWeight: number, total: number) {
  return holdings.filter((holding) => {
    if (asset !== 'All' && holding.assetClass !== asset) return false
    if (sector !== 'All' && holding.sector !== sector) return false
    if (aiBucket !== 'All' && (holding.ai.buckets[0]?.bucket ?? 'Missing') !== aiBucket) return false
    const nuisance = isNuisanceHolding(holding, total, minDollar, minWeight)
    if (sizeFilter === 'Below nuisance threshold' && !nuisance) return false
    if (sizeFilter === 'Above nuisance threshold' && nuisance) return false
    return true
  })
}
function aggregateAI(holdings: Holding[]): AIExposureClassification {
  const total = holdings.reduce((sum, holding) => sum + holding.marketValue, 0)
  const score = total ? holdings.reduce((sum, holding) => sum + holding.marketValue * holding.ai.score, 0) / total : 0
  const bucketValues: Record<string, number> = {}
  const directnessValues: Record<string, number> = {}
  holdings.forEach((holding) => {
    const aiValue = holding.marketValue * (holding.ai.score / 5)
    directnessValues[holding.ai.directness] = (directnessValues[holding.ai.directness] ?? 0) + aiValue
    holding.ai.buckets.forEach((bucket) => {
      bucketValues[bucket.bucket] = (bucketValues[bucket.bucket] ?? 0) + aiValue * (bucket.weight / 100)
    })
  })
  const bucketTotal = Object.values(bucketValues).reduce((sum, value) => sum + value, 0)
  const buckets = bucketTotal ? Object.entries(bucketValues).map(([bucket, value]) => ({ bucket: bucket as AIBucket, weight: weight(value, bucketTotal) })).sort((a, b) => b.weight - a.weight) : [{ bucket: 'Broad passive index exposure' as AIBucket, weight: 100 }]
  const directness = Object.entries(directnessValues).sort((a, b) => b[1] - a[1])[0]?.[0] as Directness | undefined
  return {
    score,
    buckets,
    directness: directness ?? 'none',
    confidence: holdings.some((holding) => holding.ai.confidence === 'low') ? 'low' : holdings.some((holding) => holding.ai.confidence === 'medium') ? 'medium' : 'high',
    source: holdings.some((holding) => holding.ai.source === 'manual') ? 'manual' : holdings.some((holding) => holding.ai.source === 'imported') ? 'imported' : holdings.some((holding) => holding.ai.source === 'default') ? 'default' : 'unknown',
    notes: uniqueText(holdings.map((holding) => holding.ai.notes).filter(Boolean)).join(' | '),
  }
}
function dominantByValue(holdings: Holding[], key: (holding: Holding) => string) {
  const totals = new Map<string, number>()
  holdings.forEach((holding) => totals.set(key(holding), (totals.get(key(holding)) ?? 0) + holding.marketValue))
  return Array.from(totals.entries()).sort((a, b) => b[1] - a[1])[0]?.[0] ?? ''
}
function uniqueText(values: string[]) {
  return Array.from(new Set(values.filter(Boolean)))
}
function cloneHolding(holding: Holding): Holding {
  return { ...holding, ai: { ...holding.ai, buckets: [...holding.ai.buckets] } }
}
function applyHoldingEditOverlay(holdings: Holding[], overlay: Record<string, Partial<Holding>>) {
  return holdings.map((holding) => {
    const edit = overlay[holding.id]
    if (!edit) return cloneHolding(holding)
    return {
      ...cloneHolding(holding),
      ...edit,
      ai: edit.ai ? { ...holding.ai, ...edit.ai, buckets: edit.ai.buckets ? [...edit.ai.buckets] : [...holding.ai.buckets] } : { ...holding.ai, buckets: [...holding.ai.buckets] },
    }
  })
}
function isHoldingEdited(id: string, overlay: Record<string, Partial<Holding>>) {
  return Boolean(overlay[id])
}
function holdingEditedInView(holding: Holding, overlay: Record<string, Partial<Holding>>) {
  return (holding.sourceHoldingIds ?? [holding.id]).some((id) => isHoldingEdited(id, overlay))
}
function holdingExcludedInView(holding: Holding, excludedIds: string[]) {
  const ids = holding.sourceHoldingIds ?? [holding.id]
  return ids.some((id) => excludedIds.includes(id))
}
function repriceHolding(holding: Holding): Holding {
  if (!holding.price || !holding.shares) return holding
  return { ...holding, marketValue: holding.shares * holding.price }
}
function isNuisanceHolding(holding: Holding, total: number, minDollar: number, minWeight: number) {
  return holding.marketValue < minDollar || weight(holding.marketValue, total) < minWeight
}
function warningStatus(holding: Holding, total: number, minDollar: number, minWeight: number, maxSingle: number) {
  return isNuisanceHolding(holding, total, minDollar, minWeight) ? 'Nuisance' : weight(holding.marketValue, total) > maxSingle ? 'High concentration' : holding.ai.source === 'unknown' ? 'Missing AI data' : 'Clear'
}
function scenarioImpactDollars(analytics: ReturnType<typeof analyze>) {
  return analytics.total * (analytics.stressImpact / 100)
}
function shortChartLabel(value: string) {
  return value
    .replace('Big tech / hyperscalers', 'Big tech')
    .replace('AI platforms / AI software', 'AI software')
    .replace('Broad passive index exposure', 'Broad index')
    .replace('Public indirect AI lab exposure', 'AI lab indirect')
    .replace('Data centers / colocation', 'Data centers')
    .replace('Cooling / thermal management', 'Cooling')
    .replace('Semiconductor equipment', 'Semi equip.')
    .replace('Grid / electrification', 'Grid')
}
function weight(value: number, total: number) { return total ? (value / total) * 100 : 0 }
function parseMoney(value: unknown) { const cleanValue = String(value ?? '').replace(/[$,%"]/g, '').replace(/,/g, '').trim(); const parsed = Number(cleanValue); return Number.isFinite(parsed) ? parsed : 0 }
function clean(value: unknown) { return String(value ?? '').replaceAll('"', '').trim() }
function assetFromRaw(raw: string): AssetClass { return raw.toLowerCase().includes('cash') ? 'Cash' : raw.toLowerCase().includes('bond') ? 'Bonds/fixed income' : 'Broad US equity' }
function inferSecurityType(ticker: string, name: string, assetClass: AssetClass, rawAssetType = ''): SecurityType {
  const text = `${ticker} ${name} ${rawAssetType}`.toLowerCase()
  if (assetClass === 'Cash' || text.includes('cash') || text.includes('money market')) return 'Cash / money market'
  if (assetClass === 'Bonds/fixed income' || text.includes('fixed income') || /\b\d{6}[a-z0-9]{3}\b/i.test(ticker)) return 'Bond / fixed income'
  if (text.includes('mutual fund')) return 'Mutual fund'
  if (text.includes('etf') || text.includes('closed end fund') || text.includes('trust etf')) return 'ETF / closed-end fund'
  if (text.includes('warrant') || text.includes('wts') || text.includes('right') || text.includes('rts')) return 'Option / warrant / right'
  if (text.includes('bitcoin') || text.includes('ethereum') || text.includes('crypto')) return 'Crypto / digital asset'
  if (rawAssetType.toLowerCase().includes('equity') || assetClass === 'Broad US equity' || assetClass === 'AI buildout sleeve' || assetClass === 'International equity') return 'Individual equity'
  return 'Other'
}
function stressExplanation(scenario: StressScenario) {
  return `Scenario sensitivity applies this scenario's simple shocks to each holding by AI bucket first, then asset class fallback, and sums the weighted impact. It is a what-if comparison tool, not a forecast, volatility model, or institutional risk system. Current scenario: ${scenario.name}.`
}
function assetShortLabel(asset: AssetClass) {
  return {
    'Broad US equity': 'US',
    'AI buildout sleeve': 'AI',
    'International equity': 'Intl',
    'Bonds/fixed income': 'Bonds',
    Cash: 'Cash',
    'Alternatives/other': 'Alt',
  }[asset]
}
function summarizeStagedAccounts(rows: Record<string, unknown>[]) {
  const summaries = new Map<string, { id: string; name: string; type: string; rows: number; value: number }>()
  rows.forEach((row) => {
    const id = String(row.__accountId ?? 'unknown-account')
    const current = summaries.get(id) ?? { id, name: String(row.__accountName ?? 'Uploaded account'), type: String(row.__accountType ?? 'Unknown account'), rows: 0, value: 0 }
    current.rows += 1
    current.value += parseMoney(row['Mkt Val (Market Value)'] ?? row.market_value ?? row['Market Value'])
    summaries.set(id, current)
  })
  return Array.from(summaries.values()).sort((a, b) => b.value - a.value)
}
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
function slug(value: string) { return value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') || 'unknown-account' }
function detectAccountType(value: string) {
  const text = value.toLowerCase()
  if (text.includes('401k') || text.includes('401(k)')) return '401k'
  if (text.includes('sep')) return 'SEP IRA'
  if (text.includes('simple')) return 'SIMPLE IRA'
  if (text.includes('rollover')) return 'Rollover IRA'
  if (text.includes('roth')) return 'Roth IRA'
  if (text.includes('traditional') || /\bira\b/.test(text)) return 'Traditional IRA'
  if (text.includes('hsa')) return 'HSA'
  if (text.includes('trust')) return 'Trust'
  if (text.includes('taxable') || text.includes('brokerage') || text.includes('joint')) return 'Taxable'
  if (text.includes('db plan') || text.includes('pension')) return 'Defined benefit plan'
  return 'Unknown account'
}
function detectAccountFromText(preamble: string, filename: string) {
  const stripped = preamble.replaceAll('"', '').trim()
  const match = stripped.match(/Positions for account\s+(.+?)\s+as of/i)
  const rawName = match?.[1] ?? filename.replace(/\.(csv|xlsx|xls)$/i, '').replace(/-Positions-.+$/i, '')
  const name = rawName.replace(/\s*\.\.\.\d+\s*/g, '').trim() || 'Uploaded account'
  return { id: slug(name), name, type: detectAccountType(name) }
}
function detectAsOfDate(preamble: string) {
  const match = preamble.match(/as of .*?,\s*(\d{4})\/(\d{2})\/(\d{2})/i)
  if (!match) return ''
  return `${match[1]}-${match[2]}-${match[3]}`
}
function accountScopes(holdings: Holding[]) {
  const accounts = Array.from(new Map(holdings.map((holding) => [holding.accountId, holding])).values())
  return [{ id: 'combined', label: 'Combined portfolio' }, ...accounts.map((holding) => ({ id: holding.accountId, label: `${holding.accountName} · ${holding.accountCategory || holding.accountType}` }))]
}
function resolveTheme(theme: ThemePreference) {
  if (theme !== 'system') return theme
  return window.matchMedia?.('(prefers-color-scheme: light)').matches ? 'light' : 'dark'
}
function loadState() {
  try {
    const parsed = JSON.parse(localStorage.getItem(storageKey) || '{}')
    const snapshots = parsed.snapshots?.length ? parsed.snapshots.map((snapshot: PortfolioSnapshot) => ({ ...snapshot, holdings: snapshot.holdings.map(migrateHolding).filter((holding): holding is Holding => Boolean(holding)) })) : sampleSnapshots()
    return { snapshots, selectedSnapshotId: parsed.selectedSnapshotId ?? 'sample-current', selectedAccountScope: parsed.selectedAccountScope ?? 'combined', selectedTemplateId: parsed.selectedTemplateId ?? 'diversified-ai-supply-chain', customTemplates: parsed.customTemplates ?? [], themePreference: parsed.themePreference ?? 'system', ledgerMode: parsed.ledgerMode ?? 'uploaded', holdingEditOverlay: parsed.holdingEditOverlay ?? {}, hiddenHoldingIds: parsed.hiddenHoldingIds ?? [], recompCandidates: parsed.recompCandidates ?? [], manualActions: parsed.manualActions ?? [], positionPlans: parsed.positionPlans ?? {}, decisionLog: parsed.decisionLog ?? [] }
  } catch {
    return { snapshots: sampleSnapshots(), selectedSnapshotId: 'sample-current', selectedAccountScope: 'combined', selectedTemplateId: 'diversified-ai-supply-chain', customTemplates: [], themePreference: 'system' as ThemePreference, ledgerMode: 'uploaded' as LedgerMode, holdingEditOverlay: {}, hiddenHoldingIds: [], recompCandidates: [], manualActions: [], positionPlans: {}, decisionLog: [] }
  }
}
function migrateHolding(holding: Holding): Holding | null {
  if (!holding.ticker || holding.ticker.toLowerCase().includes('positions total') || holding.ticker.toLowerCase() === 'no number') return null
  const ticker = holding.ticker.toUpperCase()
  const defaults = defaultClassifications[ticker] ?? {}
  const aiDefault = defaults.ai as AIExposureClassification | undefined
  const assetClass = (defaults.assetClass as AssetClass | undefined) ?? holding.assetClass
  const inferredType = inferSecurityType(ticker, holding.securityName, assetClass)
  const price = holding.price || (holding.shares ? holding.marketValue / holding.shares : undefined)
  return {
    ...holding,
    ticker,
    price,
    accountId: holding.accountId ?? slug(holding.accountName),
    accountCategory: holding.accountCategory ?? detectAccountType(holding.accountName),
    assetClass,
    securityType: (defaults.securityType as SecurityType | undefined) ?? (holding.securityType && holding.securityType !== 'Other' ? holding.securityType : inferredType),
    sector: (defaults.sector as string | undefined) ?? holding.sector,
    ai: holding.ai?.source === 'manual' ? holding.ai : aiDefault ?? holding.ai ?? { score: 0, buckets: [], directness: 'none', confidence: 'low', source: 'unknown', notes: '' },
  }
}
function persist(state: PersistedState) {
  localStorage.setItem(storageKey, JSON.stringify(state))
}

export default App
