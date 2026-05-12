import { aiBuckets, assetClasses, assetClassColors, semanticColors } from './domain'
import type { AIBucket, AssetClass, Holding } from './domain'

const fallbackColors = ['#047857', '#2563eb', '#b45309', '#6b7280', '#0f766e', '#be123c', '#475569']
const aiBucketColors: Partial<Record<AIBucket, string>> = {
  'Big tech / hyperscalers': '#047857',
  'AI platforms / AI software': '#0f766e',
  'Public indirect AI lab exposure': '#0e7490',
  Semiconductors: '#0369a1',
  'GPUs / accelerators': '#1d4ed8',
  'CPUs / general compute': '#2563eb',
  'Memory / storage': '#7c3aed',
  'Semiconductor equipment': '#0891b2',
  'Fabs / foundries': '#0d9488',
  'Photonics / optical / interconnect': '#0284c7',
  Networking: '#155e75',
  'Data centers / colocation': '#0f766e',
  'Power generation': '#b45309',
  'Grid / electrification': '#ca8a04',
  'Cooling / thermal management': '#0d9488',
  'Materials / specialty chemicals': '#be123c',
  Cybersecurity: '#475569',
  'Enterprise software': '#059669',
  'Robotics / automation': '#64748b',
  'Broad passive index exposure': semanticColors.index,
}

export const holdingColorLegend = [
  { label: 'AI exposure', color: semanticColors.ai },
  { label: 'Index / ETF / fund', color: semanticColors.index },
  { label: 'Non-AI stock', color: semanticColors.equity },
  { label: 'Cash', color: semanticColors.cash },
  { label: 'Bonds', color: semanticColors.bonds },
  { label: 'Alternatives', color: semanticColors.alternatives },
]

export function holdingColor(holding: Holding) {
  if (holding.assetClass === 'Cash') return semanticColors.cash
  if (holding.assetClass === 'Bonds/fixed income') return semanticColors.bonds
  if (isFundHolding(holding) || isBroadIndexHolding(holding)) return semanticColors.index
  if (holding.ai.score > 0) return semanticColors.ai
  if (holding.assetClass === 'Alternatives/other') return semanticColors.alternatives
  if (holding.securityType === 'Individual equity') return semanticColors.equity
  if (holding.assetClass === 'International equity') return semanticColors.equity
  if (holding.assetClass === 'Broad US equity') return semanticColors.equity
  return semanticColors.other
}

export function chartColor(name: string, index: number) {
  const lower = name.toLowerCase()
  if (assetClasses.includes(name as AssetClass)) return assetClassColors[name as AssetClass]
  if (aiBuckets.includes(name as AIBucket)) return aiBucketColors[name as AIBucket] ?? semanticColors.ai
  if (lower.includes('remaining ai buckets')) return semanticColors.other
  if (lower.includes('cash') || lower.includes('money market')) return semanticColors.cash
  if (lower.includes('broad us') || lower.includes('broad passive') || lower.includes('index') || lower.includes('s&p') || lower.includes('nasdaq')) return semanticColors.index
  if (lower.includes('bond') || lower.includes('fixed income')) return semanticColors.bonds
  if (lower.includes('international') || lower.includes('emerging')) return semanticColors.international
  if (lower.includes('ai') || lower.includes('semiconductor') || lower.includes('gpu') || lower.includes('data center') || lower.includes('grid') || lower.includes('power') || lower.includes('cooling') || lower.includes('networking') || lower.includes('photonics')) return semanticColors.ai
  if (lower.includes('alternative') || lower.includes('energy') || lower.includes('digital') || lower.includes('gold') || lower.includes('miner')) return semanticColors.alternatives
  return fallbackColors[index % fallbackColors.length]
}

function isFundHolding(holding: Holding) {
  return holding.securityType === 'ETF / closed-end fund' || holding.securityType === 'Mutual fund'
}

function isBroadIndexHolding(holding: Holding) {
  const text = `${holding.ticker} ${holding.securityName} ${holding.sector}`.toLowerCase()
  return holding.assetClass === 'Broad US equity' && (holding.sector === 'Broad Market' || /\b(index|s&p|nasdaq|dow|total stock|equal weight|dividend aristocrats|covered call|vanguard|spdr|invesco qqq)\b/.test(text))
}
