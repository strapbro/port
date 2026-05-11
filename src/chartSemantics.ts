import { aiBuckets, assetClasses, assetClassColors, semanticColors } from './domain'
import type { AIBucket, AssetClass, Holding } from './domain'

const fallbackColors = ['#047857', '#2563eb', '#b45309', '#6b7280', '#0f766e', '#be123c', '#475569']

export const holdingColorLegend = [
  { label: 'AI exposure', color: semanticColors.ai },
  { label: 'Index / ETF', color: semanticColors.index },
  { label: 'Individual equity', color: semanticColors.equity },
  { label: 'Cash', color: semanticColors.cash },
  { label: 'Bonds', color: semanticColors.bonds },
  { label: 'Alternatives', color: semanticColors.alternatives },
]

export function holdingColor(holding: Holding) {
  if (holding.assetClass === 'Cash') return semanticColors.cash
  if (holding.assetClass === 'Bonds/fixed income') return semanticColors.bonds
  if (isFundHolding(holding) || isBroadIndexHolding(holding)) return semanticColors.index
  if (holding.ai.score > 0) return semanticColors.ai
  if (holding.assetClass === 'International equity') return semanticColors.international
  if (holding.assetClass === 'Broad US equity') return semanticColors.equity
  if (holding.securityType === 'Individual equity') return semanticColors.equity
  if (holding.assetClass === 'Alternatives/other') return semanticColors.alternatives
  return semanticColors.other
}

export function chartColor(name: string, index: number) {
  const lower = name.toLowerCase()
  if (assetClasses.includes(name as AssetClass)) return assetClassColors[name as AssetClass]
  if (aiBuckets.includes(name as AIBucket)) return name === 'Broad passive index exposure' ? semanticColors.index : semanticColors.ai
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
