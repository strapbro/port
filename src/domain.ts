export type AssetClass = 'Broad US equity' | 'AI buildout sleeve' | 'International equity' | 'Bonds/fixed income' | 'Cash' | 'Alternatives/other'
export type SecurityType = 'Individual equity' | 'ETF / closed-end fund' | 'Mutual fund' | 'Bond / fixed income' | 'Cash / money market' | 'Option / warrant / right' | 'Crypto / digital asset' | 'Other'
export type Directness = 'direct' | 'indirect' | 'none'
export type ClassificationSource = 'manual' | 'default' | 'imported' | 'unknown'
export type WarningSeverity = 'high' | 'medium' | 'low'
export type ActionType = 'Trim to target %' | 'Sell dollar amount' | 'Buy dollar amount' | 'Allocate cash to bucket'
export type ThemePreference = 'system' | 'dark' | 'light'
export type LedgerMode = 'uploaded' | 'sandbox'

export type AIBucket =
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

export type BucketExposure = { bucket: AIBucket; weight: number }
export type AIExposureClassification = {
  score: number
  buckets: BucketExposure[]
  directness: Directness
  confidence: 'high' | 'medium' | 'low'
  notes: string
  source: ClassificationSource
}
export type Holding = {
  id: string
  sourceHoldingIds?: string[]
  accountId: string
  accountOwner: string
  accountName: string
  accountType: string
  accountCategory: string
  ticker: string
  securityName: string
  shares: number
  price?: number
  marketValue: number
  assetClass: AssetClass
  securityType: SecurityType
  sector: string
  ai: AIExposureClassification
  costBasis?: number
  unrealizedGainLoss?: number
  notes?: string
}
export type PortfolioSnapshot = { id: string; name: string; date: string; source: string; holdings: Holding[] }
export type StrategyTemplate = {
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
export type RecompCandidate = {
  id: string
  actionType: ActionType
  tickerOrBucket: string
  dollarAmount: number
  estimatedShares?: number
  priceUsed?: number
  beforeWeight: number
  afterWeight: number
  reason: string
  riskImpact: string
  alignmentImpact: string
  stressImpact: string
  warning?: string
}
export type ManualSandboxAction = RecompCandidate & { note: string }
export type StressScenario = { id: string; name: string; shocks: Partial<Record<AssetClass | AIBucket, number>> }
export type AppWarning = { id: string; severity: WarningSeverity; title: string; detail: string }
export type DecisionLogEntry = {
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
export type PersistedState = {
  snapshots: PortfolioSnapshot[]
  selectedSnapshotId: string
  selectedAccountScope: string
  selectedTemplateId: string
  customTemplates: StrategyTemplate[]
  themePreference: ThemePreference
  ledgerMode: LedgerMode
  holdingEditOverlay: Record<string, Partial<Holding>>
  hiddenHoldingIds: string[]
  recompCandidates: RecompCandidate[]
  manualActions: ManualSandboxAction[]
  decisionLog: DecisionLogEntry[]
}

export const assetClasses: AssetClass[] = ['Broad US equity', 'AI buildout sleeve', 'International equity', 'Bonds/fixed income', 'Cash', 'Alternatives/other']
export const aiBuckets: AIBucket[] = [
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

export const semanticColors = {
  ai: '#047857',
  index: '#2563eb',
  cash: '#b45309',
  bonds: '#6b7280',
  international: '#0f766e',
  alternatives: '#be123c',
  equity: '#0891b2',
  other: '#71717a',
}
export const assetClassColors: Record<AssetClass, string> = {
  'Broad US equity': semanticColors.index,
  'AI buildout sleeve': semanticColors.ai,
  'International equity': semanticColors.international,
  'Bonds/fixed income': semanticColors.bonds,
  Cash: semanticColors.cash,
  'Alternatives/other': semanticColors.alternatives,
}

export function ai(bucket: AIBucket, score: number, directness: Directness, extra: BucketExposure[] = [], notes = ''): AIExposureClassification {
  const primaryWeight = Math.max(0, 100 - extra.reduce((sum, item) => sum + item.weight, 0))
  return { score, directness, confidence: 'medium', source: 'default', notes, buckets: [{ bucket, weight: primaryWeight }, ...extra] }
}

function template(id: string, name: string, purpose: string, targetValues: number[], aiRange: [number, number], maxSingle: number, maxTop10: number, maxBucket: number, risk: string, drawdownRange: string, upside: number, complexity: number, pros: string[], cons: string[], bestFor: string, warningLabel: string): StrategyTemplate {
  return { id, name, purpose, targets: Object.fromEntries(assetClasses.map((asset, index) => [asset, targetValues[index]])) as Record<AssetClass, number>, aiRange, maxSingle, maxTop10, maxBucket, risk, drawdownRange, upside, complexity, pros, cons, bestFor, warningLabel }
}

export const strategyTemplates: StrategyTemplate[] = [
  template('balanced-ai-growth', 'Balanced AI 25', 'Balanced growth profile with a 25% AI buildout sleeve and normal ballast.', [40, 25, 10, 15, 5, 5], [20, 30], 7, 45, 35, 'Moderate growth', '-18% to -30%', 7, 5, ['Clear AI participation', 'Diversified core remains intact', 'Keeps bonds and cash in the mix'], ['Still exposed to tech drawdowns', 'May trail a concentrated AI rally'], 'A growth portfolio that wants AI exposure without becoming a pure tech bet.', 'Watch single-name concentration'),
  template('aggressive-ai-infrastructure', 'Aggressive AI 40', 'High-conviction profile with a 40% AI infrastructure sleeve and wider guardrails.', [30, 40, 5, 10, 5, 10], [35, 45], 10, 60, 40, 'High growth', '-25% to -42%', 9, 7, ['Highest AI participation', 'Broad compute, power, and data center exposure'], ['Higher volatility', 'Sensitive to AI capex disappointment'], 'An intentionally aggressive AI infrastructure bet.', 'Higher volatility'),
  template('ai-barbell', 'Barbell AI 35', 'Pairs a 35% AI sleeve with larger cash and bond ballast.', [25, 35, 5, 20, 10, 5], [30, 40], 8, 50, 35, 'Barbell growth', '-20% to -36%', 8, 6, ['Meaningful AI upside', 'More defensive ballast', 'Keeps dry powder available'], ['Cash and bonds can drag', 'Still carries AI theme volatility'], 'Bullish on AI, but uncomfortable with full high-beta exposure.', 'Barbell volatility'),
  {
    ...template('diversified-ai-supply-chain', 'Supply Chain AI 30', '30% AI sleeve spread across chips, power, data centers, cooling, networking, and software.', [35, 30, 10, 15, 5, 5], [25, 35], 7, 50, 30, 'Diversified growth', '-20% to -34%', 8, 7, ['Best match for broad AI infrastructure', 'Reduces mega-cap/chip dependence', 'Includes second-order beneficiaries'], ['More moving parts', 'May lag a chip-led rally'], 'A portfolio built around the whole AI supply chain.', 'Balance the AI sleeve'),
    aiInternalSplit: {
      'Big tech / hyperscalers': 20,
      Semiconductors: 25,
      'Power generation': 10,
      'Grid / electrification': 10,
      'Data centers / colocation': 10,
      'Cooling / thermal management': 5,
      Networking: 5,
      'Photonics / optical / interconnect': 5,
      'AI platforms / AI software': 5,
      Cybersecurity: 5,
    },
  },
  template('conservative-ai-participation', 'Conservative AI 15', 'Lower-volatility profile with a 15% AI sleeve and stronger bond ballast.', [40, 15, 10, 25, 7, 3], [10, 20], 5, 35, 30, 'Capital preservation', '-12% to -24%', 5, 4, ['Lower expected volatility', 'Still participates in AI buildout', 'Fits capital preservation better'], ['Less upside in an AI rally', 'May feel too conservative'], 'A portfolio where preserving capital matters more than maximizing AI upside.', 'Lower AI sleeve'),
]

export const stressScenarios: StressScenario[] = [
  { id: 'market-correction', name: 'Market correction', shocks: { 'Broad US equity': -10, 'AI buildout sleeve': -15, 'Bonds/fixed income': 0, Cash: 0, 'Alternatives/other': -5 } },
  { id: 'bear-market', name: 'Bear market', shocks: { 'Broad US equity': -20, 'AI buildout sleeve': -35, 'Bonds/fixed income': 2, Cash: 0, 'Alternatives/other': -10 } },
  { id: 'ai-disappointment', name: 'AI disappointment', shocks: { 'Broad US equity': -10, 'AI buildout sleeve': -45, 'Bonds/fixed income': 2, Cash: 0, 'Alternatives/other': -8 } },
  { id: 'ai-mania', name: 'AI mania continuation', shocks: { 'Broad US equity': 12, 'AI buildout sleeve': 35, 'Bonds/fixed income': 0, Cash: 0, 'Alternatives/other': 8 } },
  { id: 'rates-power', name: 'Rates up / power bottleneck', shocks: { 'Broad US equity': -8, 'AI buildout sleeve': -15, 'Bonds/fixed income': -8, Cash: 0, 'Power generation': 8, 'Grid / electrification': 8 } },
  { id: 'big-tech-compression', name: 'Big tech multiple compression', shocks: { 'Broad US equity': -12, 'Big tech / hyperscalers': -30, Semiconductors: -25, 'Bonds/fixed income': 1, Cash: 0 } },
  { id: 'semi-shock', name: 'Semiconductor shock', shocks: { 'Broad US equity': -8, Semiconductors: -35, 'GPUs / accelerators': -35, 'Semiconductor equipment': -35, 'Big tech / hyperscalers': -15, 'Power generation': -5, 'Grid / electrification': -5, 'Data centers / colocation': -5, 'Cooling / thermal management': -5, 'Bonds/fixed income': 1, Cash: 0 } },
]
