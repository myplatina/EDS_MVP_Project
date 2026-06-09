import type { PortfolioOutputRow } from './types';

export function formatKRW(value: number | string | undefined | null): string {
  const n = Number(value ?? 0);
  return new Intl.NumberFormat('ko-KR', {
    style: 'currency',
    currency: 'KRW',
    maximumFractionDigits: 0,
  }).format(n);
}

export function formatSignedKRW(value: number | string | undefined | null): string {
  const n = Number(value ?? 0);
  if (n > 0) return `+${formatKRW(n)}`;
  if (n < 0) return `-${formatKRW(Math.abs(n))}`;
  return formatKRW(0);
}

export function formatNumber(value: number | string | undefined | null, digits = 0): string {
  const n = Number(value ?? 0);
  return new Intl.NumberFormat('ko-KR', {
    maximumFractionDigits: digits,
  }).format(n);
}

export function formatPercent(value: number | string | undefined | null, digits = 2): string {
  const n = Number(value ?? 0);
  return `${(n * 100).toFixed(digits)}%`;
}

export function formatSignedPercent(value: number | string | undefined | null, digits = 2): string {
  const n = Number(value ?? 0);
  const abs = Math.abs(n);
  if (n > 0) return `+${(abs * 100).toFixed(digits)}%`;
  if (n < 0) return `-${(abs * 100).toFixed(digits)}%`;
  return `${(0).toFixed(digits)}%`;
}

export function signedClass(value: number | string | undefined | null): string {
  const n = Number(value ?? 0);
  if (n > 0) return 'positive';
  if (n < 0) return 'negative';
  return 'neutral';
}

export function returnClass(value: number | string | undefined | null): string {
  const n = Number(value ?? 0);
  if (n > 0) return 'return-positive';
  if (n < 0) return 'return-negative';
  return 'return-neutral';
}

export function buildHoldingKey(row: PortfolioOutputRow): string {
  return `${row.account_id}__${row.asset_id}`;
}

export function sortByValuationDesc<T extends { valuation_amount: number }>(rows: T[]): T[] {
  return rows.slice().sort((a, b) => Number(b.valuation_amount || 0) - Number(a.valuation_amount || 0));
}

// ============================================================
// 위험 모니터링 엔진 (Risk Monitoring Engine)
// ============================================================

export type RiskFlag =
  | 'DROP_WARNING'          // 총수익률 < -20%
  | 'CONCENTRATION_WARNING' // 전체 투자비중 > 15%
  | 'WEIGHT_DEVIATION';     // 계좌비중-목표비중 절대 괴리율 >= 30%

/**
 * 각 종목별 위험 플래그 Set을 계산하여 Map으로 반환
 * key: `${account_id}__${asset_id}` (buildHoldingKey와 동일)
 */
export function computeRiskFlags(items: PortfolioOutputRow[]): Map<string, Set<RiskFlag>> {
  const result = new Map<string, Set<RiskFlag>>();

  for (const item of items) {
    const key = buildHoldingKey(item);
    const flags = new Set<RiskFlag>();

    const profitRate = Number(item.profit_rate ?? 0);
    const totalWeight = Number(item.total_weight ?? 0);
    const accountWeight = Number(item.account_weight ?? 0);
    const targetWeight = Number(item.target_weight_account ?? 0);

    // Drop Warning: 총수익률 < -20%
    if (profitRate < -0.20) {
      flags.add('DROP_WARNING');
    }

    // Concentration Warning: 전체 투자비중 > 15%
    if (totalWeight > 0.15) {
      flags.add('CONCENTRATION_WARNING');
    }

    // Weight Deviation Warning: |계좌비중 - 목표비중| / 목표비중 >= 30%
    if (targetWeight > 0) {
      const deviation = Math.abs(accountWeight - targetWeight) / targetWeight;
      if (deviation >= 0.30) {
        flags.add('WEIGHT_DEVIATION');
      }
    }

    if (flags.size > 0) result.set(key, flags);
  }

  return result;
}

export type DeficitCandidate = {
  item: PortfolioOutputRow;
  gapAmount: number; // 양수 = 부족 (추가 매수 필요)
  deviationRate: number; // (계좌비중 - 목표비중) / 목표비중 — 음수 = 부족
  required_shares: number; // 부족 금액(Gap) ÷ 원화 현재가(price)
};

/**
 * 비중 괴리율이 -30% 이하(= 심각한 부족)인 종목을 추출하여
 * 포트폴리오 최상단 Action Card에 표시할 데이터 반환
 */
export function computeDeficitCandidates(items: PortfolioOutputRow[]): DeficitCandidate[] {
  const candidates: DeficitCandidate[] = [];

  for (const item of items) {
    const accountWeight = Number(item.account_weight ?? 0);
    const targetWeight = Number(item.target_weight_account ?? 0);
    if (targetWeight <= 0) continue;

    const deviationRate = (accountWeight - targetWeight) / targetWeight;
    if (deviationRate > -0.30) continue;

    // target_gap_amount: target amount - current valuation (양수 = 부족)
    const gapAmount = Number(item.target_gap_amount ?? 0);
    if (gapAmount <= 0) continue;

    const price = Number(item.price ?? 0);
    const required_shares = price > 0 ? Math.floor(gapAmount / price) : 0;

    candidates.push({ item, gapAmount, deviationRate, required_shares });
  }

  // 부족 금액 큰 순서 정렬
  return candidates.sort((a, b) => b.gapAmount - a.gapAmount);
}
