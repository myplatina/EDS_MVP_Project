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
