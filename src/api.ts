import type { ApiResponse, AppConfig, AppStatus, ChartData, DashboardData, DividendDashboard, DividendRecord, HoldingRow, PortfolioOutputRow, PriceRefreshResult, RefreshLogs } from './types';

export const APP_VERSION = '0.8.6';

const DEFAULT_API_URL = import.meta.env.VITE_EDS_API_URL || '';
const DEFAULT_API_TOKEN = '';
// GAS 웹앱 보안 인터셉터용 앱 토큰 (환경변수 VITE_EDS_APP_TOKEN으로 배포 시 주입)
const APP_SECURITY_TOKEN = import.meta.env.VITE_EDS_APP_TOKEN || '';

const STORAGE_KEY = 'eds_mvp_config';
const REFRESH_LOG_KEY = 'eds_mvp_refresh_logs_v1';
const CHART_CACHE_PREFIX = 'eds_mvp_chart_cache_v1';
const CHART_CACHE_TTL_MS = 12 * 60 * 60 * 1000;
const AUTO_PRICE_REFRESH_INTERVAL_KEY = 'eds_mvp_auto_price_refresh_interval_v1';
const AUTO_PRICE_REFRESH_MARKET_ONLY_KEY = 'eds_mvp_auto_price_refresh_market_only_v1';

export type ChartPayload = {
  asset_id: string;
  ticker?: string | number;
  market?: string;
  interval: 'D' | 'W' | 'M';
  limit?: number;
};

function makeChartCacheKey(apiUrl: string, payload: ChartPayload): string {
  const base = [apiUrl, payload.asset_id, payload.interval, payload.limit || ''].join('|');
  return `${CHART_CACHE_PREFIX}:${encodeURIComponent(base)}`;
}

function readChartCache(apiUrl: string, payload: ChartPayload): ChartData | null {
  try {
    const raw = localStorage.getItem(makeChartCacheKey(apiUrl, payload));
    if (!raw) return null;
    const parsed = JSON.parse(raw) as { savedAt: number; data: ChartData };
    if (!parsed || !parsed.savedAt || !parsed.data) return null;
    if (Date.now() - parsed.savedAt > CHART_CACHE_TTL_MS) return null;
    if (!parsed.data.items || parsed.data.items.length === 0) return null;
    return parsed.data;
  } catch {
    return null;
  }
}

function writeChartCache(apiUrl: string, payload: ChartPayload, data: ChartData): void {
  try {
    if (!data || !data.items || data.items.length === 0) return;
    localStorage.setItem(
      makeChartCacheKey(apiUrl, payload),
      JSON.stringify({ savedAt: Date.now(), data })
    );
  } catch {
    // localStorage quota 또는 private mode 오류는 무시
  }
}

export function clearChartCaches(): number {
  let count = 0;
  for (let i = localStorage.length - 1; i >= 0; i--) {
    const key = localStorage.key(i);
    if (key && key.startsWith(`${CHART_CACHE_PREFIX}:`)) {
      localStorage.removeItem(key);
      count += 1;
    }
  }
  return count;
}


export function readCachedChartDataForAsset(apiUrl: string, assetId: string, interval: 'D' | 'W' | 'M' = 'D'): ChartData | null {
  try {
    for (let i = localStorage.length - 1; i >= 0; i--) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith(`${CHART_CACHE_PREFIX}:`)) continue;
      const raw = localStorage.getItem(key);
      if (!raw) continue;
      const parsed = JSON.parse(raw) as { savedAt: number; data: ChartData };
      if (!parsed || !parsed.savedAt || !parsed.data) continue;
      if (Date.now() - parsed.savedAt > CHART_CACHE_TTL_MS) continue;
      if (parsed.data.asset_id !== assetId || parsed.data.interval !== interval) continue;
      if (!parsed.data.items || parsed.data.items.length === 0) continue;
      // apiUrl이 다른 배포의 캐시와 섞이지 않도록 키 문자열도 확인한다.
      const decodedKey = decodeURIComponent(key.replace(`${CHART_CACHE_PREFIX}:`, ''));
      if (apiUrl && !decodedKey.startsWith(`${apiUrl}|`)) continue;
      return parsed.data;
    }
  } catch {
    return null;
  }
  return null;
}

export function loadConfig(): AppConfig {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      return {
        apiUrl: parsed.apiUrl || DEFAULT_API_URL,
        token: parsed.token || DEFAULT_API_TOKEN,
      };
    } catch {
      // 저장값 파싱 실패 시 기본값 사용
    }
  }

  return {
    apiUrl: DEFAULT_API_URL,
    token: DEFAULT_API_TOKEN,
  };
}

export function saveConfig(config: AppConfig): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
}

export function clearConfig(): void {
  localStorage.removeItem(STORAGE_KEY);
}

export function clearTokenOnly(): AppConfig {
  const current = loadConfig();
  const next = { apiUrl: current.apiUrl, token: '' };
  saveConfig(next);
  return next;
}

export function clearAllLocalAppData(): number {
  const prefixes = ['eds_mvp_'];
  let count = 0;
  for (let i = localStorage.length - 1; i >= 0; i--) {
    const key = localStorage.key(i);
    if (key && prefixes.some((prefix) => key.startsWith(prefix))) {
      localStorage.removeItem(key);
      count += 1;
    }
  }
  return count;
}


export function getAutoPriceRefreshIntervalMinutes(): number {
  const raw = localStorage.getItem(AUTO_PRICE_REFRESH_INTERVAL_KEY);
  const n = Number(raw || 0);
  return [0, 1, 3, 5, 10].includes(n) ? n : 0;
}

export function setAutoPriceRefreshIntervalMinutes(minutes: number): void {
  const normalized = [0, 1, 3, 5, 10].includes(minutes) ? minutes : 0;
  localStorage.setItem(AUTO_PRICE_REFRESH_INTERVAL_KEY, String(normalized));
}

export function getAutoPriceRefreshMarketOnly(): boolean {
  const raw = localStorage.getItem(AUTO_PRICE_REFRESH_MARKET_ONLY_KEY);
  return raw === null ? true : raw === 'true';
}

export function setAutoPriceRefreshMarketOnly(enabled: boolean): void {
  localStorage.setItem(AUTO_PRICE_REFRESH_MARKET_ONLY_KEY, String(Boolean(enabled)));
}

export function getRefreshLogs(): RefreshLogs {
  try {
    const raw = localStorage.getItem(REFRESH_LOG_KEY);
    if (!raw) return {};
    return JSON.parse(raw) as RefreshLogs;
  } catch {
    return {};
  }
}

function writeRefreshLog(action: string): void {
  const logs = getRefreshLogs();
  const now = new Date().toISOString();
  const key = normalizeRefreshAction(action);
  logs[key] = now;
  localStorage.setItem(REFRESH_LOG_KEY, JSON.stringify(logs));
  window.dispatchEvent(new CustomEvent('eds-mvp-refresh-log', { detail: { action: key, timestamp: now } }));
}

function normalizeRefreshAction(action: string): keyof RefreshLogs {
  if (action === 'refreshChartData') return 'getChartData';
  if (action === 'refreshDividends') return 'getDividendDashboard';
  return action as keyof RefreshLogs;
}

function explainApiError(action: string, error: string, message?: string): string {
  const raw = [error, message].filter(Boolean).join(' · ');
  if (error === 'missing_token') return 'API token이 누락되었습니다. 설정 탭에서 token을 입력하세요.';
  if (error === 'invalid_token') return 'API token이 올바르지 않습니다. App_Settings의 api_token과 앱 설정값을 확인하세요.';
  if (error === 'unknown_action') return `Apps Script에 ${action} action이 아직 배포되지 않았습니다. Apps Script 저장 후 웹앱을 새 버전으로 배포하세요.`;
  if (error === 'missing_action') return 'API action 파라미터가 누락되었습니다. PWA/API 버전 불일치 가능성이 있습니다.';
  if (error === 'exception') return `Apps Script 실행 오류: ${message || '상세 메시지 없음'}`;
  return raw || `${action} failed`;
}

async function parseJsonResponse<T>(res: Response, action: string): Promise<ApiResponse<T>> {
  const text = await res.text();
  let json: ApiResponse<T>;
  try {
    json = JSON.parse(text) as ApiResponse<T>;
  } catch {
    throw new Error(`API 응답을 JSON으로 해석하지 못했습니다. Apps Script URL/배포 권한을 확인하세요. HTTP=${res.status}, body=${text.slice(0, 160)}`);
  }

  if (!res.ok) {
    throw new Error(`API HTTP 오류: ${res.status}. Apps Script 배포 권한 또는 URL을 확인하세요.`);
  }

  if (!json.ok) {
    throw new Error(explainApiError(action, json.error || '', json.message));
  }

  return json;
}

export class EdsApi {
  private config: AppConfig;

  constructor(config: AppConfig) {
    this.config = config;
  }

  updateConfig(config: AppConfig) {
    this.config = config;
  }

  getConfig() {
    return this.config;
  }

  private requireUrl() {
    if (!this.config.apiUrl) {
      throw new Error('Apps Script 웹앱 URL이 설정되지 않았습니다. 설정 탭에서 URL을 입력하세요.');
    }
  }

  private requireToken(action: string) {
    if (action !== 'ping' && !this.config.token) {
      throw new Error('API token이 설정되지 않았습니다. 설정 탭에서 token을 입력하세요.');
    }
  }

  private makeUrl(action: string, params: Record<string, string> = {}) {
    this.requireUrl();
    this.requireToken(action);

    const url = new URL(this.config.apiUrl);
    url.searchParams.set('action', action);
    if (action !== 'ping') url.searchParams.set('token', this.config.token);
    // GAS 보안 인터셉터용 앱 토큰 (헤더는 GAS에서 지원 안 되므로 URL 파라미터로 전달)
    if (APP_SECURITY_TOKEN) url.searchParams.set('x_eds_app_token', APP_SECURITY_TOKEN);
    Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
    return url.toString();
  }

  async get<T>(action: string, params: Record<string, string> = {}): Promise<T> {
    try {
      const res = await fetch(this.makeUrl(action, params), { method: 'GET' });
      const json = await parseJsonResponse<T>(res, action);
      writeRefreshLog(action);
      return json.data as T;
    } catch (e) {
      if (e instanceof TypeError) {
        throw new Error('API 호출 실패: 네트워크/CORS/Apps Script 배포 권한을 확인하세요. 액세스 권한은 “모든 사용자”여야 합니다.');
      }
      throw e;
    }
  }

  async post<T>(action: string, payload: unknown = {}): Promise<T> {
    this.requireUrl();
    this.requireToken(action);

    const body = JSON.stringify({
      action,
      token: this.config.token,
      // GAS 보안 인터셉터용 앱 토큰 (POST body에 포함)
      ...(APP_SECURITY_TOKEN ? { x_eds_app_token: APP_SECURITY_TOKEN } : {}),
      payload,
    });

    try {
      // text/plain은 Apps Script CORS preflight를 줄이기 위한 선택.
      const res = await fetch(this.config.apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body,
      });

      const json = await parseJsonResponse<T>(res, action);
      writeRefreshLog(action);
      return json.data as T;
    } catch (e) {
      if (e instanceof TypeError) {
        throw new Error('API 호출 실패: 네트워크/CORS/Apps Script 배포 권한을 확인하세요. 액세스 권한은 “모든 사용자”여야 합니다.');
      }
      throw e;
    }
  }

  ping() {
    return this.get<{ app: string; timestamp: string }>('ping');
  }

  getAppStatus() {
    return this.post<AppStatus>('getAppStatus', { pwa_version: APP_VERSION });
  }

  getDashboard() {
    return this.get<DashboardData>('getDashboard');
  }

  getHoldings() {
    return this.get<HoldingRow[]>('getHoldings');
  }

  getPortfolioOutput() {
    return this.get<PortfolioOutputRow[]>('getPortfolioOutput');
  }

  refreshOutput() {
    return this.get<{ message: string; output_count: number }>('refreshOutput');
  }

  refreshKrxPrices() {
    return this.post<PriceRefreshResult>('refreshKrxPrices', {});
  }

  refreshKrxPricesToMainSheet() {
    return this.post<PriceRefreshResult>('refreshKrxPricesToMainSheet', {});
  }

  refreshKrxDailyCharts() {
    return this.post<{ target_count: number; row_count: number; success_count: number; error_count: number; started_at?: string; finished_at?: string }>('refreshKrxDailyCharts', {});
  }

  previewSync() {
    return this.get<{ message: string; preview_summary: Record<string, unknown> }>('previewSync');
  }

  syncToMain(skipValidation = true) {
    // skipValidation=true(기본): 고속 숏컷 모드 — 무거운 인덱싱/검증 생략, Bulk Write 직행
    // skipValidation=false: 안전 모드 — 전체 인덱싱 + 정합성 검증 후 반영
    return this.post<{ message: string; mode: 'fast' | 'safe'; sync_result_summary: Record<string, unknown> }>(
      'syncToMain',
      { skipValidation }
    );
  }

  async getChartData(payload: ChartPayload) {
    const cached = readChartCache(this.config.apiUrl, payload);
    if (cached) return cached;

    const data = await this.post<ChartData>('getChartData', payload);
    writeChartCache(this.config.apiUrl, payload, data);
    return data;
  }

  async fetchSingleChartData(payload: Omit<ChartPayload, 'interval'>) {
    const data = await this.post<{ D: ChartData; W: ChartData; M: ChartData; asset_id: string }>('fetchSingleChartData', payload);
    if (data.D) writeChartCache(this.config.apiUrl, { ...payload, interval: 'D' }, data.D);
    if (data.W) writeChartCache(this.config.apiUrl, { ...payload, interval: 'W' }, data.W);
    if (data.M) writeChartCache(this.config.apiUrl, { ...payload, interval: 'M' }, data.M);
    return data;
  }

  async refreshChartData(payload: ChartPayload) {
    const data = await this.post<ChartData>('refreshChartData', { ...payload, force: true });
    writeChartCache(this.config.apiUrl, payload, data);
    return data;
  }

  updateHolding(payload: Partial<HoldingRow> & { holding_id: string }) {
    return this.post<{ updated: { row: number; changed_count: number }; holding_id: string }>('updateHolding', payload);
  }

  getDividendDashboard() {
    return this.post<DividendDashboard>('getDividendDashboard', {});
  }

  refreshDividends() {
    return this.post<{ imported_count: number; skipped_count: number; error_count?: number; source_sheet: string; fetched_at: string }>('refreshDividends', {});
  }

  addDividend(payload: Partial<DividendRecord>) {
    return this.post<{ dividend_id: string; row: number }>('addDividend', payload);
  }

  updateDividend(payload: Partial<DividendRecord> & { dividend_id: string }) {
    return this.post<{ dividend_id: string; row: number; changed_count: number }>('updateDividend', payload);
  }

  disableDividend(payload: { dividend_id: string }) {
    return this.post<{ dividend_id: string; row: number }>('disableDividend', payload);
  }
}
