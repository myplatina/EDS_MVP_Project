import { useEffect, useMemo, useRef, useState } from 'react';
import { APP_VERSION, EdsApi, clearAllLocalAppData, clearChartCaches, clearConfig, clearTokenOnly, getAutoPriceRefreshIntervalMinutes, getAutoPriceRefreshMarketOnly, getRefreshLogs, loadConfig, readCachedChartDataForAsset, saveConfig, setAutoPriceRefreshIntervalMinutes, setAutoPriceRefreshMarketOnly } from './api';
import type { AccountSummary, AppConfig, AppStatus, ChartData, ChartPoint, DashboardData, DividendDashboard, HoldingRow, PortfolioItem, RefreshLogs } from './types';
import { buildHoldingKey, computeDeficitCandidates, computeRiskFlags, formatKRW, formatNumber, formatPercent, formatSignedKRW, formatSignedPercent, returnClass, signedClass, sortByValuationDesc } from './utils';
import type { DeficitCandidate, RiskFlag } from './utils';

type Tab = 'home' | 'portfolio' | 'rebalance' | 'history' | 'settings';

type EditState = {
  item: PortfolioItem;
  quantity: string;
  avg_price: string;
  target_weight_account: string;
};

type DetailState =
  | { kind: 'account'; account: AccountSummary }
  | { kind: 'holding'; item: PortfolioItem };

type BeforeInstallPromptEvent = Event & {
  prompt: () => Promise<void>;
  userChoice: Promise<{ outcome: 'accepted' | 'dismissed'; platform: string }>;
};

const ACCOUNT_FILTER_KEY = 'eds_mvp_account_filter';
const REBALANCE_CASH_KEY = 'eds_mvp_rebalance_cash_v1';
const REBALANCE_MODE_KEY = 'eds_mvp_rebalance_mode_v1';
const SYNC_FAST_MODE_KEY = 'eds_mvp_sync_fast_mode_v1';

function targetPositionGapAmount(item: PortfolioItem): number {
  // App_Output.target_gap_amount = target amount - current valuation.
  // UI에서는 '현재 보유가 목표보다 많은가/적은가' 기준으로 보여야 하므로 부호를 반전한다.
  return -Number(item.target_gap_amount || 0);
}

function App() {
  const [tab, setTab] = useState<Tab>('home');
  const [config, setConfig] = useState<AppConfig>(() => loadConfig());
  const [api] = useState(() => new EdsApi(config));
  const [dashboard, setDashboard] = useState<DashboardData | null>(null);
  const [holdings, setHoldings] = useState<HoldingRow[]>([]);
  const [dividends, setDividends] = useState<DividendDashboard | null>(null);
  const [accountFilter, setAccountFilterState] = useState(() => localStorage.getItem(ACCOUNT_FILTER_KEY) || 'ALL');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [refreshLogs, setRefreshLogs] = useState<RefreshLogs>(() => getRefreshLogs());
  const [appStatus, setAppStatus] = useState<AppStatus | null>(null);
  const [edit, setEdit] = useState<EditState | null>(null);
  const [detail, setDetail] = useState<DetailState | null>(null);
  const [installPrompt, setInstallPrompt] = useState<BeforeInstallPromptEvent | null>(null);
  const [isStandalone, setIsStandalone] = useState(false);
  const [autoPriceRefreshMinutes, setAutoPriceRefreshMinutesState] = useState(() => getAutoPriceRefreshIntervalMinutes());
  const [autoPriceMarketOnly, setAutoPriceMarketOnlyState] = useState(() => getAutoPriceRefreshMarketOnly());
  const [autoPriceStatus, setAutoPriceStatus] = useState('자동 갱신 꺼짐');
  const autoPriceRunningRef = useRef(false);
  // 원장 동기화 모드: true=고속(기본, skipValidation), false=안전(전체 검증)
  const [syncFastMode, setSyncFastModeState] = useState(() => {
    const saved = localStorage.getItem(SYNC_FAST_MODE_KEY);
    return saved === null ? true : saved === 'true';
  });

  function setSyncFastMode(next: boolean) {
    setSyncFastModeState(next);
    localStorage.setItem(SYNC_FAST_MODE_KEY, String(next));
  }

  useEffect(() => {
    api.updateConfig(config);
  }, [api, config]);

  useEffect(() => {
    const logHandler = () => setRefreshLogs(getRefreshLogs());
    window.addEventListener('eds-mvp-refresh-log', logHandler);
    return () => window.removeEventListener('eds-mvp-refresh-log', logHandler);
  }, []);

  useEffect(() => {
    const standalone = window.matchMedia('(display-mode: standalone)').matches || Boolean((window.navigator as Navigator & { standalone?: boolean }).standalone);
    setIsStandalone(standalone);

    const handler = (event: Event) => {
      event.preventDefault();
      setInstallPrompt(event as BeforeInstallPromptEvent);
    };

    window.addEventListener('beforeinstallprompt', handler);
    return () => window.removeEventListener('beforeinstallprompt', handler);
  }, []);

  function setAccountFilter(next: string) {
    setAccountFilterState(next);
    localStorage.setItem(ACCOUNT_FILTER_KEY, next);
  }

  function setAutoPriceRefreshMinutes(next: number) {
    setAutoPriceRefreshIntervalMinutes(next);
    setAutoPriceRefreshMinutesState(next);
    setAutoPriceStatus(next > 0 ? `자동 갱신 ${next}분 주기로 설정` : '자동 갱신 꺼짐');
  }

  function setAutoPriceMarketOnly(next: boolean) {
    setAutoPriceRefreshMarketOnly(next);
    setAutoPriceMarketOnlyState(next);
    setAutoPriceStatus(next ? '시장 시간에만 자동 갱신' : '장중 여부와 무관하게 자동 갱신');
  }

  function isMarketRefreshWindowNow(): boolean {
    const now = new Date();
    const day = now.getDay();
    const minutes = now.getHours() * 60 + now.getMinutes();

    // KST 기준 국내 정규장: 월~금 09:00~15:40
    const isKrxWindow = day >= 1 && day <= 5 && minutes >= 9 * 60 && minutes <= 15 * 60 + 40;

    // KST 기준 미국장 대략 범위. DST/휴장일은 엄밀 반영하지 않고 자동 갱신 편의용으로 넓게 잡는다.
    // 월~금 22:00~23:59, 화~토 00:00~06:30
    const isUsEveningWindow = day >= 1 && day <= 5 && minutes >= 22 * 60;
    const isUsEarlyMorningWindow = day >= 2 && day <= 6 && minutes <= 6 * 60 + 30;

    return isKrxWindow || isUsEveningWindow || isUsEarlyMorningWindow;
  }

  async function loadData() {
    setLoading(true);
    setError('');
    setMessage('');
    try {
      const [nextDashboard, nextHoldings, nextDividends] = await Promise.all([
        api.getDashboard(),
        api.getHoldings(),
        api.getDividendDashboard().catch(() => null),
      ]);
      setDashboard(nextDashboard);
      setHoldings(nextHoldings);
      setDividends(nextDividends);
      setRefreshLogs(getRefreshLogs());
      setMessage('데이터 새로고침 완료');
    } catch (e) {
      setError(formatErrorForDisplay(e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    if (config.apiUrl && config.token) {
      loadData();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!config.apiUrl || !config.token || autoPriceRefreshMinutes <= 0) {
      setAutoPriceStatus('자동 갱신 꺼짐');
      return;
    }

    setAutoPriceStatus(`자동 갱신 대기 중 · ${autoPriceRefreshMinutes}분 주기`);

    const tick = async () => {
      if (autoPriceRunningRef.current) return;
      if (document.hidden) {
        setAutoPriceStatus('앱이 백그라운드라 자동 갱신 대기');
        return;
      }
      if (autoPriceMarketOnly && !isMarketRefreshWindowNow()) {
        setAutoPriceStatus('국내/미국 시장 시간이 아니라 자동 갱신 대기');
        return;
      }

      autoPriceRunningRef.current = true;
      setAutoPriceStatus('현재가 자동 갱신 중...');
      try {
        const result = await api.refreshKrxPrices();
        const [nextDashboard, nextHoldings] = await Promise.all([api.getDashboard(), api.getHoldings()]);
        setDashboard(nextDashboard);
        setHoldings(nextHoldings);
        setRefreshLogs(getRefreshLogs());
        setAutoPriceStatus(`자동 갱신 완료 · 성공 ${result.success_count} · 실패 ${result.error_count}`);
      } catch (e) {
        setAutoPriceStatus(`자동 갱신 실패 · ${formatErrorForDisplay(e)}`);
      } finally {
        autoPriceRunningRef.current = false;
      }
    };

    const timer = window.setInterval(tick, autoPriceRefreshMinutes * 60 * 1000);
    return () => window.clearInterval(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [api, config.apiUrl, config.token, autoPriceRefreshMinutes, autoPriceMarketOnly]);

  const holdingByKey = useMemo(() => {
    const map = new Map<string, HoldingRow>();
    holdings.forEach((h) => map.set(`${h.account_id}__${h.asset_id}`, h));
    return map;
  }, [holdings]);

  const portfolioItems: PortfolioItem[] = useMemo(() => {
    const rows = dashboard?.output ?? [];
    return rows.map((row) => ({
      ...row,
      holding_id: holdingByKey.get(buildHoldingKey(row))?.holding_id,
    }));
  }, [dashboard, holdingByKey]);

  const filteredItems = useMemo(() => {
    const rows = accountFilter === 'ALL'
      ? portfolioItems
      : portfolioItems.filter((row) => row.account_id === accountFilter);
    return sortByValuationDesc(rows);
  }, [accountFilter, portfolioItems]);

  const visibleSummary = useMemo(() => {
    if (!dashboard) return null;
    if (accountFilter === 'ALL') return dashboard.summary;
    const account = dashboard.accounts.find((a) => a.account_id === accountFilter);
    if (!account) return dashboard.summary;
    return {
      total_valuation: account.valuation_amount,
      total_invested: account.invested_amount,
      total_profit: account.profit_amount,
      total_profit_rate: account.profit_rate,
      account_count: 1,
      holding_count: account.holding_count,
    };
  }, [accountFilter, dashboard]);

  function updateConfig(next: AppConfig) {
    const trimmed = {
      apiUrl: next.apiUrl.trim(),
      token: next.token.trim(),
    };
    setConfig(trimmed);
    saveConfig(trimmed);
    api.updateConfig(trimmed);
    setMessage('설정 저장 완료');
    setError('');
  }

  function resetConfig() {
    const confirmed = window.confirm('저장된 API URL과 token을 이 브라우저에서 삭제합니다. 계속할까요?');
    if (!confirmed) return;
    clearConfig();
    const blank = { apiUrl: '', token: '' };
    setConfig(blank);
    api.updateConfig(blank);
    setDashboard(null);
    setHoldings([]);
    setMessage('설정 초기화 완료');
    setError('');
  }

  function openEdit(item: PortfolioItem) {
    if (!item.holding_id) {
      setError('holding_id를 찾지 못했습니다. getHoldings 결과를 확인하세요.');
      return;
    }
    setEdit({
      item,
      quantity: String(item.quantity ?? ''),
      avg_price: String(item.avg_price ?? ''),
      target_weight_account: String(Number(item.target_weight_account ?? 0) * 100),
    });
  }

  async function saveEdit() {
    if (!edit || !edit.item.holding_id) return;
    setLoading(true);
    setError('');
    setMessage('');
    try {
      await api.updateHolding({
        holding_id: edit.item.holding_id,
        quantity: Number(edit.quantity),
        avg_price: Number(edit.avg_price),
        target_weight_account: Number(edit.target_weight_account) / 100,
      });
      setEdit(null);
      await loadData();
      setMessage('종목 수정 완료. 원장 반영 전 previewSync를 확인하세요.');
    } catch (e) {
      setError(formatErrorForDisplay(e));
    } finally {
      setLoading(false);
    }
  }

  async function runAction(label: string, fn: () => Promise<unknown>) {
    setLoading(true);
    setError('');
    setMessage('');
    try {
      const result = await fn();
      setMessage(`${label} 완료: ${summarizeResult(result)}`);
      await loadData();
    } catch (e) {
      setError(formatErrorForDisplay(e));
    } finally {
      setLoading(false);
    }
  }

  async function installPwa() {
    if (!installPrompt) {
      setMessage('현재 브라우저에서는 설치 프롬프트가 준비되지 않았습니다. 크롬 메뉴의 “홈 화면에 추가”를 사용하세요.');
      return;
    }
    await installPrompt.prompt();
    const choice = await installPrompt.userChoice;
    setInstallPrompt(null);
    setMessage(choice.outcome === 'accepted' ? 'PWA 설치가 시작되었습니다.' : 'PWA 설치가 취소되었습니다.');
  }

  function handleSyncToMain() {
    const modeLabel = syncFastMode ? '⚡ 고속 모드 (skipValidation)' : '🛡️ 안전 모드 (전체 검증)';
    const typed = window.prompt(
      `원장 반영(syncToMain) — 현재: ${modeLabel}\n\n기존 2. 종목현황 시트를 실제 수정합니다. previewSync 확인 후 계속하려면 SYNC를 입력하세요.`
    );
    if (typed !== 'SYNC') {
      setMessage('원장 반영 취소');
      return;
    }
    // syncFastMode=true → skipValidation: true (고속 숏컷 모드)
    // syncFastMode=false → skipValidation: false (안전 모드, 전체 검증)
    runAction('syncToMain', () => api.syncToMain(syncFastMode));
  }

  function handleRefreshKrxPrices() {
    const confirmed = window.confirm('KIS Open API로 국내/미국 보유종목 현재가를 갱신하고 App_Output을 재계산합니다. 원본 2. 종목현황은 수정하지 않습니다. 계속할까요?');
    if (!confirmed) {
      setMessage('현재가 갱신 취소');
      return;
    }
    runAction('refreshKrxPrices', () => api.refreshKrxPrices());
  }

  function handleRefreshKrxPricesToMainSheet() {
    const confirmed = window.confirm('KIS 현재가를 App_Prices/App_Output에 반영하고, 원본 2. 종목현황의 국내 종목 K/M열과 미국 종목 L/M열을 값으로 갱신합니다. 기존 GOOGLEFINANCE/네이버 수식이 값으로 대체될 수 있습니다. 계속할까요?');
    if (!confirmed) {
      setMessage('원본 시트 가격 반영 취소');
      return;
    }
    runAction('refreshKrxPricesToMainSheet', () => api.refreshKrxPricesToMainSheet());
  }


  function handleRefreshKrxDailyCharts() {
    const confirmed = window.confirm('KIS Open API로 국내 KRX 보유종목의 일봉 차트 캐시를 갱신합니다. 실행 시간이 길 수 있으며 원본 2. 종목현황은 수정하지 않습니다. 계속할까요?');
    if (!confirmed) {
      setMessage('국내 일봉 차트 갱신 취소');
      return;
    }
    runAction('refreshKrxDailyCharts', () => api.refreshKrxDailyCharts());
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <div className="eyebrow">Google Sheets Portfolio Terminal · v{APP_VERSION} · © 삼평동불나방 · myplatina@gmail.com</div>
          <h1>ED's MVP</h1>
        </div>
        <button className="icon-button" onClick={loadData} disabled={loading}>↻</button>
      </header>

      {message && <div className="notice success">{message}</div>}
      {error && <div className="notice error">{error}</div>}
      {loading && <div className="notice loading">처리 중...</div>}

      <main className="content">
        {tab === 'home' && dashboard && visibleSummary && (
          <HomePage
            dashboard={dashboard}
            summary={visibleSummary}
            accountFilter={accountFilter}
            setAccountFilter={setAccountFilter}
            onAccountDetail={(account) => setDetail({ kind: 'account', account })}
            onHoldingDetail={(item) => setDetail({ kind: 'holding', item })}
            onRefreshPrices={handleRefreshKrxPrices}
            onRefreshDailyCharts={handleRefreshKrxDailyCharts}
            refreshLogs={refreshLogs}
            loading={loading}
          />
        )}

        {tab === 'portfolio' && dashboard && (
          <PortfolioPage
            dashboard={dashboard}
            items={filteredItems}
            accountFilter={accountFilter}
            setAccountFilter={setAccountFilter}
            onEdit={openEdit}
            onDetail={(item) => setDetail({ kind: 'holding', item })}
          />
        )}

        {tab === 'rebalance' && dashboard && (
          <RebalancePage
            dashboard={dashboard}
            items={portfolioItems}
            accountFilter={accountFilter}
            setAccountFilter={setAccountFilter}
            onHoldingDetail={(item) => setDetail({ kind: 'holding', item })}
            apiUrl={config.apiUrl}
          />
        )}

        {tab === 'history' && dashboard && (
          <DividendPage
            dashboard={dashboard}
            dividends={dividends}
            api={api}
            onReload={loadData}
            onHoldingDetail={(item) => setDetail({ kind: 'holding', item })}
          />
        )}

        {tab === 'history' && !dashboard && <HistoryPage />}

        {tab === 'settings' && (
          <SettingsPage
            config={config}
            updateConfig={updateConfig}
            resetConfig={resetConfig}
            onPing={() => runAction('ping', () => api.ping())}
            onRefreshOutput={() => runAction('refreshOutput', () => api.refreshOutput())}
            onRefreshPrices={handleRefreshKrxPrices}
            onRefreshPricesToMainSheet={handleRefreshKrxPricesToMainSheet}
            onRefreshKrxDailyCharts={handleRefreshKrxDailyCharts}
            autoPriceRefreshMinutes={autoPriceRefreshMinutes}
            onAutoPriceRefreshMinutesChange={setAutoPriceRefreshMinutes}
            autoPriceMarketOnly={autoPriceMarketOnly}
            onAutoPriceMarketOnlyChange={setAutoPriceMarketOnly}
            autoPriceStatus={autoPriceStatus}
            onPreviewSync={() => runAction('previewSync', () => api.previewSync())}
            onSyncToMain={handleSyncToMain}
            syncFastMode={syncFastMode}
            onSyncFastModeChange={setSyncFastMode}
            onInstall={installPwa}
            installAvailable={Boolean(installPrompt)}
            isStandalone={isStandalone}
            appStatus={appStatus}
            refreshLogs={refreshLogs}
            onFetchStatus={() => runAction('getAppStatus', async () => {
              const status = await api.getAppStatus();
              setAppStatus(status);
              return status;
            })}
            onClearChartCache={() => {
              const count = clearChartCaches();
              setMessage(`차트 캐시 삭제 완료: ${count}건`);
            }}
            onClearTokenOnly={() => {
              const next = clearTokenOnly();
              setConfig(next);
              api.updateConfig(next);
              setMessage('API token만 삭제 완료');
            }}
            onClearAllLocalData={() => {
              const confirmed = window.confirm('API 설정, 계좌 필터, 차트 캐시, 새로고침 로그 등 이 브라우저의 ED MVP 로컬 설정을 모두 삭제합니다. 계속할까요?');
              if (!confirmed) return;
              const count = clearAllLocalAppData();
              const blank = { apiUrl: '', token: '' };
              setConfig(blank);
              api.updateConfig(blank);
              setDashboard(null);
              setHoldings([]);
              setDividends(null);
              setRefreshLogs({});
              setMessage(`로컬 앱 데이터 삭제 완료: ${count}건`);
            }}
          />
        )}

        {!dashboard && tab !== 'settings' && (
          <EmptyState onGoSettings={() => setTab('settings')} onRefresh={loadData} />
        )}
      </main>

      <nav className="bottom-nav">
        <button className={tab === 'home' ? 'active' : ''} onClick={() => setTab('home')}>홈</button>
        <button className={tab === 'portfolio' ? 'active' : ''} onClick={() => setTab('portfolio')}>포트폴리오</button>
        <button className={tab === 'rebalance' ? 'active' : ''} onClick={() => setTab('rebalance')}>리밸런싱</button>
        <button className={tab === 'history' ? 'active' : ''} onClick={() => setTab('history')}>내역</button>
        <button className={tab === 'settings' ? 'active' : ''} onClick={() => setTab('settings')}>설정</button>
      </nav>

      {detail && dashboard && (
        <DetailModal
          detail={detail}
          dashboard={dashboard}
          dividends={dividends}
          api={api}
          onClose={() => setDetail(null)}
          onEditHolding={(item) => {
            setDetail(null);
            openEdit(item);
          }}
        />
      )}

      {edit && (
        <div className="modal-backdrop" onClick={() => setEdit(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <h2>{edit.item.asset_name}</h2>
            <p className="muted">{edit.item.account_name} · {edit.item.ticker}</p>
            <label>
              수량
              <input value={edit.quantity} onChange={(e) => setEdit({ ...edit, quantity: e.target.value })} inputMode="decimal" />
            </label>
            <label>
              평단가 KRW
              <input value={edit.avg_price} onChange={(e) => setEdit({ ...edit, avg_price: e.target.value })} inputMode="decimal" />
            </label>
            <label>
              목표비중 %
              <input value={edit.target_weight_account} onChange={(e) => setEdit({ ...edit, target_weight_account: e.target.value })} inputMode="decimal" />
            </label>
            <div className="warning-box">미국 ETF의 평단가는 앱 표시용 KRW 환산값이다. USD 원장 평단에는 자동 반영하지 않는다.</div>
            <div className="modal-actions">
              <button className="ghost" onClick={() => setEdit(null)}>취소</button>
              <button onClick={saveEdit} disabled={loading}>저장</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function AccountFilter({ dashboard, value, onChange }: { dashboard: DashboardData; value: string; onChange: (v: string) => void }) {
  return (
    <select className="select" value={value} onChange={(e) => onChange(e.target.value)}>
      <option value="ALL">전체 계좌</option>
      {dashboard.accounts.map((account) => (
        <option key={account.account_id} value={account.account_id}>
          {account.account_name} · {account.broker}
        </option>
      ))}
    </select>
  );
}

function HomePage({ dashboard, summary, accountFilter, setAccountFilter, onAccountDetail, onHoldingDetail, onRefreshPrices, onRefreshDailyCharts, refreshLogs, loading }: {
  dashboard: DashboardData;
  summary: DashboardData['summary'];
  accountFilter: string;
  setAccountFilter: (v: string) => void;
  onAccountDetail: (account: AccountSummary) => void;
  onHoldingDetail: (item: PortfolioItem) => void;
  onRefreshPrices: () => void;
  onRefreshDailyCharts: () => void;
  refreshLogs: RefreshLogs;
  loading: boolean;
}) {
  // 위험 플래그 + 부족 종목 연산
  const riskFlags = useMemo(() => computeRiskFlags(dashboard.output), [dashboard.output]);
  const deficitCandidates = useMemo(() => computeDeficitCandidates(dashboard.output), [dashboard.output]);

  return (
    <section className="page-stack">
      {/* Quick Action Bar */}
      <QuickActionBar
        onRefreshPrices={onRefreshPrices}
        onRefreshDailyCharts={onRefreshDailyCharts}
        refreshLogs={refreshLogs}
        loading={loading}
      />

      {/* Deficit Action Cards */}
      {deficitCandidates.length > 0 && (
        <section className="card deficit-action-card">
          <div className="section-title">📌 추가 매수 후보 (비중 부족 -30% 초과)</div>
          {deficitCandidates.slice(0, 3).map((c) => (
            <DeficitCard key={`${c.item.account_id}__${c.item.asset_id}`} candidate={c} onDetail={() => onHoldingDetail(c.item)} />
          ))}
        </section>
      )}

      <AccountFilter dashboard={dashboard} value={accountFilter} onChange={setAccountFilter} />

      <section className="hero-card">
        <div className="muted">종 평가액</div>
        <div className="hero-value">{formatKRW(summary.total_valuation)}</div>
        <div className="summary-grid">
          <Metric label="투자원금" value={formatKRW(summary.total_invested)} />
          <Metric label="평가손익" value={formatKRW(summary.total_profit)} className={signedClass(summary.total_profit)} />
          <Metric label="수익률" value={formatPercent(summary.total_profit_rate)} className={returnClass(summary.total_profit_rate)} />
          <Metric label="보유종목" value={`${summary.holding_count}개`} />
        </div>
      </section>

      <section className="card">
        <div className="section-title">계좌별 비중</div>
        <div className="bar-list">
          {dashboard.accounts.map((account) => (
            <button key={account.account_id} className="bar-row account-row-button" onClick={() => onAccountDetail(account)}>
              <div className="bar-label">
                <span>{account.account_name}</span>
                <strong>{formatPercent(account.weight)}</strong>
              </div>
              <div className="bar-bg"><div className="bar-fill" style={{ width: `${Math.max(account.weight * 100, 2)}%` }} /></div>
              <div className="account-subline">
                <span>평가액 {formatKRW(account.valuation_amount)}</span>
                <span className={returnClass(account.profit_rate)}>수익률 {formatPercent(account.profit_rate)}</span>
              </div>
              <div className="tap-hint">탭해서 계좌 상세 보기</div>
            </button>
          ))}
        </div>
      </section>

      <section className="card">
        <div className="section-title">상위 보유 종목</div>
        <HoldingList rows={dashboard.top_holdings} compact onDetail={onHoldingDetail} riskFlags={riskFlags} />
      </section>
    </section>
  );
}

/** Quick Action Bar 컴포넌트 — 홈 상단에 배치되는 빠른 액션 바 */
function QuickActionBar({ onRefreshPrices, onRefreshDailyCharts, refreshLogs, loading }: {
  onRefreshPrices: () => void;
  onRefreshDailyCharts: () => void;
  refreshLogs: RefreshLogs;
  loading: boolean;
}) {
  const priceUpdated = refreshLogs.refreshKrxPrices ? formatLogTime(refreshLogs.refreshKrxPrices) : '-';
  const chartUpdated = refreshLogs.refreshKrxDailyCharts ? formatLogTime(refreshLogs.refreshKrxDailyCharts) : '-';

  return (
    <section className="quick-action-bar">
      <div className="quick-action-item">
        <button
          className="quick-action-btn"
          onClick={onRefreshPrices}
          disabled={loading}
          title="KIS 현재가 갱신"
        >
          <span className="quick-action-icon">⚡</span>
          <span>현재가 갱신</span>
        </button>
        <div className="quick-action-time">마지막 갱신: {priceUpdated}</div>
      </div>
      <div className="quick-action-item">
        <button
          className="quick-action-btn"
          onClick={onRefreshDailyCharts}
          disabled={loading}
          title="일별 차트 배치 갱신"
        >
          <span className="quick-action-icon">📊</span>
          <span>차트 갱신</span>
        </button>
        <div className="quick-action-time">마지막 갱신: {chartUpdated}</div>
      </div>
    </section>
  );
}

/** 부족 종목 Action Card 컴포넌트 */
function DeficitCard({ candidate, onDetail }: { candidate: DeficitCandidate; onDetail: () => void }) {
  const deviationPct = (candidate.deviationRate * 100).toFixed(1);
  return (
    <button className="deficit-card-row" onClick={onDetail}>
      <div className="deficit-card-info">
        <span className="deficit-badge">Deficit</span>
        <strong className="deficit-name">{candidate.item.asset_name}</strong>
        <span className="deficit-account muted">{candidate.item.account_name}</span>
      </div>
      <div className="deficit-card-numbers">
        <span className="deficit-amount">+{formatKRW(candidate.gapAmount)}</span>
        <span className="deficit-deviation negative">비중 괴리 {deviationPct}%</span>
      </div>
      <div className="deficit-card-action muted small">추가 매수 후보 →</div>
    </button>
  );
}

function PortfolioPage({ dashboard, items, accountFilter, setAccountFilter, onEdit, onDetail }: {
  dashboard: DashboardData;
  items: PortfolioItem[];
  accountFilter: string;
  setAccountFilter: (v: string) => void;
  onEdit: (item: PortfolioItem) => void;
  onDetail: (item: PortfolioItem) => void;
}) {
  return (
    <section className="page-stack">
      <AccountFilter dashboard={dashboard} value={accountFilter} onChange={setAccountFilter} />
      <section className="card">
        <div className="section-title">보유 종목 {items.length}개</div>
        <HoldingList rows={items} onEdit={onEdit} onDetail={onDetail} />
      </section>
    </section>
  );
}

function HoldingList({ rows, compact = false, onEdit, onDetail, riskFlags }: {
  rows: PortfolioItem[];
  compact?: boolean;
  onEdit?: (item: PortfolioItem) => void;
  onDetail?: (item: PortfolioItem) => void;
  riskFlags?: Map<string, Set<RiskFlag>>;
}) {
  return (
    <div className="holding-list">
      {rows.map((row) => {
        const flags = riskFlags?.get(buildHoldingKey(row));
        return (
          <article key={`${row.account_id}-${row.asset_id}`} className="holding-card clickable-card" onClick={() => onDetail?.(row)}>
            <div className="holding-main">
              <div>
                <h3>
                  {row.asset_name}
                  {flags?.has('DROP_WARNING') && <span className="risk-dot risk-dot--drop" title="손실 확대 (-20% 이하)">●</span>}
                  {flags?.has('CONCENTRATION_WARNING') && <span className="risk-dot risk-dot--concentration" title="비중 과다 (전체 15% 초과)">●</span>}
                  {flags?.has('WEIGHT_DEVIATION') && <span className="risk-dot risk-dot--deviation" title="비중 괴리 (목표 대비 ±30% 이상)">●</span>}
                </h3>
                <div className="muted small">{row.account_name} · {row.ticker}</div>
              </div>
              <div className="right">
                <strong>{formatKRW(row.valuation_amount)}</strong>
                <span className={returnClass(row.profit_rate)}>{formatPercent(row.profit_rate)}</span>
              </div>
            </div>
            {!compact && (
              <div className="holding-detail">
                <Metric label="수량" value={formatNumber(row.quantity, 4)} />
                <Metric label="평단" value={formatKRW(row.avg_price)} />
                <Metric label="현재가" value={formatKRW(row.price)} />
                <Metric label="계좌비중" value={formatPercent(row.account_weight)} />
                <Metric label="목표비중" value={formatPercent(row.target_weight_account)} />
                <Metric label="목표 대비" value={formatSignedKRW(targetPositionGapAmount(row))} className={signedClass(targetPositionGapAmount(row))} />
              </div>
            )}
            {!compact && onEdit && <button className="full-width ghost" onClick={(event) => { event.stopPropagation(); onEdit(row); }}>수정</button>}
          </article>
        );
      })}
    </div>
  );
}


type RebalanceMode = 'conservative' | 'target' | 'aggressive';

type RebalanceAllocation = {
  item: PortfolioItem;
  amount: number;
  remainingGap: number;
};

type RebalanceSellAllocation = {
  item: PortfolioItem;
  amount: number;
  remainingExcess: number;
};


type RiskSeverity = '주의' | '경고' | '재검토';

type RiskSignal = {
  id: string;
  severity: RiskSeverity;
  type: 'LOSS_ALERT' | 'DRAWDOWN_ALERT' | 'BUY_WITH_CAUTION' | 'TAKE_PROFIT_CANDIDATE' | 'CONCENTRATION_ALERT' | 'DATA_ALERT';
  title: string;
  item?: PortfolioItem;
  description: string;
  metrics: string[];
  actions: string[];
  score: number;
};

type RebalanceAccountRow = {
  account: AccountSummary;
  shortfall: number;
  excess: number;
  cash: number;
  usableCash: number;
  rebalanceBudget: number;
  netShortage: number;
  residualCash: number;
  sellRequired: number;
  sellAllocations: RebalanceSellAllocation[];
  allocations: RebalanceAllocation[];
};

function RebalancePage({ dashboard, items, accountFilter, setAccountFilter, onHoldingDetail, apiUrl }: {
  dashboard: DashboardData;
  items: PortfolioItem[];
  accountFilter: string;
  setAccountFilter: (v: string) => void;
  onHoldingDetail: (item: PortfolioItem) => void;
  apiUrl: string;
}) {
  const [mode, setModeState] = useState<RebalanceMode>(() => {
    const saved = localStorage.getItem(REBALANCE_MODE_KEY) as RebalanceMode | null;
    return saved || 'target';
  });
  const [cashByAccount, setCashByAccount] = useState<Record<string, string>>(() => readRebalanceCash());

  function setMode(next: RebalanceMode) {
    setModeState(next);
    localStorage.setItem(REBALANCE_MODE_KEY, next);
  }

  function updateCash(accountId: string, value: string) {
    const normalized = value.replace(/[^0-9.]/g, '');
    const next = { ...cashByAccount, [accountId]: normalized };
    setCashByAccount(next);
    localStorage.setItem(REBALANCE_CASH_KEY, JSON.stringify(next));
  }

  function clearCash() {
    const confirmed = window.confirm('계좌별 예수금 입력값을 이 브라우저에서 모두 삭제합니다. 계속할까요?');
    if (!confirmed) return;
    setCashByAccount({});
    localStorage.removeItem(REBALANCE_CASH_KEY);
  }

  const visibleItems = accountFilter === 'ALL'
    ? items
    : items.filter((item) => item.account_id === accountFilter);

  const buyCandidates = visibleItems
    .filter((item) => targetShortfallAmount(item) > 0)
    .sort((a, b) => targetShortfallAmount(b) - targetShortfallAmount(a));

  const excessCandidates = visibleItems
    .filter((item) => targetExcessAmount(item) > 0)
    .sort((a, b) => targetExcessAmount(b) - targetExcessAmount(a));

  const riskSignals = useMemo(
    () => buildRiskSignals(visibleItems, apiUrl).slice(0, 12),
    [visibleItems, apiUrl],
  );

  const accountRebalanceRows: RebalanceAccountRow[] = dashboard.accounts.map((account) => {
    const accountItems = items.filter((item) => item.account_id === account.account_id);
    const shortfall = accountItems.reduce((sum, item) => sum + targetShortfallAmount(item), 0);
    const excess = accountItems.reduce((sum, item) => sum + targetExcessAmount(item), 0);
    const cash = Number(cashByAccount[account.account_id] || 0);
    const usableCash = cashToUseByMode(cash, mode);
    const rebalanceBudget = excess + usableCash;
    const netShortage = Math.max(shortfall - rebalanceBudget, 0);
    const residualCash = Math.max(rebalanceBudget - shortfall, 0);
    const sellRequired = Math.min(excess, Math.max(shortfall - usableCash, 0));
    const sellAllocations = makeRebalanceSellAllocations(accountItems, sellRequired, mode);
    const allocations = makeRebalanceAllocations(accountItems, rebalanceBudget, mode);
    return { account, shortfall, excess, cash, usableCash, rebalanceBudget, netShortage, residualCash, sellRequired, sellAllocations, allocations };
  });

  const visibleAccountRows = accountFilter === 'ALL'
    ? accountRebalanceRows
    : accountRebalanceRows.filter((row) => row.account.account_id === accountFilter);

  const totalShortfall = visibleAccountRows.reduce((sum, row) => sum + row.shortfall, 0);
  const totalExcess = visibleAccountRows.reduce((sum, row) => sum + row.excess, 0);
  const totalCash = visibleAccountRows.reduce((sum, row) => sum + row.cash, 0);
  const totalUsableCash = visibleAccountRows.reduce((sum, row) => sum + row.usableCash, 0);
  const totalBudget = visibleAccountRows.reduce((sum, row) => sum + row.rebalanceBudget, 0);
  const totalNetShortage = visibleAccountRows.reduce((sum, row) => sum + row.netShortage, 0);
  const totalResidualCash = visibleAccountRows.reduce((sum, row) => sum + row.residualCash, 0);
  const totalAllocated = visibleAccountRows.reduce((sum, row) => sum + row.allocations.reduce((s, a) => s + a.amount, 0), 0);

  return (
    <section className="page-stack rebalance-page">
      <AccountFilter dashboard={dashboard} value={accountFilter} onChange={setAccountFilter} />

      <section className="hero-card rebalance-hero">
        <div className="muted">리밸런싱 후 추가 필요 현금</div>
        <div className="hero-value">{formatKRW(totalNetShortage)}</div>
        <div className="summary-grid">
          <Metric label="부족 총액" value={formatKRW(totalNetShortage)} className={totalNetShortage > 0 ? 'shortfall' : 'neutral'} />
          <Metric label="초과 총액" value={formatKRW(totalResidualCash)} className={totalResidualCash > 0 ? 'excess' : 'neutral'} />
          <Metric label="매수 필요액" value={formatKRW(totalShortfall)} className="shortfall" />
          <Metric label="매도 재원" value={formatKRW(totalExcess)} className="excess" />
          <Metric label="사용 예수금" value={formatKRW(totalUsableCash)} />
          <Metric label="매수 제안" value={formatKRW(totalAllocated)} />
        </div>
        <p className="muted small">부족/초과 총액은 목표 초과 종목 매도분 + 입력 예수금으로 부족 종목을 매수한다고 가정한 뒤의 순부족/잔여 현금이다.</p>
      </section>

      <section className="card rebalance-mode-card">
        <div className="section-title">리밸런싱 모드</div>
        <div className="mode-grid">
          <button className={mode === 'conservative' ? 'active' : ''} onClick={() => setMode('conservative')}>
            <strong>보수적</strong>
            <span>초과분 전액 + 예수금 50% · 부족 큰 순서</span>
          </button>
          <button className={mode === 'target' ? 'active' : ''} onClick={() => setMode('target')}>
            <strong>목표 근접</strong>
            <span>초과분 전액 + 예수금 100% · 부족액 비례</span>
          </button>
          <button className={mode === 'aggressive' ? 'active' : ''} onClick={() => setMode('aggressive')}>
            <strong>공격적</strong>
            <span>초과분 전액 + 예수금 100% · 큰 부족 순차</span>
          </button>
        </div>
      </section>

      <section className="card risk-signal-card">
        <div className="section-title">리스크 신호</div>
        <div className="risk-summary-grid">
          <Metric label="재검토" value={String(riskSignals.filter((signal) => signal.severity === '재검토').length)} className="risk-critical" />
          <Metric label="경고" value={String(riskSignals.filter((signal) => signal.severity === '경고').length)} className="risk-warning" />
          <Metric label="주의" value={String(riskSignals.filter((signal) => signal.severity === '주의').length)} className="risk-caution" />
        </div>
        <p className="muted small">목표비중과 별개로 추가매수 보류, 손실 확대, 고점 대비 낙폭, 비중 과집중, 데이터 이상 여부를 검토하는 보조 신호다. 매수/매도 지시가 아니라 확인 후보로만 사용한다.</p>
        <RiskSignalList signals={riskSignals} onSelect={onHoldingDetail} />
      </section>

      <section className="card">
        <div className="section-title">계좌별 예수금 입력</div>
        <div className="cash-input-list">
          {dashboard.accounts.map((account) => (
            <label key={account.account_id} className="cash-input-row">
              <div>
                <strong>{account.account_name}</strong>
                <span className="muted small">{account.broker} · 매수 필요 {formatKRW(accountRebalanceRows.find((row) => row.account.account_id === account.account_id)?.shortfall || 0)}</span>
              </div>
              <input
                inputMode="numeric"
                value={cashByAccount[account.account_id] || ''}
                onChange={(event) => updateCash(account.account_id, event.target.value)}
                placeholder="예수금"
              />
            </label>
          ))}
        </div>
        <div className="row-actions">
          <button className="ghost" onClick={clearCash}>예수금 입력값 초기화</button>
        </div>
        <p className="muted small">입력값은 이 브라우저 localStorage에만 저장된다. 리밸런싱 재원은 목표 초과분 매도 가능액 + 모드별 사용 예수금으로 계산한다.</p>
      </section>

      <section className="card">
        <div className="section-title">계좌별 리밸런싱 요약</div>
        <div className="rebalance-account-list">
          {visibleAccountRows.map((row) => {
            const accountItems = items.filter((item) => item.account_id === row.account.account_id);
            const accountExcessItems = accountItems
              .filter((item) => targetExcessAmount(item) > 0)
              .sort((a, b) => targetExcessAmount(b) - targetExcessAmount(a));
            const plannedSellMap = new Map(row.sellAllocations.map((allocation) => [`${allocation.item.account_id}__${allocation.item.asset_id}`, allocation.amount]));
            return (
              <div key={row.account.account_id} className="rebalance-account-card">
                <div className="rebalance-account-head">
                  <div>
                    <strong>{row.account.account_name}</strong>
                    <div className="muted small">{row.account.broker} · {row.account.account_type}</div>
                  </div>
                  <div className="right">
                    <strong>{formatKRW(row.account.valuation_amount)}</strong>
                    <span className="muted small">평가액</span>
                  </div>
                </div>
                <div className="summary-grid compact-summary-grid">
                  <Metric label="매수 필요" value={formatKRW(row.shortfall)} className="shortfall" />
                  <Metric label="매도 재원" value={formatKRW(row.excess)} className="excess" />
                  <Metric label="사용 예수금" value={formatKRW(row.usableCash)} />
                  <Metric label="부족" value={formatKRW(row.netShortage)} className={row.netShortage > 0 ? 'shortfall' : 'neutral'} />
                  <Metric label="잔여" value={formatKRW(row.residualCash)} className={row.residualCash > 0 ? 'excess' : 'neutral'} />
                  <Metric label="입력 예수금" value={formatKRW(row.cash)} />
                </div>
                <div className="rebalance-action-block">
                  <div className="rebalance-subtitle">
                    <strong>매도 제안</strong>
                    <span className="muted small">필요 매도 {formatKRW(row.sellRequired)} · 초과분 {formatKRW(row.excess)}</span>
                  </div>
                  {accountExcessItems.length > 0 ? (
                    <div className="allocation-list">
                      {accountExcessItems.slice(0, 5).map((item) => {
                        const key = `${item.account_id}__${item.asset_id}`;
                        const plannedSell = plannedSellMap.get(key) || 0;
                        const excessAmount = targetExcessAmount(item);
                        return (
                          <button key={`${item.account_id}-${item.asset_id}-sell`} className="allocation-row sell-allocation-row" onClick={() => onHoldingDetail(item)}>
                            <div>
                              <strong>{item.asset_name}</strong>
                              <span className="muted small">{item.ticker} · 초과 {formatKRW(excessAmount)}</span>
                            </div>
                            <div className="right">
                              <strong className={plannedSell > 0 ? 'excess' : 'neutral'}>{formatKRW(plannedSell > 0 ? plannedSell : excessAmount)}</strong>
                              <span className="muted small">{plannedSell > 0 ? '매도 제안' : '초과 후보'}</span>
                            </div>
                          </button>
                        );
                      })}
                    </div>
                  ) : (
                    <div className="empty-inline">목표비중을 초과한 매도 후보가 없다.</div>
                  )}
                </div>

                <div className="rebalance-action-block">
                  <div className="rebalance-subtitle">
                    <strong>매수 제안</strong>
                    <span className="muted small">제안 합계 {formatKRW(row.allocations.reduce((sum, allocation) => sum + allocation.amount, 0))}</span>
                  </div>
                  {row.allocations.length > 0 ? (
                    <div className="allocation-list">
                      {row.allocations.slice(0, 5).map((allocation) => (
                        <button key={`${allocation.item.account_id}-${allocation.item.asset_id}-buy`} className="allocation-row" onClick={() => onHoldingDetail(allocation.item)}>
                          <div>
                            <strong>{allocation.item.asset_name}</strong>
                            <span className="muted small">{allocation.item.ticker} · 부족 {formatKRW(targetShortfallAmount(allocation.item))}</span>
                          </div>
                          <div className="right">
                            <strong className="shortfall">{formatKRW(allocation.amount)}</strong>
                            <span className="muted small">매수 제안</span>
                          </div>
                        </button>
                      ))}
                    </div>
                  ) : (
                    <div className="empty-inline">리밸런싱 재원이 없거나 추가 매수 후보가 없다.</div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </section>

      <section className="card">
        <div className="section-title">목표비중 미달 Top 5</div>
        <RebalanceCandidateList rows={buyCandidates.slice(0, 5)} kind="buy" onSelect={onHoldingDetail} />
      </section>

      <section className="card">
        <div className="section-title">목표비중 초과 Top 5</div>
        <RebalanceCandidateList rows={excessCandidates.slice(0, 5)} kind="excess" onSelect={onHoldingDetail} />
      </section>
    </section>
  );
}


function RiskSignalList({ signals, onSelect }: { signals: RiskSignal[]; onSelect: (item: PortfolioItem) => void }) {
  if (signals.length === 0) {
    return <div className="empty-inline">현재 기준으로 특별한 리스크 신호가 없다.</div>;
  }

  return (
    <div className="risk-signal-list">
      {signals.map((signal) => (
        <button
          key={signal.id}
          className={`risk-signal-row risk-${riskSeverityClass(signal.severity)}`}
          onClick={() => signal.item && onSelect(signal.item)}
          disabled={!signal.item}
        >
          <div className="risk-signal-head">
            <span className={`risk-badge risk-${riskSeverityClass(signal.severity)}`}>{signal.severity}</span>
            <strong>{signal.title}</strong>
          </div>
          {signal.item && <div className="muted small">{signal.item.account_name} · {signal.item.ticker} · {signal.item.asset_name}</div>}
          <p>{signal.description}</p>
          <div className="risk-metric-list">
            {signal.metrics.map((metric) => <span key={metric}>{metric}</span>)}
          </div>
          <div className="risk-action-list">
            {signal.actions.map((action) => <span key={action}>{action}</span>)}
          </div>
        </button>
      ))}
    </div>
  );
}

function RebalanceCandidateList({ rows, kind, onSelect }: { rows: PortfolioItem[]; kind: 'buy' | 'excess'; onSelect: (item: PortfolioItem) => void }) {
  if (rows.length === 0) return <div className="empty-inline">대상 종목이 없다.</div>;
  return (
    <div className="rebalance-candidate-list">
      {rows.map((item, index) => {
        const amount = kind === 'buy' ? targetShortfallAmount(item) : targetExcessAmount(item);
        const gapRate = item.account_weight - item.target_weight_account;
        return (
          <button key={`${item.account_id}-${item.asset_id}`} className="rebalance-candidate-row" onClick={() => onSelect(item)}>
            <span className="candidate-rank">{index + 1}</span>
            <div className="candidate-body">
              <strong>{item.asset_name}</strong>
              <span className="muted small">{item.account_name} · {item.ticker} · 현재 {formatPercent(item.account_weight)} / 목표 {formatPercent(item.target_weight_account)}</span>
              <div className="bar-bg candidate-bar">
                <div className={kind === 'buy' ? 'bar-fill shortfall-fill' : 'bar-fill excess-fill'} style={{ width: `${Math.min(Math.max(Math.abs(gapRate) * 100, 3), 100)}%` }} />
              </div>
            </div>
            <div className="right">
              <strong className={kind === 'buy' ? 'shortfall' : 'excess'}>{formatKRW(amount)}</strong>
              <span className="muted small">{kind === 'buy' ? '부족' : '초과'}</span>
            </div>
          </button>
        );
      })}
    </div>
  );
}


function buildRiskSignals(items: PortfolioItem[], apiUrl: string): RiskSignal[] {
  const signals: RiskSignal[] = [];
  const seen = new Set<string>();

  function push(signal: RiskSignal) {
    if (seen.has(signal.id)) return;
    seen.add(signal.id);
    signals.push(signal);
  }

  items.forEach((item) => {
    const profitRate = Number(item.profit_rate || 0);
    const shortfall = targetShortfallAmount(item);
    const excess = targetExcessAmount(item);
    const totalWeight = Number(item.total_weight || 0);
    const price = Number(item.price || 0);
    const valuation = Number(item.valuation_amount || 0);
    const priceFetchedAt = String(item.price_fetched_at || '');

    if (price <= 0 || valuation <= 0) {
      push({
        id: `DATA_PRICE_${item.account_id}_${item.asset_id}`,
        severity: '재검토',
        type: 'DATA_ALERT',
        title: '가격 데이터 확인 필요',
        item,
        description: '현재가 또는 평가액이 0 이하로 계산되어 포트폴리오 판단값이 왜곡될 수 있다.',
        metrics: [`현재가 ${formatKRW(price)}`, `평가액 ${formatKRW(valuation)}`],
        actions: ['현재가 갱신', '원장 가격/평가액 확인', 'App_Prices 확인'],
        score: 9000,
      });
    }

    if (isStalePriceFetchedAt(priceFetchedAt, 3)) {
      push({
        id: `DATA_STALE_${item.account_id}_${item.asset_id}`,
        severity: '주의',
        type: 'DATA_ALERT',
        title: '현재가 갱신 시각 확인',
        item,
        description: '가격 갱신 시각이 오래되었을 수 있다. 리밸런싱 전 현재가 갱신을 권장한다.',
        metrics: [`가격 갱신 ${priceFetchedAt || '없음'}`],
        actions: ['현재가 갱신', '미국 종목 KIS 갱신 결과 확인'],
        score: 1000,
      });
    }

    if (profitRate <= -0.3) {
      push(makeLossSignal(item, '재검토', profitRate, '손실률이 -30% 이하로 확대되었다. 목표비중 유지 여부와 보유 논리 재검토가 필요하다.', 8000));
    } else if (profitRate <= -0.2) {
      push(makeLossSignal(item, '경고', profitRate, '손실률이 -20% 이하로 확대되었다. 추가매수 또는 보유 지속 판단 전 점검이 필요하다.', 6000));
    } else if (profitRate <= -0.1) {
      push(makeLossSignal(item, '주의', profitRate, '손실률이 -10% 이하로 내려갔다. 일시 조정인지 구조적 부진인지 관찰 대상이다.', 3000));
    }

    if (shortfall > 0 && profitRate <= -0.25) {
      push(makeBuyCautionSignal(item, '재검토', shortfall, profitRate, '목표비중상 추가매수 후보지만 손실률이 커서 기계적 물타기 전 재검토가 필요하다.', 8500));
    } else if (shortfall > 0 && profitRate <= -0.15) {
      push(makeBuyCautionSignal(item, '경고', shortfall, profitRate, '목표비중상 추가매수 후보지만 손실률이 커서 추가매수 전 확인이 필요하다.', 6500));
    } else if (shortfall > 0 && profitRate <= -0.08) {
      push(makeBuyCautionSignal(item, '주의', shortfall, profitRate, '목표비중상 추가매수 후보이나 손실 구간이다. 매수 속도 조절 검토 대상이다.', 3500));
    }

    if (excess > 0 && profitRate >= 0.5) {
      push(makeTakeProfitSignal(item, '경고', excess, profitRate, '목표비중 초과 상태에서 수익률도 높다. 일부 이익 실현 또는 비중 조절 후보로 볼 수 있다.', 6200));
    } else if (excess > 0 && profitRate >= 0.3) {
      push(makeTakeProfitSignal(item, '주의', excess, profitRate, '목표비중 초과 상태에서 수익률이 높다. 추가 매수는 보류하고 비중 관리가 필요하다.', 3600));
    }

    if (totalWeight >= 0.2) {
      push(makeConcentrationSignal(item, '경고', totalWeight, '단일 종목 전체 비중이 20% 이상이다. 포트폴리오 집중 위험을 점검해야 한다.', 6200));
    } else if (totalWeight >= 0.15) {
      push(makeConcentrationSignal(item, '주의', totalWeight, '단일 종목 전체 비중이 15% 이상이다. 의도한 핵심 비중인지 확인할 필요가 있다.', 3200));
    }

    const cachedDaily = readCachedChartDataForAsset(apiUrl, item.asset_id, 'D');
    const drawdown = calculateDrawdown(cachedDaily);
    if (drawdown !== null) {
      if (drawdown <= -0.3) {
        push(makeDrawdownSignal(item, '재검토', drawdown, '캐시된 일봉 기준 최근 고점 대비 -30% 이상 하락했다. 추세 훼손 여부 재검토가 필요하다.', 7600));
      } else if (drawdown <= -0.2) {
        push(makeDrawdownSignal(item, '경고', drawdown, '캐시된 일봉 기준 최근 고점 대비 -20% 이상 하락했다. 추가매수 전 낙폭 원인 확인이 필요하다.', 5600));
      } else if (drawdown <= -0.1) {
        push(makeDrawdownSignal(item, '주의', drawdown, '캐시된 일봉 기준 최근 고점 대비 -10% 이상 하락했다. 단기 낙폭 관찰 대상이다.', 2600));
      }
    }
  });

  return signals.sort((a, b) => riskSeverityScore(b.severity) - riskSeverityScore(a.severity) || b.score - a.score);
}

function makeLossSignal(item: PortfolioItem, severity: RiskSeverity, profitRate: number, description: string, score: number): RiskSignal {
  return {
    id: `LOSS_${item.account_id}_${item.asset_id}_${severity}`,
    severity,
    type: 'LOSS_ALERT',
    title: '손실 확대',
    item,
    description,
    metrics: [`수익률 ${formatPercent(profitRate)}`, `평가손익 ${formatKRW(item.profit_amount)}`],
    actions: ['보유 논리 확인', '추가매수 속도 조절', '목표비중 재검토'],
    score: score + Math.abs(profitRate) * 100,
  };
}

function makeBuyCautionSignal(item: PortfolioItem, severity: RiskSeverity, shortfall: number, profitRate: number, description: string, score: number): RiskSignal {
  return {
    id: `BUY_CAUTION_${item.account_id}_${item.asset_id}_${severity}`,
    severity,
    type: 'BUY_WITH_CAUTION',
    title: '추가매수 주의',
    item,
    description,
    metrics: [`부족 ${formatKRW(shortfall)}`, `수익률 ${formatPercent(profitRate)}`],
    actions: ['기계적 매수 보류', '동일 자산군 비교', '목표비중 유지 여부 확인'],
    score: score + shortfall / 100000,
  };
}

function makeTakeProfitSignal(item: PortfolioItem, severity: RiskSeverity, excess: number, profitRate: number, description: string, score: number): RiskSignal {
  return {
    id: `TAKE_PROFIT_${item.account_id}_${item.asset_id}_${severity}`,
    severity,
    type: 'TAKE_PROFIT_CANDIDATE',
    title: '이익 실현 후보',
    item,
    description,
    metrics: [`초과 ${formatKRW(excess)}`, `수익률 ${formatPercent(profitRate)}`],
    actions: ['추가매수 보류', '일부 매도 검토', '목표비중 재확인'],
    score: score + excess / 100000,
  };
}

function makeConcentrationSignal(item: PortfolioItem, severity: RiskSeverity, totalWeight: number, description: string, score: number): RiskSignal {
  return {
    id: `CONCENTRATION_${item.account_id}_${item.asset_id}_${severity}`,
    severity,
    type: 'CONCENTRATION_ALERT',
    title: '비중 과집중',
    item,
    description,
    metrics: [`전체비중 ${formatPercent(totalWeight)}`, `평가액 ${formatKRW(item.valuation_amount)}`],
    actions: ['전체 포트폴리오 내 역할 확인', '상한 비중 설정 검토', '리밸런싱 매도 후보 확인'],
    score: score + totalWeight * 100,
  };
}

function makeDrawdownSignal(item: PortfolioItem, severity: RiskSeverity, drawdown: number, description: string, score: number): RiskSignal {
  return {
    id: `DRAWDOWN_${item.account_id}_${item.asset_id}_${severity}`,
    severity,
    type: 'DRAWDOWN_ALERT',
    title: '고점 대비 낙폭',
    item,
    description,
    metrics: [`고점 대비 ${formatPercent(drawdown)}`, '일봉 캐시 기준'],
    actions: ['차트 추세 확인', '추가매수 전 원인 점검', '동일 자산군 비교'],
    score: score + Math.abs(drawdown) * 100,
  };
}

function calculateDrawdown(data: ChartData | null): number | null {
  const items = data?.items ?? [];
  if (items.length < 10) return null;
  const closes = items.map((point) => Number(point.close || 0)).filter((value) => value > 0);
  if (closes.length < 10) return null;
  const latest = closes[closes.length - 1];
  const high = Math.max(...closes);
  if (high <= 0) return null;
  return latest / high - 1;
}

function isStalePriceFetchedAt(value: string, staleDays: number): boolean {
  if (!value) return false;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return false;
  return Date.now() - date.getTime() > staleDays * 24 * 60 * 60 * 1000;
}

function riskSeverityScore(severity: RiskSeverity): number {
  if (severity === '재검토') return 3;
  if (severity === '경고') return 2;
  return 1;
}

function riskSeverityClass(severity: RiskSeverity): 'critical' | 'warning' | 'caution' {
  if (severity === '재검토') return 'critical';
  if (severity === '경고') return 'warning';
  return 'caution';
}

function readRebalanceCash(): Record<string, string> {
  try {
    const raw = localStorage.getItem(REBALANCE_CASH_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch {
    return {};
  }
}

function targetShortfallAmount(item: PortfolioItem): number {
  return Math.max(Number(item.target_gap_amount || 0), 0);
}

function targetExcessAmount(item: PortfolioItem): number {
  return Math.max(-Number(item.target_gap_amount || 0), 0);
}

function cashToUseByMode(cash: number, mode: RebalanceMode): number {
  if (cash <= 0) return 0;
  return mode === 'conservative' ? cash * 0.5 : cash;
}

function makeRebalanceSellAllocations(items: PortfolioItem[], sellNeed: number, mode: RebalanceMode): RebalanceSellAllocation[] {
  const candidates = items
    .filter((item) => targetExcessAmount(item) > 0)
    .sort((a, b) => targetExcessAmount(b) - targetExcessAmount(a));

  if (sellNeed <= 0 || candidates.length === 0) return [];

  if (mode === 'target') {
    const totalExcess = candidates.reduce((sum, item) => sum + targetExcessAmount(item), 0);
    if (totalExcess <= 0) return [];
    return candidates
      .map((item) => {
        const excess = targetExcessAmount(item);
        const amount = Math.min(excess, sellNeed * (excess / totalExcess));
        return { item, amount, remainingExcess: excess - amount };
      })
      .filter((row) => row.amount > 0);
  }

  let remainingNeed = sellNeed;
  const allocations: RebalanceSellAllocation[] = [];
  candidates.forEach((item) => {
    if (remainingNeed <= 0) return;
    const excess = targetExcessAmount(item);
    const amount = Math.min(excess, remainingNeed);
    if (amount > 0) {
      allocations.push({ item, amount, remainingExcess: excess - amount });
      remainingNeed -= amount;
    }
  });
  return allocations;
}

function makeRebalanceAllocations(items: PortfolioItem[], budget: number, mode: RebalanceMode): RebalanceAllocation[] {
  const candidates = items
    .filter((item) => targetShortfallAmount(item) > 0)
    .sort((a, b) => targetShortfallAmount(b) - targetShortfallAmount(a));

  if (budget <= 0 || candidates.length === 0) return [];

  if (mode === 'target') {
    const totalGap = candidates.reduce((sum, item) => sum + targetShortfallAmount(item), 0);
    if (totalGap <= 0) return [];
    return candidates
      .map((item) => {
        const gap = targetShortfallAmount(item);
        const amount = Math.min(gap, budget * (gap / totalGap));
        return { item, amount, remainingGap: gap - amount };
      })
      .filter((row) => row.amount > 0);
  }

  let remainingBudget = budget;
  const allocations: RebalanceAllocation[] = [];
  candidates.forEach((item) => {
    if (remainingBudget <= 0) return;
    const gap = targetShortfallAmount(item);
    const amount = Math.min(gap, remainingBudget);
    if (amount > 0) {
      allocations.push({ item, amount, remainingGap: gap - amount });
      remainingBudget -= amount;
    }
  });
  return allocations;
}


function DetailModal({ detail, dashboard, dividends, api, onClose, onEditHolding }: {
  detail: DetailState;
  dashboard: DashboardData;
  dividends: DividendDashboard | null;
  api: EdsApi;
  onClose: () => void;
  onEditHolding: (item: PortfolioItem) => void;
}) {
  if (detail.kind === 'account') {
    const account = detail.account;
    const accountItems = sortByValuationDesc(
      dashboard.output.filter((row) => row.account_id === account.account_id),
    );
    return (
      <div className="modal-backdrop" onClick={onClose}>
        <div className="modal detail-modal" onClick={(event) => event.stopPropagation()}>
          <div className="modal-header">
            <div>
              <h2>{account.account_name}</h2>
              <p className="muted">{account.broker} · {account.account_type}</p>
            </div>
            <button className="ghost close-button" onClick={onClose}>닫기</button>
          </div>

          <div className="summary-grid">
            <Metric label="평가액" value={formatKRW(account.valuation_amount)} />
            <Metric label="총손익" value={formatKRW(account.profit_amount)} className={signedClass(account.profit_amount)} />
            <Metric label="수익률" value={formatPercent(account.profit_rate)} className={returnClass(account.profit_rate)} />
            <Metric label="전체비중" value={formatPercent(account.weight)} />
          </div>

          <section className="detail-section">
            <div className="section-title">계좌 내 종목 구성</div>
            <CompositionRing rows={accountItems} mode="account" centerLabel={formatPercent(account.weight)} />
            <BreakdownList rows={accountItems} mode="account" onSelect={(item) => onEditHolding(item)} />
          </section>
        </div>
      </div>
    );
  }

  const item = detail.item;
  const currentWeight = item.account_weight;
  const targetWeight = item.target_weight_account;
  const gapRate = currentWeight - targetWeight;
  const targetGapAmount = targetPositionGapAmount(item);

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal detail-modal" onClick={(event) => event.stopPropagation()}>
        <div className="modal-header">
          <div>
            <h2>{item.asset_name}</h2>
            <p className="muted">{item.account_name} · {item.ticker}</p>
          </div>
          <button className="ghost close-button" onClick={onClose}>닫기</button>
        </div>

        <section className="detail-hero">
          <div>
            <span className="muted small">평가액</span>
            <strong>{formatKRW(item.valuation_amount)}</strong>
          </div>
          <div className="right">
            <span className={returnClass(item.profit_rate)}>{formatPercent(item.profit_rate)}</span>
            <span className={signedClass(item.profit_amount)}>{formatKRW(item.profit_amount)}</span>
          </div>
        </section>

        <div className="summary-grid">
          <Metric label="수량" value={formatNumber(item.quantity, 4)} />
          <Metric label="평단" value={formatKRW(item.avg_price)} />
          <Metric label="현재가" value={formatKRW(item.price)} />
          <Metric label="전체비중" value={formatPercent(item.total_weight)} />
        </div>

        <section className="detail-section">
          <div className="section-title">목표비중 비교</div>
          <ComparisonBar label="현재 계좌비중" value={currentWeight} />
          <ComparisonBar label="목표 계좌비중" value={targetWeight} variant="target" />
          <div className="insight-card">
            <span>목표 대비</span>
            <strong className={signedClass(targetGapAmount)}>{formatSignedKRW(targetGapAmount)}</strong>
            <p className="muted small">
              {targetGapAmount < 0
                ? '목표 대비 부족. 추가 매수 후보.'
                : targetGapAmount > 0
                  ? '목표 대비 초과. 매수 보류 또는 일부 매도 후보.'
                  : '목표비중과 일치.'}
            </p>
            <p className="muted small">비중 차이 {formatSignedPercent(gapRate)}</p>
          </div>
        </section>


        <HoldingDividendSummary item={item} dividends={dividends} />

        <section className="detail-section">
          <div className="section-title">가격 차트</div>
          <AssetChart item={item} api={api} />
        </section>

        <button onClick={() => onEditHolding(item)}>이 종목 수정</button>
      </div>
    </div>
  );
}


function AssetChart({ item, api }: { item: PortfolioItem; api: EdsApi }) {
  const market = getMarketFromAssetId(item);

  if (market === 'KRX') {
    return <KisLineChart item={item} api={api} />;
  }

  return <TradingViewChart item={item} />;
}

function KisLineChart({ item, api }: { item: PortfolioItem; api: EdsApi }) {
  const [interval, setInterval] = useState<'D' | 'W' | 'M'>('D');
  const [chartData, setChartData] = useState<ChartData | null>(null);
  const [chartCache, setChartCache] = useState<Record<string, ChartData>>({});
  const [status, setStatus] = useState<'idle' | 'loading' | 'ready' | 'error'>('idle');
  const [chartError, setChartError] = useState('');
  const [lastLoadSource, setLastLoadSource] = useState<'memory' | 'local' | 'network' | ''>('');

  const ticker = normalizeKrxTicker(item.ticker);

  async function loadChart(nextInterval: 'D' | 'W' | 'M', force = false) {
    if (!force && chartCache[nextInterval]) {
      setChartData(chartCache[nextInterval]);
      setLastLoadSource('memory');
      setStatus('ready');
      setChartError('');
      return;
    }

    setStatus('loading');
    setChartError('');
    try {
      const data = force
        ? await api.refreshChartData({ asset_id: item.asset_id, ticker, market: 'KRX', interval: nextInterval })
        : await api.getChartData({ asset_id: item.asset_id, ticker, market: 'KRX', interval: nextInterval });
      setChartCache((prev) => ({ ...prev, [nextInterval]: data }));
      setChartData(data);
      setLastLoadSource(force ? 'network' : 'local');
      setStatus('ready');
    } catch (e) {
      setChartError(e instanceof Error ? e.message : String(e));
      setStatus('error');
    }
  }

  useEffect(() => {
    setChartCache({});
    setChartData(null);
    setInterval('D');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [item.asset_id]);

  useEffect(() => {
    loadChart(interval, false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [item.asset_id, interval]);

  const items = chartData?.items ?? [];
  const latest = items[items.length - 1];
  const first = items[0];
  const change = first && latest && first.close > 0 ? latest.close - first.close : 0;
  const changeRate = first && latest && first.close > 0 ? change / first.close : 0;

  return (
    <div className="tv-card kis-chart-card">
      <div className="tv-header">
        <div>
          <div className="muted small">KIS 국내 차트</div>
          <strong>{ticker} · {intervalLabel(interval)}</strong>
        </div>
        <div className="tv-intervals" role="group" aria-label="차트 주기 선택">
          <button className={interval === 'D' ? 'active' : ''} onClick={() => setInterval('D')}>일</button>
          <button className={interval === 'W' ? 'active' : ''} onClick={() => setInterval('W')}>주</button>
          <button className={interval === 'M' ? 'active' : ''} onClick={() => setInterval('M')}>월</button>
        </div>
      </div>

      <div className="chart-stat-row">
        <Metric label="최근 종가" value={latest ? formatKRW(latest.close) : '-'} />
        <Metric label="기간 등락" value={latest && first ? formatKRW(change) : '-'} className={signedClass(change)} />
        <Metric label="기간 수익률" value={latest && first ? formatPercent(changeRate) : '-'} className={returnClass(changeRate)} />
      </div>

      <div className="kis-chart-frame">
        {status === 'loading' && <div className="tv-overlay muted small">KIS 차트 데이터 로딩 중...</div>}
        {status === 'error' && (
          <div className="tv-overlay tv-error">
            <strong>차트를 불러오지 못했습니다.</strong>
            <span>{chartError}</span>
            <button className="ghost" onClick={() => loadChart(interval, true)}>다시 가져오기</button>
          </div>
        )}
        {status === 'ready' && items.length === 0 && (
          <div className="tv-overlay tv-error">
            <strong>캐시된 차트 데이터가 없습니다.</strong>
            <span>이 주기의 데이터는 처음 클릭 시 KIS에서 가져오도록 설정되어 있다.</span>
            <button className="ghost" onClick={() => loadChart(interval, true)}>지금 가져오기</button>
          </div>
        )}
        {items.length > 0 && <LineChart points={items} />}
      </div>

      <div className="chart-actions">
        <span className="muted small">{items.length}개 데이터 · source: kis · load: {loadSourceLabel(lastLoadSource)}</span>
        <button className="ghost" onClick={() => loadChart(interval, true)} disabled={status === 'loading'}>차트 새로고침</button>
      </div>
      <p className="muted small tv-note">일봉은 선캐시, 주봉/월봉은 최초 조회 시 KIS에서 가져와 App_ChartPrices에 저장한다. 같은 종목/주기는 브라우저 캐시를 우선 사용한다.</p>
    </div>
  );
}


function normalizeKrxTicker(value: string | number): string {
  const raw = String(value || '').trim().toUpperCase().replace(/\s/g, '');
  if (/^\d+$/.test(raw) && raw.length < 6) return raw.padStart(6, '0');
  return raw;
}

function loadSourceLabel(source: 'memory' | 'local' | 'network' | ''): string {
  if (source === 'memory') return 'memory';
  if (source === 'local') return 'cache/api';
  if (source === 'network') return 'refresh';
  return '-';
}

function LineChart({ points }: { points: ChartPoint[] }) {
  const width = 720;
  const height = 260;
  const padding = 24;
  const closes = points.map((p) => Number(p.close) || 0).filter((v) => v > 0);
  const min = Math.min(...closes);
  const max = Math.max(...closes);
  const range = Math.max(max - min, 1);
  const last = points[points.length - 1];
  const first = points[0];
  const positive = last && first ? last.close >= first.close : true;

  const coords = points.map((point, index) => {
    const x = padding + (index / Math.max(points.length - 1, 1)) * (width - padding * 2);
    const y = height - padding - ((Number(point.close) - min) / range) * (height - padding * 2);
    return { x, y, point };
  });

  const path = coords.map((c, index) => `${index === 0 ? 'M' : 'L'} ${c.x.toFixed(2)} ${c.y.toFixed(2)}`).join(' ');
  const areaPath = `${path} L ${width - padding} ${height - padding} L ${padding} ${height - padding} Z`;

  return (
    <div className="line-chart-wrap">
      <svg className="line-chart" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="종가 라인 차트">
        <line x1={padding} y1={padding} x2={padding} y2={height - padding} className="chart-axis" />
        <line x1={padding} y1={height - padding} x2={width - padding} y2={height - padding} className="chart-axis" />
        <line x1={padding} y1={padding} x2={width - padding} y2={padding} className="chart-grid" />
        <line x1={padding} y1={height / 2} x2={width - padding} y2={height / 2} className="chart-grid" />
        <path d={areaPath} className={positive ? 'chart-area positive' : 'chart-area negative'} />
        <path d={path} className={positive ? 'chart-line positive' : 'chart-line negative'} />
        {coords.length > 0 && (
          <circle cx={coords[coords.length - 1].x} cy={coords[coords.length - 1].y} r="4" className={positive ? 'chart-dot positive' : 'chart-dot negative'} />
        )}
      </svg>
      <div className="chart-scale">
        <span>{formatKRW(max)}</span>
        <span>{formatKRW(min)}</span>
      </div>
      <div className="chart-dates">
        <span>{formatChartDate(points[0]?.date)}</span>
        <span>{formatChartDate(points[points.length - 1]?.date)}</span>
      </div>
    </div>
  );
}

function intervalLabel(interval: 'D' | 'W' | 'M') {
  if (interval === 'W') return '주봉';
  if (interval === 'M') return '월봉';
  return '일봉';
}

function formatChartDate(value?: string) {
  const text = String(value || '');
  if (text.length === 8) return `${text.slice(0, 4)}.${text.slice(4, 6)}.${text.slice(6, 8)}`;
  return text || '-';
}

function getMarketFromAssetId(item: PortfolioItem): string {
  return String(item.asset_id || '').split('_')[0].toUpperCase();
}


function TradingViewChart({ item }: { item: PortfolioItem }) {
  const [interval, setInterval] = useState<'D' | 'W' | 'M'>('D');
  const [status, setStatus] = useState<'loading' | 'ready' | 'error'>('loading');
  const containerRef = useRef<HTMLDivElement | null>(null);

  const symbol = useMemo(() => getTradingViewSymbol(item), [item]);

  useEffect(() => {
    let cancelled = false;
    const container = containerRef.current;

    if (!container || !symbol) return undefined;

    setStatus('loading');
    container.innerHTML = '';

    const widgetContainer = document.createElement('div');
    widgetContainer.id = `tv_${sanitizeWidgetId(String(item.asset_id || item.ticker))}_${interval}_${Date.now()}`;
    widgetContainer.className = 'tradingview-widget-container__widget';
    container.appendChild(widgetContainer);

    ensureTradingViewScript()
      .then(() => {
        if (cancelled) return;
        const TradingView = (window as Window & { TradingView?: { widget: new (config: Record<string, unknown>) => unknown } }).TradingView;
        if (!TradingView) throw new Error('TradingView widget script was not loaded');

        new TradingView.widget({
          autosize: true,
          symbol,
          interval,
          timezone: 'Asia/Seoul',
          theme: 'dark',
          style: '1',
          locale: 'kr',
          toolbar_bg: '#111827',
          enable_publishing: false,
          hide_side_toolbar: true,
          allow_symbol_change: true,
          save_image: false,
          hideideas: true,
          studies: [],
          container_id: widgetContainer.id,
        });
        setStatus('ready');
      })
      .catch(() => {
        if (!cancelled) setStatus('error');
      });

    return () => {
      cancelled = true;
      if (container) container.innerHTML = '';
    };
  }, [item.asset_id, item.ticker, symbol, interval]);

  return (
    <div className="tv-card">
      <div className="tv-header">
        <div>
          <div className="muted small">차트 심볼</div>
          <strong>{symbol}</strong>
        </div>
        <div className="tv-intervals" role="group" aria-label="차트 주기 선택">
          <button className={interval === 'D' ? 'active' : ''} onClick={() => setInterval('D')}>일</button>
          <button className={interval === 'W' ? 'active' : ''} onClick={() => setInterval('W')}>주</button>
          <button className={interval === 'M' ? 'active' : ''} onClick={() => setInterval('M')}>월</button>
        </div>
      </div>

      <div className="tv-frame">
        <div ref={containerRef} className="tv-widget-container" />
        {status === 'loading' && <div className="tv-overlay muted small">TradingView 차트 로딩 중...</div>}
        {status === 'error' && (
          <div className="tv-overlay tv-error">
            <strong>차트를 불러오지 못했습니다.</strong>
            <span>TradingView에서 이 심볼을 지원하지 않거나 네트워크가 차단된 상태일 수 있습니다.</span>
          </div>
        )}
      </div>

      <p className="muted small tv-note">
        일부 국내 ETF, 특히 영문이 섞인 신규 KRX 코드는 TradingView 심볼 매핑이 맞지 않을 수 있다. 이 경우 위젯 내 심볼 검색으로 직접 변경 가능.
      </p>
    </div>
  );
}

function ensureTradingViewScript(): Promise<void> {
  const win = window as Window & { TradingView?: unknown; __edsTvScriptPromise?: Promise<void> };

  if (win.TradingView) return Promise.resolve();
  if (win.__edsTvScriptPromise) return win.__edsTvScriptPromise;

  win.__edsTvScriptPromise = new Promise((resolve, reject) => {
    const existing = document.querySelector<HTMLScriptElement>('script[data-eds-tv-script="true"]');
    if (existing) {
      existing.addEventListener('load', () => resolve(), { once: true });
      existing.addEventListener('error', () => reject(new Error('TradingView script load failed')), { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = 'https://s3.tradingview.com/tv.js';
    script.async = true;
    script.dataset.edsTvScript = 'true';
    script.onload = () => resolve();
    script.onerror = () => reject(new Error('TradingView script load failed'));
    document.head.appendChild(script);
  });

  return win.__edsTvScriptPromise;
}

function getTradingViewSymbol(item: PortfolioItem): string {
  const assetId = String(item.asset_id || '').toUpperCase();
  const ticker = String(item.ticker || '').toUpperCase().trim();
  const [marketFromAsset] = assetId.split('_');

  if (marketFromAsset === 'NASDAQ') return `NASDAQ:${ticker}`;
  if (marketFromAsset === 'NYSE') return `NYSE:${ticker}`;
  if (marketFromAsset === 'AMEX') return `AMEX:${ticker}`;
  if (marketFromAsset === 'KRX') return `KRX:${ticker}`;

  if (/^\d{6}$/.test(ticker) || /^[0-9A-Z]{6}$/.test(ticker)) return `KRX:${ticker}`;
  if (/^[A-Z]{1,5}$/.test(ticker)) return ticker;

  return ticker || assetId;
}

function sanitizeWidgetId(value: string): string {
  return value.replace(/[^a-zA-Z0-9_]/g, '_');
}

function CompositionRing({ rows, mode, centerLabel }: { rows: PortfolioItem[]; mode: 'account' | 'total'; centerLabel: string }) {
  const gradient = buildConicGradient(rows, mode);
  return (
    <div className="ring-layout">
      <div className="ring-chart" style={{ background: gradient }}>
        <div className="ring-center">
          <span>{mode === 'account' ? '계좌비중' : '전체비중'}</span>
          <strong>{centerLabel}</strong>
        </div>
      </div>
      <div className="ring-legend">
        {rows.slice(0, 6).map((row, index) => (
          <div key={`${row.account_id}-${row.asset_id}`} className="legend-row">
            <span className="legend-dot" style={{ background: chartColor(index) }} />
            <span>{row.asset_name}</span>
            <strong>{formatPercent(mode === 'account' ? row.account_weight : row.total_weight)}</strong>
          </div>
        ))}
      </div>
    </div>
  );
}

function BreakdownList({ rows, mode, onSelect }: { rows: PortfolioItem[]; mode: 'account' | 'total'; onSelect?: (item: PortfolioItem) => void }) {
  return (
    <div className="breakdown-list">
      {rows.map((row) => {
        const weight = mode === 'account' ? row.account_weight : row.total_weight;
        return (
          <button key={`${row.account_id}-${row.asset_id}`} className="breakdown-row" onClick={() => onSelect?.(row)}>
            <div className="bar-label">
              <span>{row.asset_name}</span>
              <strong>{formatPercent(weight)}</strong>
            </div>
            <div className="bar-bg"><div className="bar-fill" style={{ width: `${Math.max(weight * 100, 2)}%` }} /></div>
            <div className="account-subline">
              <span>{formatKRW(row.valuation_amount)}</span>
              <span className={returnClass(row.profit_rate)}>{formatPercent(row.profit_rate)}</span>
            </div>
          </button>
        );
      })}
    </div>
  );
}

function ComparisonBar({ label, value, variant = 'current' }: { label: string; value: number; variant?: 'current' | 'target' }) {
  return (
    <div className="comparison-row">
      <div className="bar-label">
        <span>{label}</span>
        <strong>{formatPercent(value)}</strong>
      </div>
      <div className="bar-bg">
        <div className={`bar-fill ${variant === 'target' ? 'target-fill' : ''}`} style={{ width: `${Math.max(Math.min(value * 100, 100), 2)}%` }} />
      </div>
    </div>
  );
}

function buildConicGradient(rows: PortfolioItem[], mode: 'account' | 'total'): string {
  if (rows.length === 0) return 'conic-gradient(rgba(148, 163, 184, 0.18) 0 100%)';

  let cursor = 0;
  const parts = rows.slice(0, 8).map((row, index) => {
    const raw = Number(mode === 'account' ? row.account_weight : row.total_weight) || 0;
    const size = Math.max(raw * 100, 0);
    const start = cursor;
    const end = cursor + size;
    cursor = end;
    return `${chartColor(index)} ${start}% ${end}%`;
  });

  if (cursor < 100) parts.push(`rgba(148, 163, 184, 0.16) ${cursor}% 100%`);
  return `conic-gradient(${parts.join(', ')})`;
}

function chartColor(index: number): string {
  const colors = ['#22c55e', '#38bdf8', '#f59e0b', '#ef4444', '#a78bfa', '#14b8a6', '#f97316', '#e879f9'];
  return colors[index % colors.length];
}


function HoldingDividendSummary({ item, dividends }: { item: PortfolioItem; dividends: DividendDashboard | null }) {
  const summary = dividends?.by_asset.find((row) => row.account_id === item.account_id && row.asset_id === item.asset_id);

  if (!summary) {
    return (
      <section className="detail-section">
        <div className="section-title">배당 요약</div>
        <div className="empty-inline">등록된 배당 내역이 없다.</div>
      </section>
    );
  }

  return (
    <section className="detail-section">
      <div className="section-title">배당 요약</div>
      <div className="summary-grid">
        <Metric label="누적 배당" value={formatKRW(summary.total_net)} />
        <Metric label="올해 배당" value={formatKRW(summary.this_year_net)} />
        <Metric label="전년 배당" value={formatKRW(summary.last_year_net)} />
        <Metric label="YoY" value={summary.last_year_net > 0 ? formatPercent(summary.yoy_growth_rate) : '-'} className={returnClass(summary.yoy_growth_rate)} />
      </div>
      <p className="muted small">최근 배당일 {summary.latest_date || '-'} · 기록 {summary.record_count}건</p>
    </section>
  );
}


type DividendFilter = 'ALL' | 'CURRENT' | 'OTHER_ACCOUNT' | 'SOLD';

function DividendPage({ dashboard, dividends, api, onReload, onHoldingDetail }: {
  dashboard: DashboardData;
  dividends: DividendDashboard | null;
  api: EdsApi;
  onReload: () => Promise<void>;
  onHoldingDetail: (item: PortfolioItem) => void;
}) {
  const [filter, setFilter] = useState<DividendFilter>('ALL');
  const [refreshing, setRefreshing] = useState(false);
  const [localMessage, setLocalMessage] = useState('');
  const [localError, setLocalError] = useState('');

  const filteredRecords = useMemo(() => {
    const rows = dividends?.records ?? [];
    return rows.filter((row) => dividendFilterMatches(row.holding_status, filter));
  }, [dividends, filter]);

  const filteredAssetRows = useMemo(() => {
    const rows = dividends?.by_asset ?? [];
    return rows.filter((row) => dividendFilterMatches(row.holding_status, filter));
  }, [dividends, filter]);

  const filteredSummary = useMemo(() => buildDividendSummaryFromRecords(filteredRecords), [filteredRecords]);
  const filteredMonthly = useMemo(() => buildMonthlyDividendRows(filteredRecords), [filteredRecords]);
  const recentRows = useMemo(() => filteredRecords.slice().sort((a, b) => String(b.dividend_date).localeCompare(String(a.dividend_date))).slice(0, 30), [filteredRecords]);

  async function refreshDividendData() {
    const confirmed = window.confirm('원본 `6. 배당내역`을 기준으로 App_Dividends를 다시 생성합니다. 기존 App_Dividends는 덮어써질 수 있습니다. 주요 변경 전에는 Google Sheet 사본 백업을 권장합니다. 계속할까요?');
    if (!confirmed) return;
    setRefreshing(true);
    setLocalError('');
    setLocalMessage('');
    try {
      const result = await api.refreshDividends();
      setLocalMessage(`6. 배당내역 이관 완료: imported=${formatNumber(result.imported_count)}건, skipped=${formatNumber(result.skipped_count || 0)}건, error=${formatNumber(result.error_count || 0)}건`);
      await onReload();
    } catch (e) {
      setLocalError(formatErrorForDisplay(e));
    } finally {
      setRefreshing(false);
    }
  }

  function findHolding(row: { account_id: string; asset_id: string }) {
    return dashboard.output.find((item) => item.account_id === row.account_id && item.asset_id === row.asset_id);
  }

  return (
    <section className="page-stack">
      <section className="hero-card dividend-hero">
        <div className="muted">누적 배당금</div>
        <div className="hero-value">{formatKRW(filteredSummary.totalNet)}</div>
        <div className="summary-grid">
          <Metric label="올해 배당" value={formatKRW(filteredSummary.thisYearNet)} />
          <Metric label="최근 12개월" value={formatKRW(filteredSummary.recent12mNet)} />
          <Metric label="월평균" value={formatKRW(filteredSummary.monthlyAverage12m)} />
          <Metric label="기록 수" value={`${formatNumber(filteredSummary.recordCount)}건`} />
        </div>
        <p className="muted small">기준 원장: Google Sheet `6. 배당내역` → App_Dividends 이관 데이터</p>
      </section>

      {localMessage && <div className="notice success">{localMessage}</div>}
      {localError && <div className="notice error">{localError}</div>}

      <section className="card dividend-control-card">
        <div className="section-title">배당 필터</div>
        <div className="filter-row dividend-filter-row">
          <button className={filter === 'ALL' ? 'active' : ''} onClick={() => setFilter('ALL')}>전체</button>
          <button className={filter === 'CURRENT' ? 'active' : ''} onClick={() => setFilter('CURRENT')}>현재 보유</button>
          <button className={filter === 'OTHER_ACCOUNT' ? 'active' : ''} onClick={() => setFilter('OTHER_ACCOUNT')}>다른 계좌 보유</button>
          <button className={filter === 'SOLD' ? 'active' : ''} onClick={() => setFilter('SOLD')}>미보유</button>
        </div>
        <div className="dividend-status-grid">
          <Metric label="현재 보유 종목" value={formatKRW(dividends?.summary.current_holdings_net ?? 0)} />
          <Metric label="다른 계좌 보유" value={formatKRW(dividends?.summary.other_account_holdings_net ?? 0)} />
          <Metric label="미보유 종목" value={formatKRW(dividends?.summary.sold_holdings_net ?? 0)} />
          <Metric label="전체 기록" value={`${formatNumber(dividends?.summary.record_count ?? 0)}건`} />
        </div>
        <button className="ghost" onClick={refreshDividendData} disabled={refreshing}>{refreshing ? '이관 중...' : '6. 배당내역 다시 가져오기'}</button>
        <p className="muted small">원본 `6. 배당내역`을 수정한 뒤 이 버튼으로 App_Dividends를 갱신한다.</p>
      </section>

      <section className="card">
        <div className="section-title">월별 배당 차트</div>
        <DividendMonthlyChart rows={filteredMonthly} />
      </section>

      <section className="card">
        <div className="section-title">계좌별 배당</div>
        <div className="dividend-list">
          {(dividends?.by_account ?? []).map((row) => (
            <div key={row.account_id} className="dividend-row static">
              <div>
                <strong>{row.account_name}</strong>
                <div className="muted small">{row.broker || '-'} · {row.account_type || '-'} · {row.record_count}건</div>
              </div>
              <div className="right">
                <strong>{formatKRW(row.total_net)}</strong>
                <span className="muted small">올해 {formatKRW(row.this_year_net)}</span>
              </div>
            </div>
          ))}
          {(dividends?.by_account ?? []).length === 0 && <div className="empty-inline">계좌별 배당 데이터가 없다.</div>}
        </div>
      </section>

      <section className="card">
        <div className="section-title">종목별 배당 요약</div>
        <div className="dividend-list">
          {filteredAssetRows.slice(0, 40).map((row) => {
            const item = findHolding(row);
            const content = (
              <>
                <div>
                  <strong>{row.asset_name}</strong>
                  <div className="muted small">{row.account_name} · {row.ticker} · {row.record_count}건</div>
                  <span className={`status-pill ${dividendStatusClass(row.holding_status)}`}>{dividendStatusLabel(row.holding_status)}</span>
                </div>
                <div className="right">
                  <strong>{formatKRW(row.total_net)}</strong>
                  <span className="muted small">올해 {formatKRW(row.this_year_net)}</span>
                </div>
              </>
            );
            return item ? (
              <button key={`${row.account_id}-${row.asset_id}`} className="dividend-row" onClick={() => onHoldingDetail(item)}>{content}</button>
            ) : (
              <div key={`${row.account_id}-${row.asset_id}`} className="dividend-row static">{content}</div>
            );
          })}
          {filteredAssetRows.length === 0 && <div className="empty-inline">조건에 맞는 배당 내역이 없다.</div>}
        </div>
      </section>

      <section className="card">
        <div className="section-title">최근 배당 내역</div>
        <div className="dividend-list">
          {recentRows.map((row) => (
            <div key={row.dividend_id} className="dividend-row static">
              <div>
                <strong>{row.asset_name}</strong>
                <div className="muted small">{row.dividend_date} · {row.account_name} · {row.ticker}</div>
                <span className={`status-pill ${dividendStatusClass(row.holding_status)}`}>{dividendStatusLabel(row.holding_status)}</span>
              </div>
              <div className="right">
                <strong>{formatKRW(row.net_amount_krw ?? row.net_amount)}</strong>
                <span className="muted small">원본 {row.source_row ? `${row.source_row}행` : '-'}</span>
              </div>
            </div>
          ))}
          {recentRows.length === 0 && <div className="empty-inline">최근 배당 기록이 없다.</div>}
        </div>
      </section>
    </section>
  );
}

function dividendFilterMatches(status: string | undefined, filter: DividendFilter): boolean {
  if (filter === 'ALL') return true;
  if (filter === 'CURRENT') return status === 'current';
  if (filter === 'OTHER_ACCOUNT') return status === 'other_account_current';
  if (filter === 'SOLD') return status === 'sold';
  return true;
}

function dividendStatusLabel(status?: string): string {
  if (status === 'current') return '현재 보유';
  if (status === 'other_account_current') return '다른 계좌 보유';
  if (status === 'sold') return '미보유';
  return '상태 미확인';
}

function dividendStatusClass(status?: string): string {
  if (status === 'current') return 'status-current';
  if (status === 'other_account_current') return 'status-other';
  if (status === 'sold') return 'status-sold';
  return 'status-unknown';
}

function buildDividendSummaryFromRecords(records: NonNullable<DividendDashboard['records']>) {
  const now = new Date();
  const thisYear = now.getFullYear();
  const start12 = new Date(now.getFullYear(), now.getMonth() - 11, 1);
  const totalNet = records.reduce((sum, row) => sum + Number(row.net_amount_krw ?? row.net_amount ?? 0), 0);
  const thisYearNet = records.filter((row) => Number(row.year ?? String(row.dividend_date).slice(0, 4)) === thisYear).reduce((sum, row) => sum + Number(row.net_amount_krw ?? row.net_amount ?? 0), 0);
  const recent12mNet = records.filter((row) => {
    const d = parseDividendDate(row.dividend_date);
    return d ? d >= start12 : false;
  }).reduce((sum, row) => sum + Number(row.net_amount_krw ?? row.net_amount ?? 0), 0);
  return {
    totalNet,
    thisYearNet,
    recent12mNet,
    monthlyAverage12m: recent12mNet / 12,
    recordCount: records.length,
  };
}

function buildMonthlyDividendRows(records: NonNullable<DividendDashboard['records']>) {
  const map = new Map<string, { month_key: string; total_net: number; record_count: number }>();
  records.forEach((row) => {
    const date = String(row.dividend_date || '');
    const key = /^\d{4}-\d{2}/.test(date) ? date.slice(0, 7) : `${row.year}-${String(row.month || 0).padStart(2, '0')}`;
    if (!map.has(key)) map.set(key, { month_key: key, total_net: 0, record_count: 0 });
    const item = map.get(key)!;
    item.total_net += Number(row.net_amount_krw ?? row.net_amount ?? 0);
    item.record_count += 1;
  });
  return Array.from(map.values()).sort((a, b) => a.month_key.localeCompare(b.month_key)).slice(-18);
}

function DividendMonthlyChart({ rows }: { rows: { month_key: string; total_net: number; record_count: number }[] }) {
  if (rows.length === 0) return <div className="empty-inline">월별 배당 데이터가 없다.</div>;
  const max = Math.max(...rows.map((row) => Number(row.total_net || 0)), 1);
  return (
    <div className="monthly-dividend-chart">
      {rows.map((row) => (
        <div className="monthly-bar-item" key={row.month_key}>
          <div className="monthly-bar-value">{formatKRW(row.total_net)}</div>
          <div className="monthly-bar-track">
            <div className="monthly-bar-fill" style={{ height: `${Math.max((row.total_net / max) * 100, 3)}%` }} />
          </div>
          <div className="monthly-bar-label">{formatMonthLabel(row.month_key)}</div>
        </div>
      ))}
    </div>
  );
}

function formatMonthLabel(monthKey: string) {
  const [year, month] = String(monthKey).split('-');
  if (!year || !month) return monthKey;
  return `${year.slice(2)}.${month}`;
}

function parseDividendDate(value: string) {
  const text = String(value || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  return new Date(Number(text.slice(0, 4)), Number(text.slice(5, 7)) - 1, Number(text.slice(8, 10)));
}

function HistoryPage() {
  return (
    <section className="card empty-card">
      <h2>내역</h2>
      <p>1차 MVP에서는 계좌총액/배당/입출금 내역 입력을 제외했다.</p>
      <p className="muted">다음 단계에서 App_AccountValues, App_Dividends를 연결하면 된다.</p>
    </section>
  );
}

function SettingsPage({
  config,
  updateConfig,
  resetConfig,
  onPing,
  onRefreshOutput,
  onRefreshPrices,
  onRefreshPricesToMainSheet,
  onRefreshKrxDailyCharts,
  autoPriceRefreshMinutes,
  onAutoPriceRefreshMinutesChange,
  autoPriceMarketOnly,
  onAutoPriceMarketOnlyChange,
  autoPriceStatus,
  onPreviewSync,
  onSyncToMain,
  syncFastMode,
  onSyncFastModeChange,
  onInstall,
  installAvailable,
  isStandalone,
  appStatus,
  refreshLogs,
  onFetchStatus,
  onClearChartCache,
  onClearTokenOnly,
  onClearAllLocalData,
}: {
  config: AppConfig;
  updateConfig: (config: AppConfig) => void;
  resetConfig: () => void;
  onPing: () => void;
  onRefreshOutput: () => void;
  onRefreshPrices: () => void;
  onRefreshPricesToMainSheet: () => void;
  onRefreshKrxDailyCharts: () => void;
  autoPriceRefreshMinutes: number;
  onAutoPriceRefreshMinutesChange: (minutes: number) => void;
  autoPriceMarketOnly: boolean;
  onAutoPriceMarketOnlyChange: (enabled: boolean) => void;
  autoPriceStatus: string;
  onPreviewSync: () => void;
  onSyncToMain: () => void;
  syncFastMode: boolean;
  onSyncFastModeChange: (enabled: boolean) => void;
  onInstall: () => void;
  installAvailable: boolean;
  isStandalone: boolean;
  appStatus: AppStatus | null;
  refreshLogs: RefreshLogs;
  onFetchStatus: () => void;
  onClearChartCache: () => void;
  onClearTokenOnly: () => void;
  onClearAllLocalData: () => void;
}) {
  const [apiUrl, setApiUrl] = useState(config.apiUrl);
  const [token, setToken] = useState(config.token);

  useEffect(() => {
    setApiUrl(config.apiUrl);
    setToken(config.token);
  }, [config.apiUrl, config.token]);

  const [advancedOpen, setAdvancedOpen] = useState(false);

  return (
    <section className="page-stack">
      {/* === 항상 노출: API 설정 === */}
      <section className="card form-card">
        <div className="section-title">API 설정</div>
        <label>
          Apps Script 웹앱 URL
          <input value={apiUrl} onChange={(e) => setApiUrl(e.target.value)} placeholder="https://script.google.com/macros/s/.../exec" />
        </label>
        <label>
          API token
          <input type="password" value={token} onChange={(e) => setToken(e.target.value)} placeholder="App_Settings api_token" autoComplete="off" />
        </label>
        <div className="button-row">
          <button onClick={() => updateConfig({ apiUrl, token })}>설정 저장</button>
          <button className="ghost" onClick={resetConfig}>URL/token 초기화</button>
        </div>
      </section>

      {/* === 항상 노출: 현재가 자동 갱신 === */}
      <section className="card action-card">
        <div className="section-title">현재가 자동 갱신</div>
        <div className="segmented-control wrap">
          {[0, 1, 3, 5, 10].map((minutes) => (
            <button
              key={minutes}
              className={autoPriceRefreshMinutes === minutes ? 'active' : ''}
              onClick={() => onAutoPriceRefreshMinutesChange(minutes)}
            >
              {minutes === 0 ? '끄기' : `${minutes}분`}
            </button>
          ))}
        </div>
        <label className="checkbox-line">
          <input
            type="checkbox"
            checked={autoPriceMarketOnly}
            onChange={(e) => onAutoPriceMarketOnlyChange(e.target.checked)}
          />
          시장 시간에만 자동 갱신
        </label>
        <div className="kv-list compact">
          <div><span>자동 갱신 상태</span><strong>{autoPriceStatus}</strong></div>
        </div>
        <p className="muted small">앱이 열려 있고 화면이 활성 상태일 때만 App_Prices/App_Output을 갱신한다. 원본 `2. 종목현황`은 자동으로 수정하지 않는다.</p>
      </section>

      {/* === 항상 노출: 연결/동기화 === */}
      <section className="card action-card">
        <div className="section-title">연결/동기화</div>
        <button className="ghost" onClick={onPing}>ping 테스트</button>
        <button className="ghost" onClick={onRefreshOutput}>App_Output 재계산</button>
        <button className="ghost" onClick={onRefreshPrices}>현재가 갱신(KIS)</button>
        <button className="ghost" onClick={onRefreshKrxDailyCharts}>국내 일봉 차트 갱신(KIS)</button>
        <button className="ghost" onClick={onRefreshPricesToMainSheet}>원본 시트 가격 반영(KIS)</button>
        <button className="ghost" onClick={onPreviewSync}>원장 반영 미리보기</button>

        <div className="section-subtitle">원장 동기화 모드</div>
        <div className="segmented-control">
          <button
            className={syncFastMode ? 'active' : ''}
            onClick={() => onSyncFastModeChange(true)}
            title="인덱싱/검증 생략, Bulk Write 직행 — 평상시 권장"
          >
            ⚡ 고속 모드
          </button>
          <button
            className={!syncFastMode ? 'active' : ''}
            onClick={() => onSyncFastModeChange(false)}
            title="전체 인덱싱 + 정합성 검증 후 반영 — 데이터 이상 의심 시 사용"
          >
            🛡️ 안전 모드
          </button>
        </div>
        <p className="muted small">
          {syncFastMode
            ? '⚡ 고속 모드: 인덱싱/검증 생략, 2차원 배열 Bulk Write 직행. 평상시 권장.'
            : '🛡️ 안전 모드: 전체 인덱싱 + 정합성 검증 후 반영. 데이터 이상 의심 시 사용.'}
        </p>
        <button className="danger" onClick={onSyncToMain}>기존 원장 실제 반영</button>
        <div className="warning-box">
          백업 권고: `기존 원장 실제 반영`, `6. 배당내역 다시 가져오기`, 대량 import/sync 전에는 Google Sheet 사본을 먼저 만들어두는 것이 안전하다. 현재가 갱신은 App_Prices/App_Output만 갱신한다. 국내 일봉 차트 갱신은 App_ChartPrices만 갱신한다. 원본 시트 가격 반영(KIS)은 2. 종목현황의 국내 K/M열, 미국 L/M열을 값으로 수정하므로 실행 전 백업을 권장한다.
        </div>
      </section>

      {/* === 항상 노출: 캐시/설정 삭제 === */}
      <section className="card action-card">
        <div className="section-title">캐시/설정 삭제</div>
        <button className="ghost" onClick={onClearChartCache}>차트 캐시 삭제</button>
        <button className="ghost" onClick={onClearTokenOnly}>API token만 삭제</button>
        <button className="danger" onClick={onClearAllLocalData}>이 브라우저의 앱 설정 전체 삭제</button>
        <p className="muted small">차트 캐시는 브라우저에 저장된 일/주/月 차트 데이터다. 삭제해도 원장/시트 데이터는 삭제되지 않는다.</p>
      </section>

      {/* === Advanced Settings 아코디언 (기본 접힘) === */}
      <section className="card advanced-settings-card">
        <button
          className="advanced-settings-toggle"
          onClick={() => setAdvancedOpen((prev) => !prev)}
          aria-expanded={advancedOpen}
        >
          <span>⚙️ Advanced Settings</span>
          <span className="advanced-settings-chevron">{advancedOpen ? '▲' : '▼'}</span>
        </button>

        {advancedOpen && (
          <div className="advanced-settings-body">
            <section className="card form-card">
              <div className="section-title">버전/상태</div>
              <div className="status-grid">
                <div className="status-pill">PWA v{APP_VERSION}</div>
                <div className="status-pill">Apps Script {appStatus?.apps_script_version ? `v${appStatus.apps_script_version}` : '미확인'}</div>
                <div className="status-pill">실행 모드 {isStandalone ? '설치 앱' : '브라우저'}</div>
                <div className="status-pill">Token {config.token ? '저장됨' : '미설정'}</div>
              </div>
              <div className="kv-list">
                <div><span>API URL</span><strong>{config.apiUrl ? maskApiUrl(config.apiUrl) : '미설정'}</strong></div>
                <div><span>Spreadsheet</span><strong>{appStatus?.spreadsheet_name || '-'}</strong></div>
                <div><span>API 시각</span><strong>{formatLogTime(appStatus?.api_time)}</strong></div>
              </div>
              <button className="ghost" onClick={onFetchStatus}>상태 확인</button>
              <p className="muted small">Apps Script 버전은 `getAppStatus` action이 배포된 뒤 표시된다.</p>
            </section>

            <section className="card action-card">
              <div className="section-title">데이터 새로고침 로그</div>
              <div className="kv-list">
                <div><span>홈/포트폴리오</span><strong>{formatLogTime(refreshLogs.getDashboard)}</strong></div>
                <div><span>배당 대시보드</span><strong>{formatLogTime(refreshLogs.getDividendDashboard)}</strong></div>
                <div><span>개별 종목 차트 조회</span><strong>{formatLogTime(refreshLogs.getChartData)}</strong></div>
                <div><span>종목 차트 일괄 갱신</span><strong>{formatLogTime(refreshLogs.refreshKrxDailyCharts)}</strong></div>
                <div><span>현재가</span><strong>{formatLogTime(refreshLogs.refreshKrxPrices)}</strong></div>
                <div><span>원본 시트 가격 반영</span><strong>{formatLogTime(refreshLogs.refreshKrxPricesToMainSheet)}</strong></div>
              </div>
            </section>

            <section className="card action-card">
              <div className="section-title">PWA 설치</div>
              <p className="muted small">
                {isStandalone
                  ? '현재 앱 설치 모드로 실행 중이다.'
                  : '현재 브라우저 탭 모드다. 안드로이드 크롬에서는 메뉴 > 홈 화면에 추가로 설치할 수 있다.'}
              </p>
              <button className="ghost" onClick={onInstall} disabled={isStandalone && !installAvailable}>
                {installAvailable ? '앱 설치하기' : '설치 안내 보기'}
              </button>
            </section>
          </div>
        )}
      </section>
    </section>
  );
}

function EmptyState({ onGoSettings, onRefresh }: { onGoSettings: () => void; onRefresh: () => void }) {
  return (
    <section className="card empty-card">
      <h2>데이터가 없다</h2>
      <p>설정에서 Apps Script 웹앱 URL과 token을 입력한 뒤 새로고침하면 된다.</p>
      <div className="button-row">
        <button onClick={onGoSettings}>설정 열기</button>
        <button className="ghost" onClick={onRefresh}>새로고침</button>
      </div>
    </section>
  );
}

function Metric({ label, value, className = '' }: { label: string; value: string; className?: string }) {
  return (
    <div className="metric">
      <span>{label}</span>
      <strong className={className}>{value}</strong>
    </div>
  );
}


function formatErrorForDisplay(error: unknown): string {
  const text = error instanceof Error ? error.message : String(error);
  if (text.includes('invalid_token') || text.includes('token')) return text;
  if (text.includes('unknown_action')) return text;
  if (text.includes('Failed to fetch')) return 'API 호출 실패: 네트워크, CORS, Apps Script 배포 권한을 확인하세요.';
  return text;
}

function maskApiUrl(url: string): string {
  if (!url) return '';
  try {
    const parsed = new URL(url);
    const parts = parsed.pathname.split('/').filter(Boolean);
    const idx = parts.indexOf('s');
    const deployId = idx >= 0 ? parts[idx + 1] : '';
    const maskedId = deployId ? `${deployId.slice(0, 8)}…${deployId.slice(-6)}` : '…';
    return `${parsed.origin}/macros/s/${maskedId}/exec`;
  } catch {
    return url.length > 36 ? `${url.slice(0, 22)}…${url.slice(-10)}` : url;
  }
}

function formatLogTime(value?: string): string {
  if (!value) return '-';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString('ko-KR', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
}

function summarizeResult(result: unknown): string {
  if (!result || typeof result !== 'object') return String(result ?? '');
  const data = result as Record<string, unknown>;

  if ('preview_summary' in data) return JSON.stringify(data.preview_summary);
  if ('sync_result_summary' in data) return JSON.stringify(data.sync_result_summary);
  if ('refreshed_count' in data) return `success=${data.success_count ?? 0}, error=${data.error_count ?? 0}, updated=${data.updated_price_count ?? 0}, output=${data.output_count ?? '-'}`;
  if ('row_count' in data && 'target_count' in data) return `target=${data.target_count ?? 0}, rows=${data.row_count ?? 0}, success=${data.success_count ?? 0}, error=${data.error_count ?? 0}`;
  if ('imported_count' in data) return `imported=${data.imported_count ?? 0}, skipped=${data.skipped_count ?? 0}, error=${data.error_count ?? 0}`;
  if ('apps_script_version' in data) return `Apps Script v${data.apps_script_version ?? '-'}, sheet=${data.spreadsheet_name ?? '-'}`;
  if ('output_count' in data) return `output_count=${data.output_count}`;
  if ('timestamp' in data) return String(data.timestamp);

  return JSON.stringify(data);
}

export default App;
