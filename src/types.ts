export type ApiResponse<T> = {
  ok: boolean;
  action?: string;
  timestamp?: string;
  data?: T;
  error?: string;
  message?: string;
};

export type DashboardSummary = {
  total_valuation: number;
  total_invested: number;
  total_profit: number;
  total_profit_rate: number;
  account_count: number;
  holding_count: number;
};

export type AccountSummary = {
  account_id: string;
  account_name: string;
  broker: string;
  account_type: string;
  valuation_amount: number;
  invested_amount: number;
  profit_amount: number;
  profit_rate: number;
  weight: number;
  holding_count: number;
};

export type PortfolioOutputRow = {
  account_id: string;
  account_name: string;
  broker: string;
  account_type: string;
  asset_id: string;
  ticker: string | number;
  asset_name: string;
  quantity: number;
  avg_price: number;
  price: number;
  invested_amount: number;
  valuation_amount: number;
  profit_amount: number;
  profit_rate: number;
  account_weight: number;
  total_weight: number;
  target_weight_account: number;
  target_gap_rate: number;
  target_gap_amount: number;
  currency: string;
  price_source: string;
  price_fetched_at: string;
  updated_at: string;
  change_amount?: number;
  change_rate?: number;
};

export type HoldingRow = {
  holding_id: string;
  account_id: string;
  asset_id: string;
  ticker: string | number;
  asset_name: string;
  quantity: number;
  avg_price: number;
  target_weight_account: number;
  memo: string;
  enabled: string;
  created_at: string;
  updated_at: string;
};

export type DashboardData = {
  summary: DashboardSummary;
  accounts: AccountSummary[];
  top_holdings: PortfolioOutputRow[];
  output: PortfolioOutputRow[];
};

export type AppConfig = {
  apiUrl: string;
  token: string;
};

export type PortfolioItem = PortfolioOutputRow & {
  holding_id?: string;
};

export type ChartPoint = {
  date: string;
  open: number;
  high: number;
  low: number;
  close: number;
  volume: number;
  source?: string;
  fetched_at?: string;
  main_sheet_updated_count?: number;
};

export type ChartData = {
  asset_id: string;
  interval: 'D' | 'W' | 'M';
  count: number;
  items: ChartPoint[];
  message?: string;
};

export type DividendRecord = {
  dividend_id: string;
  source_sheet?: string;
  source_row?: number;
  dividend_date: string;
  year?: number;
  month?: number;
  day?: number;
  broker?: string;
  account_id: string;
  account_name?: string;
  account_type?: string;
  asset_id: string;
  market?: string;
  ticker: string | number;
  asset_name: string;
  krw_amount?: number;
  foreign_amount?: number;
  gross_amount: number;
  tax_amount: number;
  net_amount: number;
  net_amount_krw?: number;
  currency: string;
  holding_status?: 'current' | 'other_account_current' | 'sold' | string;
  current_holding_id?: string;
  memo?: string;
  enabled?: string;
  created_at?: string;
  updated_at?: string;
};

export type DividendAssetSummary = {
  account_id: string;
  account_name: string;
  broker?: string;
  account_type?: string;
  asset_id: string;
  ticker: string | number;
  asset_name: string;
  holding_status?: 'current' | 'other_account_current' | 'sold' | string;
  total_gross: number;
  total_tax: number;
  total_net: number;
  total_net_krw?: number;
  this_year_net: number;
  last_year_net: number;
  yoy_growth_rate: number;
  record_count: number;
  latest_date: string;
};

export type DividendAccountSummary = {
  account_id: string;
  account_name: string;
  broker?: string;
  account_type?: string;
  total_net: number;
  total_net_krw?: number;
  this_year_net: number;
  last_year_net: number;
  yoy_growth_rate: number;
  record_count: number;
};

export type DividendMonthlySummary = {
  month_key: string;
  year: number;
  month: number;
  total_net: number;
  current_net: number;
  other_account_net: number;
  sold_net: number;
  record_count: number;
};

export type DividendDashboard = {
  summary: {
    total_gross: number;
    total_tax: number;
    total_net: number;
    this_year_net: number;
    last_year_net: number;
    yoy_growth_rate: number;
    record_count: number;
    total_net_krw?: number;
    recent_12m_net?: number;
    monthly_average_12m?: number;
    current_holdings_net?: number;
    other_account_holdings_net?: number;
    sold_holdings_net?: number;
  };
  by_asset: DividendAssetSummary[];
  by_account: DividendAccountSummary[];
  monthly?: DividendMonthlySummary[];
  recent: DividendRecord[];
  records: DividendRecord[];
};


export type RefreshLogs = {
  getDashboard?: string;
  getDividendDashboard?: string;
  getChartData?: string;
  refreshKrxPrices?: string;
  refreshKrxPricesToMainSheet?: string;
  refreshKrxDailyCharts?: string;
  refreshOutput?: string;
  previewSync?: string;
  syncToMain?: string;
  getAppStatus?: string;
  ping?: string;
  [key: string]: string | undefined;
};

export type AppStatus = {
  app?: string;
  pwa_version?: string;
  apps_script_version?: string;
  spreadsheet_name?: string;
  spreadsheet_id?: string;
  api_time?: string;
  timezone?: string;
  sheets?: Record<string, { last_row?: number; last_column?: number; updated_at?: string }>;
  settings?: Record<string, unknown>;
};

export type PriceRefreshResult = {
  refreshed_count: number;
  success_count: number;
  error_count: number;
  skipped_count: number;
  updated_price_count: number;
  output_count?: number;
  source?: string;
  fetched_at?: string;
  main_sheet_updated_count?: number;
  results?: Array<{
    asset_id: string;
    ticker: string | number;
    status: string;
    price?: number;
    change_amount?: number;
    change_rate?: number;
    message?: string;
  }>;
};
