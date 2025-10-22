/**
 * Bloomberg Terminal Integration Types
 * Data structures for market data, securities, and real-time feeds
 */

export interface BloombergConfig {
  apiKey: string;
  apiSecret: string;
  serverHost: string;
  serverPort: number;
  timeout?: number;
  enableWebSocket?: boolean;
}

export interface SecurityIdentifier {
  ticker: string;
  exchange?: string;
  securityType: 'EQUITY' | 'BOND' | 'COMMODITY' | 'CURRENCY' | 'INDEX' | 'FUTURE' | 'OPTION';
  isin?: string;
  cusip?: string;
  sedol?: string;
}

export interface MarketDataPoint {
  security: SecurityIdentifier;
  timestamp: Date;
  bid?: number;
  ask?: number;
  last: number;
  volume: number;
  open?: number;
  high?: number;
  low?: number;
  close?: number;
  vwap?: number;
  change?: number;
  changePercent?: number;
}

export interface HistoricalDataRequest {
  security: SecurityIdentifier;
  fields: string[];
  startDate: Date;
  endDate: Date;
  period?: 'DAILY' | 'WEEKLY' | 'MONTHLY' | 'QUARTERLY' | 'YEARLY';
}

export interface HistoricalDataPoint {
  date: Date;
  [field: string]: any;
}

export interface ReferenceDataRequest {
  securities: SecurityIdentifier[];
  fields: string[];
}

export interface ReferenceDataResponse {
  security: SecurityIdentifier;
  data: Record<string, any>;
  errors?: string[];
}

export interface BloombergSubscription {
  id: string;
  securities: SecurityIdentifier[];
  fields: string[];
  callback: (data: MarketDataPoint) => void;
}

export interface BloombergError {
  code: string;
  message: string;
  security?: SecurityIdentifier;
  timestamp: Date;
}

export interface BloombergResponse<T> {
  success: boolean;
  data?: T;
  error?: BloombergError;
}

// Common Bloomberg field codes
export const BLOOMBERG_FIELDS = {
  LAST_PRICE: 'PX_LAST',
  BID: 'PX_BID',
  ASK: 'PX_ASK',
  VOLUME: 'PX_VOLUME',
  OPEN: 'PX_OPEN',
  HIGH: 'PX_HIGH',
  LOW: 'PX_LOW',
  CLOSE: 'PX_CLOSE',
  VWAP: 'VWAP',
  MARKET_CAP: 'CUR_MKT_CAP',
  PE_RATIO: 'PE_RATIO',
  DIVIDEND_YIELD: 'DVD_YIELD',
  BETA: 'BETA',
  EPS: 'EARNINGS_PER_SH',
  COMPANY_NAME: 'NAME',
  COUNTRY: 'COUNTRY_ISO',
  CURRENCY: 'CRNCY',
} as const;
