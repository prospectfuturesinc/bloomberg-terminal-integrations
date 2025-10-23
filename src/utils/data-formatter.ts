/**
 * Data formatting utilities for Excel export
 */

import { MarketDataPoint, HistoricalDataPoint } from '../types/bloomberg.types';

export class DataFormatter {
  /**
   * Format number with specified decimal places
   */
  static formatNumber(value: number, decimals: number = 2): string {
    return value.toFixed(decimals);
  }

  /**
   * Format percentage
   */
  static formatPercent(value: number, decimals: number = 2): string {
    return `${(value * 100).toFixed(decimals)}%`;
  }

  /**
   * Format currency
   */
  static formatCurrency(value: number, currency: string = 'USD', decimals: number = 2): string {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency,
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    }).format(value);
  }

  /**
   * Format date
   */
  static formatDate(date: Date, format: 'short' | 'long' | 'iso' = 'short'): string {
    switch (format) {
      case 'short':
        return date.toLocaleDateString('en-US');
      case 'long':
        return date.toLocaleDateString('en-US', {
          year: 'numeric',
          month: 'long',
          day: 'numeric',
        });
      case 'iso':
        return date.toISOString();
      default:
        return date.toLocaleDateString();
    }
  }

  /**
   * Format time
   */
  static formatTime(date: Date, includeSeconds: boolean = false): string {
    return date.toLocaleTimeString('en-US', {
      hour: '2-digit',
      minute: '2-digit',
      second: includeSeconds ? '2-digit' : undefined,
    });
  }

  /**
   * Format market data point for display
   */
  static formatMarketData(data: MarketDataPoint): Record<string, any> {
    return {
      Ticker: data.security.ticker,
      Exchange: data.security.exchange || 'N/A',
      Last: this.formatNumber(data.last),
      Bid: data.bid ? this.formatNumber(data.bid) : 'N/A',
      Ask: data.ask ? this.formatNumber(data.ask) : 'N/A',
      Volume: data.volume.toLocaleString(),
      Change: data.change ? this.formatNumber(data.change) : 'N/A',
      'Change %': data.changePercent ? this.formatPercent(data.changePercent / 100) : 'N/A',
      Timestamp: this.formatDate(data.timestamp) + ' ' + this.formatTime(data.timestamp),
    };
  }

  /**
   * Convert market data array to Excel-friendly format
   */
  static marketDataToExcelRows(data: MarketDataPoint[]): any[][] {
    return data.map((point) => {
      const formatted = this.formatMarketData(point);
      return Object.values(formatted);
    });
  }

  /**
   * Get market data headers
   */
  static getMarketDataHeaders(): string[] {
    return [
      'Ticker',
      'Exchange',
      'Last',
      'Bid',
      'Ask',
      'Volume',
      'Change',
      'Change %',
      'Timestamp',
    ];
  }

  /**
   * Calculate change statistics
   */
  static calculateChange(current: number, previous: number): { change: number; changePercent: number } {
    const change = current - previous;
    const changePercent = (change / previous) * 100;
    return { change, changePercent };
  }

  /**
   * Aggregate historical data
   */
  static aggregateHistoricalData(
    data: HistoricalDataPoint[],
    period: 'weekly' | 'monthly' | 'quarterly'
  ): HistoricalDataPoint[] {
    // Implementation for data aggregation
    // This is a placeholder - real implementation would group and aggregate data
    return data;
  }

  /**
   * Calculate moving average
   */
  static calculateMovingAverage(data: number[], period: number): number[] {
    const result: number[] = [];
    for (let i = 0; i < data.length; i++) {
      if (i < period - 1) {
        result.push(NaN);
      } else {
        const sum = data.slice(i - period + 1, i + 1).reduce((a, b) => a + b, 0);
        result.push(sum / period);
      }
    }
    return result;
  }

  /**
   * Sanitize data for Excel export
   */
  static sanitizeForExcel(value: any): any {
    if (value === null || value === undefined) {
      return '';
    }
    if (typeof value === 'string') {
      // Remove any Excel formula characters
      return value.replace(/^[=+\-@]/, "'$&");
    }
    return value;
  }
}
