/**
 * Excel Connector Types
 * Data structures for Excel integration and data export
 */

import { MarketDataPoint, HistoricalDataPoint } from './bloomberg.types';

export interface ExcelConfig {
  autoRefresh?: boolean;
  refreshInterval?: number; // milliseconds
  enableFormatting?: boolean;
  includeHeaders?: boolean;
  sheetName?: string;
}

export interface ExcelWorksheetConfig {
  name: string;
  data: any[];
  headers?: string[];
  startRow?: number;
  startColumn?: number;
  formatting?: ExcelFormattingOptions;
}

export interface ExcelFormattingOptions {
  headerStyle?: {
    font?: { bold?: boolean; color?: string; size?: number };
    fill?: { type: string; pattern: string; fgColor?: string };
    alignment?: { horizontal?: string; vertical?: string };
  };
  dataStyle?: {
    numberFormat?: string;
    alignment?: { horizontal?: string; vertical?: string };
  };
  conditionalFormatting?: ConditionalFormat[];
  columnWidths?: number[];
}

export interface ConditionalFormat {
  type: 'colorScale' | 'dataBar' | 'iconSet' | 'expression';
  column: string;
  rule: any;
}

export interface ExcelExportOptions {
  fileName: string;
  filePath?: string;
  worksheets: ExcelWorksheetConfig[];
  includeCharts?: boolean;
  protection?: {
    password?: string;
    lockStructure?: boolean;
  };
}

export interface LiveDataRange {
  workbook: string;
  worksheet: string;
  range: string;
  securities: string[];
  fields: string[];
  updateFrequency: number; // milliseconds
}

export interface ExcelDataTransform {
  source: MarketDataPoint[] | HistoricalDataPoint[];
  transformType: 'pivot' | 'aggregate' | 'filter' | 'sort';
  options?: any;
}

export interface ExcelConnectionStatus {
  connected: boolean;
  lastUpdate?: Date;
  activeRanges: number;
  errors?: string[];
}
