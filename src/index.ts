/**
 * Bloomberg Terminal Excel Connector
 * Main entry point for the integration library
 */

// Export core classes
export { BloombergClient } from './connectors/bloomberg-client';
export { ExcelConnector } from './connectors/excel/excel-connector';

// Export types
export * from './types/bloomberg.types';
export * from './types/excel.types';

// Export utilities
export { Logger } from './utils/logger';
export { DataFormatter } from './utils/data-formatter';

// Re-export for convenience
import { BloombergClient } from './connectors/bloomberg-client';
import { ExcelConnector } from './connectors/excel/excel-connector';
import { BloombergConfig, ExcelConfig } from './types/bloomberg.types';

/**
 * Create a configured Bloomberg Terminal Excel integration
 */
export function createBloombergExcelConnector(
  bloombergConfig: BloombergConfig,
  excelConfig?: ExcelConfig
): { bloomberg: BloombergClient; excel: ExcelConnector } {
  const bloomberg = new BloombergClient(bloombergConfig);
  const excel = new ExcelConnector(bloomberg, excelConfig);

  return { bloomberg, excel };
}

/**
 * Version information
 */
export const VERSION = '1.0.0';

/**
 * Default configurations
 */
export const DEFAULT_BLOOMBERG_CONFIG: Partial<BloombergConfig> = {
  timeout: 30000,
  enableWebSocket: true,
};

export const DEFAULT_EXCEL_CONFIG: ExcelConfig = {
  autoRefresh: false,
  refreshInterval: 5000,
  enableFormatting: true,
  includeHeaders: true,
  sheetName: 'Market Data',
};
