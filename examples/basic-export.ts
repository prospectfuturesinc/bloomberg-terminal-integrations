/**
 * Basic Example: Export Bloomberg market data to Excel
 */

import { createBloombergExcelConnector, BLOOMBERG_FIELDS } from '../src/index';
import * as dotenv from 'dotenv';

dotenv.config({ path: './config/.env' });

async function basicExport() {
  // Create connector
  const { bloomberg, excel } = createBloombergExcelConnector({
    apiKey: process.env.BLOOMBERG_API_KEY!,
    apiSecret: process.env.BLOOMBERG_API_SECRET!,
    serverHost: process.env.BLOOMBERG_SERVER_HOST!,
    serverPort: parseInt(process.env.BLOOMBERG_SERVER_PORT!),
    enableWebSocket: true,
  });

  try {
    console.log('Exporting market data to Excel...');

    // Define securities to export
    const securities = [
      { ticker: 'AAPL', exchange: 'NASDAQ', securityType: 'EQUITY' as const },
      { ticker: 'MSFT', exchange: 'NASDAQ', securityType: 'EQUITY' as const },
      { ticker: 'GOOGL', exchange: 'NASDAQ', securityType: 'EQUITY' as const },
      { ticker: 'TSLA', exchange: 'NASDAQ', securityType: 'EQUITY' as const },
    ];

    // Define fields to retrieve
    const fields = [
      BLOOMBERG_FIELDS.LAST_PRICE,
      BLOOMBERG_FIELDS.BID,
      BLOOMBERG_FIELDS.ASK,
      BLOOMBERG_FIELDS.VOLUME,
      BLOOMBERG_FIELDS.MARKET_CAP,
    ];

    // Export to Excel
    const filePath = await excel.exportMarketData(
      securities,
      fields,
      'market-data.xlsx',
      process.env.EXCEL_OUTPUT_PATH
    );

    console.log(`✓ Export complete: ${filePath}`);
  } catch (error) {
    console.error('Export failed:', error);
  } finally {
    // Clean up
    bloomberg.disconnect();
    excel.disconnect();
  }
}

// Run example
basicExport();
