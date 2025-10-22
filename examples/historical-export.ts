/**
 * Example: Export historical Bloomberg data to Excel
 */

import { createBloombergExcelConnector, BLOOMBERG_FIELDS } from '../src/index';
import * as dotenv from 'dotenv';

dotenv.config({ path: './config/.env' });

async function historicalExport() {
  const { bloomberg, excel } = createBloombergExcelConnector({
    apiKey: process.env.BLOOMBERG_API_KEY!,
    apiSecret: process.env.BLOOMBERG_API_SECRET!,
    serverHost: process.env.BLOOMBERG_SERVER_HOST!,
    serverPort: parseInt(process.env.BLOOMBERG_SERVER_PORT!),
  });

  try {
    console.log('Exporting historical data to Excel...');

    // Define security
    const security = {
      ticker: 'AAPL',
      exchange: 'NASDAQ',
      securityType: 'EQUITY' as const,
    };

    // Define date range (last 90 days)
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 90);

    // Define fields
    const fields = [
      BLOOMBERG_FIELDS.CLOSE,
      BLOOMBERG_FIELDS.VOLUME,
      BLOOMBERG_FIELDS.HIGH,
      BLOOMBERG_FIELDS.LOW,
    ];

    // Export historical data
    const filePath = await excel.exportHistoricalData(
      security,
      fields,
      startDate,
      endDate,
      'historical-data.xlsx',
      process.env.EXCEL_OUTPUT_PATH
    );

    console.log(`✓ Historical export complete: ${filePath}`);
  } catch (error) {
    console.error('Export failed:', error);
  } finally {
    bloomberg.disconnect();
    excel.disconnect();
  }
}

// Run example
historicalExport();
