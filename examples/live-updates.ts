/**
 * Advanced Example: Live updating Excel workbook with real-time data
 */

import { createBloombergExcelConnector, BLOOMBERG_FIELDS } from '../src/index';
import * as dotenv from 'dotenv';
import * as path from 'path';

dotenv.config({ path: './config/.env' });

async function liveUpdates() {
  const { bloomberg, excel } = createBloombergExcelConnector(
    {
      apiKey: process.env.BLOOMBERG_API_KEY!,
      apiSecret: process.env.BLOOMBERG_API_SECRET!,
      serverHost: process.env.BLOOMBERG_SERVER_HOST!,
      serverPort: parseInt(process.env.BLOOMBERG_SERVER_PORT!),
      enableWebSocket: true,
    },
    {
      autoRefresh: true,
      refreshInterval: 5000, // Update every 5 seconds
      enableFormatting: true,
    }
  );

  try {
    console.log('Setting up live Excel workbook...');

    // Define live data ranges
    const ranges = [
      {
        workbook: 'live-market-data',
        worksheet: 'Tech Stocks',
        range: 'A1:F10',
        securities: ['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'META', 'TSLA', 'NVDA', 'AMD', 'INTC'],
        fields: [
          BLOOMBERG_FIELDS.LAST_PRICE,
          BLOOMBERG_FIELDS.BID,
          BLOOMBERG_FIELDS.ASK,
          BLOOMBERG_FIELDS.VOLUME,
        ],
        updateFrequency: 5000,
      },
    ];

    const outputPath = path.join(
      process.env.EXCEL_OUTPUT_PATH || './output',
      'live-market-data.xlsx'
    );

    // Create live workbook
    await excel.createLiveWorkbook('live-market-data', ranges, outputPath);

    console.log(`✓ Live workbook created: ${outputPath}`);
    console.log('Press Ctrl+C to stop...');

    // Listen for updates
    excel.on('liveDataUpdate', (data) => {
      console.log(`Update received: ${data.data.security.ticker} - ${data.data.last}`);
    });

    excel.on('workbookRefreshed', (data) => {
      console.log(`Workbook refreshed: ${new Date().toLocaleTimeString()}`);
    });

    // Keep the process running
    process.on('SIGINT', () => {
      console.log('\nShutting down...');
      excel.closeLiveWorkbook('live-market-data');
      bloomberg.disconnect();
      excel.disconnect();
      process.exit(0);
    });
  } catch (error) {
    console.error('Live update failed:', error);
    bloomberg.disconnect();
    excel.disconnect();
  }
}

// Run example
liveUpdates();
