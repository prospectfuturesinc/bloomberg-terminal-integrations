/**
 * Advanced Example: Custom multi-sheet Excel export with formatting
 */

import { BloombergClient, ExcelConnector, BLOOMBERG_FIELDS } from '../src/index';
import * as dotenv from 'dotenv';

dotenv.config({ path: './config/.env' });

async function customExport() {
  // Initialize clients
  const bloomberg = new BloombergClient({
    apiKey: process.env.BLOOMBERG_API_KEY!,
    apiSecret: process.env.BLOOMBERG_API_SECRET!,
    serverHost: process.env.BLOOMBERG_SERVER_HOST!,
    serverPort: parseInt(process.env.BLOOMBERG_SERVER_PORT!),
  });

  const excel = new ExcelConnector(bloomberg, {
    enableFormatting: true,
    includeHeaders: true,
  });

  try {
    console.log('Creating custom Excel report...');

    // Fetch data for different sectors
    const techStocks = [
      { ticker: 'AAPL', exchange: 'NASDAQ', securityType: 'EQUITY' as const },
      { ticker: 'MSFT', exchange: 'NASDAQ', securityType: 'EQUITY' as const },
      { ticker: 'GOOGL', exchange: 'NASDAQ', securityType: 'EQUITY' as const },
    ];

    const energyStocks = [
      { ticker: 'XOM', exchange: 'NYSE', securityType: 'EQUITY' as const },
      { ticker: 'CVX', exchange: 'NYSE', securityType: 'EQUITY' as const },
      { ticker: 'COP', exchange: 'NYSE', securityType: 'EQUITY' as const },
    ];

    const fields = [
      BLOOMBERG_FIELDS.LAST_PRICE,
      BLOOMBERG_FIELDS.MARKET_CAP,
      BLOOMBERG_FIELDS.PE_RATIO,
      BLOOMBERG_FIELDS.DIVIDEND_YIELD,
    ];

    // Fetch market data
    const techResponse = await bloomberg.getRealTimeData(techStocks, fields);
    const energyResponse = await bloomberg.getRealTimeData(energyStocks, fields);

    if (!techResponse.success || !energyResponse.success) {
      throw new Error('Failed to fetch market data');
    }

    // Prepare data for export
    const techData = techResponse.data!.map((point) => [
      point.security.ticker,
      point.last,
      point.volume,
      point.changePercent,
    ]);

    const energyData = energyResponse.data!.map((point) => [
      point.security.ticker,
      point.last,
      point.volume,
      point.changePercent,
    ]);

    // Create custom formatted Excel file
    const filePath = await excel.exportToExcel({
      fileName: 'sector-analysis.xlsx',
      filePath: process.env.EXCEL_OUTPUT_PATH,
      worksheets: [
        {
          name: 'Technology',
          headers: ['Ticker', 'Last Price', 'Volume', 'Change %'],
          data: techData,
          formatting: {
            headerStyle: {
              font: { bold: true, color: 'FFFFFF', size: 12 },
              fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: '0066CC',
              },
              alignment: { horizontal: 'center', vertical: 'middle' },
            },
            columnWidths: [15, 15, 20, 15],
            conditionalFormatting: [
              {
                type: 'colorScale',
                column: 'D:D',
                rule: {},
              },
            ],
          },
        },
        {
          name: 'Energy',
          headers: ['Ticker', 'Last Price', 'Volume', 'Change %'],
          data: energyData,
          formatting: {
            headerStyle: {
              font: { bold: true, color: 'FFFFFF', size: 12 },
              fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: '006600',
              },
              alignment: { horizontal: 'center', vertical: 'middle' },
            },
            columnWidths: [15, 15, 20, 15],
          },
        },
        {
          name: 'Summary',
          headers: ['Sector', 'Avg Price', 'Total Volume', 'Avg Change %'],
          data: [
            [
              'Technology',
              (techData.reduce((sum, row) => sum + row[1], 0) / techData.length).toFixed(2),
              techData.reduce((sum, row) => sum + row[2], 0).toLocaleString(),
              (techData.reduce((sum, row) => sum + row[3], 0) / techData.length).toFixed(2),
            ],
            [
              'Energy',
              (energyData.reduce((sum, row) => sum + row[1], 0) / energyData.length).toFixed(2),
              energyData.reduce((sum, row) => sum + row[2], 0).toLocaleString(),
              (energyData.reduce((sum, row) => sum + row[3], 0) / energyData.length).toFixed(2),
            ],
          ],
          formatting: {
            headerStyle: {
              font: { bold: true, color: 'FFFFFF', size: 14 },
              fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: '333333',
              },
              alignment: { horizontal: 'center', vertical: 'middle' },
            },
            columnWidths: [20, 20, 20, 20],
          },
        },
      ],
    });

    console.log(`✓ Custom export complete: ${filePath}`);
  } catch (error) {
    console.error('Export failed:', error);
  } finally {
    bloomberg.disconnect();
    excel.disconnect();
  }
}

// Run example
customExport();
