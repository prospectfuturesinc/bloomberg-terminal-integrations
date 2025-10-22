# Bloomberg Terminal Excel Connector

A powerful TypeScript/Node.js library for integrating Bloomberg Terminal market data with Microsoft Excel. Export real-time and historical financial data to Excel workbooks with advanced formatting, live updates, and automated data synchronization.

## Features

- **Real-Time Market Data**: Stream live prices, volumes, and market data directly to Excel
- **Historical Data Export**: Export historical time series data with flexible date ranges
- **Live Updating Workbooks**: Create Excel files that automatically refresh with real-time data
- **Advanced Formatting**: Apply custom styles, conditional formatting, and color scales
- **Multi-Sheet Exports**: Create complex workbooks with multiple worksheets and data sources
- **WebSocket Support**: Real-time data streaming via WebSocket connections
- **Type-Safe**: Full TypeScript support with comprehensive type definitions
- **Event-Driven**: React to data updates and connection changes via events

## Installation

```bash
npm install bloomberg-terminal-integrations
```

## Prerequisites

- Node.js 18.0 or higher
- Bloomberg Terminal API access (API key and secret)
- Bloomberg Terminal with API connectivity enabled

## Quick Start

### 1. Configure Environment

Create a `.env` file based on the example:

```bash
cp config/.env.example .env
```

Edit `.env` with your Bloomberg API credentials:

```env
BLOOMBERG_API_KEY=your_api_key_here
BLOOMBERG_API_SECRET=your_api_secret_here
BLOOMBERG_SERVER_HOST=api.bloomberg.com
BLOOMBERG_SERVER_PORT=8194
```

### 2. Basic Usage

```typescript
import { createBloombergExcelConnector, BLOOMBERG_FIELDS } from 'bloomberg-terminal-integrations';

// Create connector
const { bloomberg, excel } = createBloombergExcelConnector({
  apiKey: process.env.BLOOMBERG_API_KEY!,
  apiSecret: process.env.BLOOMBERG_API_SECRET!,
  serverHost: 'api.bloomberg.com',
  serverPort: 8194,
  enableWebSocket: true,
});

// Define securities
const securities = [
  { ticker: 'AAPL', exchange: 'NASDAQ', securityType: 'EQUITY' },
  { ticker: 'MSFT', exchange: 'NASDAQ', securityType: 'EQUITY' },
];

// Export to Excel
const filePath = await excel.exportMarketData(
  securities,
  [BLOOMBERG_FIELDS.LAST_PRICE, BLOOMBERG_FIELDS.VOLUME],
  'market-data.xlsx',
  './output'
);

console.log(`Exported to: ${filePath}`);
```

## API Reference

### BloombergClient

The main client for interacting with Bloomberg Terminal API.

#### Constructor

```typescript
const bloomberg = new BloombergClient({
  apiKey: string,
  apiSecret: string,
  serverHost: string,
  serverPort: number,
  timeout?: number,
  enableWebSocket?: boolean,
});
```

#### Methods

##### getRealTimeData()

Fetch real-time market data for specified securities.

```typescript
const response = await bloomberg.getRealTimeData(
  securities: SecurityIdentifier[],
  fields: string[]
);
```

##### getHistoricalData()

Fetch historical time series data.

```typescript
const response = await bloomberg.getHistoricalData({
  security: SecurityIdentifier,
  fields: string[],
  startDate: Date,
  endDate: Date,
  period: 'DAILY' | 'WEEKLY' | 'MONTHLY',
});
```

##### subscribe()

Subscribe to real-time market data updates via WebSocket.

```typescript
const subscriptionId = bloomberg.subscribe(
  securities: SecurityIdentifier[],
  fields: string[],
  callback: (data: MarketDataPoint) => void
);
```

##### searchSecurities()

Search for securities by ticker or name.

```typescript
const response = await bloomberg.searchSecurities(query: string, limit?: number);
```

### ExcelConnector

Handles Excel file creation and data export.

#### Constructor

```typescript
const excel = new ExcelConnector(bloombergClient: BloombergClient, config?: ExcelConfig);
```

#### Methods

##### exportMarketData()

Export real-time market data to Excel.

```typescript
const filePath = await excel.exportMarketData(
  securities: SecurityIdentifier[],
  fields: string[],
  fileName: string,
  filePath?: string
);
```

##### exportHistoricalData()

Export historical data to Excel.

```typescript
const filePath = await excel.exportHistoricalData(
  security: SecurityIdentifier,
  fields: string[],
  startDate: Date,
  endDate: Date,
  fileName: string,
  filePath?: string
);
```

##### exportToExcel()

Create custom Excel workbooks with multiple sheets and advanced formatting.

```typescript
const filePath = await excel.exportToExcel({
  fileName: string,
  filePath?: string,
  worksheets: ExcelWorksheetConfig[],
  includeCharts?: boolean,
});
```

##### createLiveWorkbook()

Create a live-updating Excel workbook.

```typescript
await excel.createLiveWorkbook(
  workbookId: string,
  ranges: LiveDataRange[],
  filePath: string
);
```

### Bloomberg Fields

Common Bloomberg field codes are available in the `BLOOMBERG_FIELDS` constant:

```typescript
import { BLOOMBERG_FIELDS } from 'bloomberg-terminal-integrations';

// Available fields:
BLOOMBERG_FIELDS.LAST_PRICE    // 'PX_LAST'
BLOOMBERG_FIELDS.BID           // 'PX_BID'
BLOOMBERG_FIELDS.ASK           // 'PX_ASK'
BLOOMBERG_FIELDS.VOLUME        // 'PX_VOLUME'
BLOOMBERG_FIELDS.OPEN          // 'PX_OPEN'
BLOOMBERG_FIELDS.HIGH          // 'PX_HIGH'
BLOOMBERG_FIELDS.LOW           // 'PX_LOW'
BLOOMBERG_FIELDS.CLOSE         // 'PX_CLOSE'
BLOOMBERG_FIELDS.MARKET_CAP    // 'CUR_MKT_CAP'
BLOOMBERG_FIELDS.PE_RATIO      // 'PE_RATIO'
BLOOMBERG_FIELDS.DIVIDEND_YIELD // 'DVD_YIELD'
```

## Examples

### Export Real-Time Market Data

```typescript
import { createBloombergExcelConnector, BLOOMBERG_FIELDS } from 'bloomberg-terminal-integrations';

const { bloomberg, excel } = createBloombergExcelConnector(config);

const securities = [
  { ticker: 'AAPL', exchange: 'NASDAQ', securityType: 'EQUITY' },
  { ticker: 'GOOGL', exchange: 'NASDAQ', securityType: 'EQUITY' },
];

const filePath = await excel.exportMarketData(
  securities,
  [BLOOMBERG_FIELDS.LAST_PRICE, BLOOMBERG_FIELDS.VOLUME, BLOOMBERG_FIELDS.MARKET_CAP],
  'tech-stocks.xlsx'
);
```

### Export Historical Data

```typescript
const security = { ticker: 'AAPL', exchange: 'NASDAQ', securityType: 'EQUITY' };

const endDate = new Date();
const startDate = new Date();
startDate.setDate(startDate.getDate() - 90); // Last 90 days

const filePath = await excel.exportHistoricalData(
  security,
  [BLOOMBERG_FIELDS.CLOSE, BLOOMBERG_FIELDS.VOLUME],
  startDate,
  endDate,
  'aapl-history.xlsx'
);
```

### Create Live Updating Workbook

```typescript
const ranges = [{
  workbook: 'live-data',
  worksheet: 'Tech Stocks',
  range: 'A1:F10',
  securities: ['AAPL', 'MSFT', 'GOOGL'],
  fields: [BLOOMBERG_FIELDS.LAST_PRICE, BLOOMBERG_FIELDS.VOLUME],
  updateFrequency: 5000, // Update every 5 seconds
}];

await excel.createLiveWorkbook('live-data', ranges, './output/live-data.xlsx');

// Listen for updates
excel.on('liveDataUpdate', (data) => {
  console.log(`Update: ${data.data.security.ticker} - $${data.data.last}`);
});
```

### Custom Multi-Sheet Export

```typescript
await excel.exportToExcel({
  fileName: 'portfolio-report.xlsx',
  filePath: './reports',
  worksheets: [
    {
      name: 'Equities',
      headers: ['Ticker', 'Price', 'Volume', 'Change %'],
      data: equitiesData,
      formatting: {
        headerStyle: {
          font: { bold: true, color: 'FFFFFF' },
          fill: { type: 'pattern', pattern: 'solid', fgColor: '0066CC' },
        },
        columnWidths: [15, 15, 20, 15],
      },
    },
    {
      name: 'Bonds',
      headers: ['ISIN', 'Price', 'Yield', 'Maturity'],
      data: bondsData,
      formatting: {
        headerStyle: {
          font: { bold: true, color: 'FFFFFF' },
          fill: { type: 'pattern', pattern: 'solid', fgColor: '006600' },
        },
      },
    },
  ],
});
```

## Running Examples

The repository includes several example scripts:

```bash
# Basic market data export
npm run dev examples/basic-export.ts

# Historical data export
npm run dev examples/historical-export.ts

# Live updating workbook
npm run dev examples/live-updates.ts

# Custom multi-sheet export
npm run dev examples/custom-export.ts
```

## Project Structure

```
bloomberg-terminal-integrations/
├── src/
│   ├── connectors/
│   │   ├── bloomberg-client.ts      # Bloomberg API client
│   │   └── excel/
│   │       └── excel-connector.ts   # Excel export functionality
│   ├── types/
│   │   ├── bloomberg.types.ts       # Bloomberg data types
│   │   └── excel.types.ts           # Excel configuration types
│   ├── utils/
│   │   ├── logger.ts                # Logging utility
│   │   └── data-formatter.ts        # Data formatting helpers
│   └── index.ts                     # Main entry point
├── examples/                         # Example scripts
├── config/                          # Configuration files
├── tests/                           # Test files
└── dist/                            # Compiled output
```

## Development

### Build

```bash
npm run build
```

### Run Tests

```bash
npm test
```

### Linting

```bash
npm run lint
```

## Configuration Options

### BloombergConfig

- `apiKey` (string): Bloomberg API key
- `apiSecret` (string): Bloomberg API secret
- `serverHost` (string): Bloomberg server hostname
- `serverPort` (number): Bloomberg server port
- `timeout` (number, optional): Request timeout in milliseconds (default: 30000)
- `enableWebSocket` (boolean, optional): Enable WebSocket connections (default: true)

### ExcelConfig

- `autoRefresh` (boolean, optional): Enable automatic refresh (default: false)
- `refreshInterval` (number, optional): Refresh interval in milliseconds (default: 5000)
- `enableFormatting` (boolean, optional): Enable cell formatting (default: true)
- `includeHeaders` (boolean, optional): Include column headers (default: true)
- `sheetName` (string, optional): Default sheet name (default: 'Market Data')

## Error Handling

The library uses a consistent error handling pattern:

```typescript
const response = await bloomberg.getRealTimeData(securities, fields);

if (!response.success) {
  console.error('Error:', response.error?.message);
  // Handle error
} else {
  // Process response.data
}
```

## Events

### BloombergClient Events

- `connected`: WebSocket connection established
- `disconnected`: WebSocket connection closed
- `error`: Connection or API error occurred
- `marketData`: Market data update received

### ExcelConnector Events

- `exported`: Excel file successfully exported
- `liveDataUpdate`: Live data range updated
- `workbookRefreshed`: Live workbook refreshed and saved

## Performance Considerations

- **WebSocket vs HTTP**: Use WebSocket connections for real-time data streaming
- **Batch Requests**: Combine multiple securities in single requests when possible
- **Live Updates**: Set appropriate refresh intervals (5-10 seconds recommended)
- **Data Volume**: Large historical datasets may take time to export
- **Memory Usage**: Monitor memory with live workbooks containing many securities

## Security

- Store API credentials in environment variables, never in code
- Use secure connections (HTTPS/WSS) for all Bloomberg API communications
- Implement proper access controls for generated Excel files
- Follow Bloomberg Terminal API usage policies and rate limits

## License

See [LICENSE.md](LICENSE.md) for details.

## Support

For issues and questions:
- GitHub Issues: [bloomberg-terminal-integrations/issues](https://github.com/prospectfuturesinc/bloomberg-terminal-integrations/issues)
- Email: support@prospectfutures.com

## Contributing

Contributions are welcome! Please follow these guidelines:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## Acknowledgments

- Built for [Prospect Futures, Inc.](https://prospectfutures.com)
- Bloomberg Terminal API integration
- ExcelJS library for Excel file generation

## Changelog

### Version 1.0.0 (2025-10-22)

- Initial release
- Bloomberg Terminal API client
- Excel connector with real-time and historical data export
- Live updating workbook support
- Advanced formatting and multi-sheet exports
- WebSocket support for real-time data streaming
- Comprehensive TypeScript type definitions
- Example scripts and documentation
