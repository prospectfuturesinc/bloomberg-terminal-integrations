/**
 * Excel Connector for Bloomberg Terminal
 * Exports Bloomberg market data to Excel with real-time updates
 */

import ExcelJS from 'exceljs';
import { EventEmitter } from 'events';
import { BloombergClient } from '../bloomberg-client';
import {
  ExcelConfig,
  ExcelWorksheetConfig,
  ExcelExportOptions,
  LiveDataRange,
  ExcelConnectionStatus,
  ExcelFormattingOptions,
} from '../../types/excel.types';
import { MarketDataPoint, SecurityIdentifier } from '../../types/bloomberg.types';
import { Logger } from '../../utils/logger';
import * as path from 'path';
import * as fs from 'fs';

export class ExcelConnector extends EventEmitter {
  private bloombergClient: BloombergClient;
  private config: ExcelConfig;
  private logger: Logger;
  private activeWorkbooks: Map<string, ExcelJS.Workbook>;
  private liveRanges: Map<string, LiveDataRange>;
  private updateIntervals: Map<string, NodeJS.Timeout>;

  constructor(bloombergClient: BloombergClient, config: ExcelConfig = {}) {
    super();
    this.bloombergClient = bloombergClient;
    this.config = {
      autoRefresh: config.autoRefresh ?? false,
      refreshInterval: config.refreshInterval ?? 5000,
      enableFormatting: config.enableFormatting ?? true,
      includeHeaders: config.includeHeaders ?? true,
      sheetName: config.sheetName ?? 'Market Data',
    };
    this.logger = new Logger('ExcelConnector');
    this.activeWorkbooks = new Map();
    this.liveRanges = new Map();
    this.updateIntervals = new Map();
  }

  /**
   * Export market data to Excel file
   */
  async exportToExcel(options: ExcelExportOptions): Promise<string> {
    try {
      const workbook = new ExcelJS.Workbook();

      // Set workbook properties
      workbook.creator = 'Bloomberg Terminal Integration';
      workbook.created = new Date();
      workbook.modified = new Date();

      // Add worksheets
      for (const worksheetConfig of options.worksheets) {
        await this.addWorksheet(workbook, worksheetConfig);
      }

      // Determine file path
      const filePath = options.filePath
        ? path.join(options.filePath, options.fileName)
        : options.fileName;

      // Ensure directory exists
      const directory = path.dirname(filePath);
      if (!fs.existsSync(directory)) {
        fs.mkdirSync(directory, { recursive: true });
      }

      // Save workbook
      await workbook.xlsx.writeFile(filePath);

      this.logger.info(`Excel file exported successfully: ${filePath}`);
      this.emit('exported', { filePath, worksheets: options.worksheets.length });

      return filePath;
    } catch (error) {
      this.logger.error('Failed to export Excel file', error);
      throw error;
    }
  }

  /**
   * Add worksheet with data to workbook
   */
  private async addWorksheet(
    workbook: ExcelJS.Workbook,
    config: ExcelWorksheetConfig
  ): Promise<void> {
    const worksheet = workbook.addWorksheet(config.name);

    // Add headers if configured
    if (config.headers && this.config.includeHeaders) {
      const headerRow = worksheet.addRow(config.headers);

      if (this.config.enableFormatting && config.formatting?.headerStyle) {
        headerRow.eachCell((cell) => {
          cell.font = config.formatting!.headerStyle!.font || { bold: true };
          cell.fill = config.formatting!.headerStyle!.fill || {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0066CC' },
          };
          cell.alignment = config.formatting!.headerStyle!.alignment || {
            horizontal: 'center',
            vertical: 'middle',
          };
        });
      }
    }

    // Add data rows
    config.data.forEach((row) => {
      const dataRow = Array.isArray(row) ? row : Object.values(row);
      worksheet.addRow(dataRow);
    });

    // Apply column widths
    if (config.formatting?.columnWidths) {
      config.formatting.columnWidths.forEach((width, index) => {
        const column = worksheet.getColumn(index + 1);
        column.width = width;
      });
    } else {
      // Auto-fit columns
      worksheet.columns.forEach((column) => {
        let maxLength = 0;
        column.eachCell?.({ includeEmpty: true }, (cell) => {
          const columnLength = cell.value ? cell.value.toString().length : 10;
          if (columnLength > maxLength) {
            maxLength = columnLength;
          }
        });
        column.width = Math.min(maxLength + 2, 50);
      });
    }

    // Apply conditional formatting
    if (config.formatting?.conditionalFormatting) {
      this.applyConditionalFormatting(worksheet, config.formatting.conditionalFormatting);
    }

    this.logger.info(`Added worksheet: ${config.name} with ${config.data.length} rows`);
  }

  /**
   * Apply conditional formatting to worksheet
   */
  private applyConditionalFormatting(
    worksheet: ExcelJS.Worksheet,
    formats: any[]
  ): void {
    formats.forEach((format) => {
      if (format.type === 'colorScale') {
        worksheet.addConditionalFormatting({
          ref: format.column,
          rules: [
            {
              type: 'colorScale',
              cfvo: [
                { type: 'min', value: undefined },
                { type: 'max', value: undefined },
              ],
              color: [{ argb: 'FFFF0000' }, { argb: 'FF00FF00' }],
            },
          ],
        });
      }
    });
  }

  /**
   * Export real-time market data to Excel
   */
  async exportMarketData(
    securities: SecurityIdentifier[],
    fields: string[],
    fileName: string,
    filePath?: string
  ): Promise<string> {
    try {
      // Fetch real-time data from Bloomberg
      const response = await this.bloombergClient.getRealTimeData(securities, fields);

      if (!response.success || !response.data) {
        throw new Error(`Failed to fetch market data: ${response.error?.message}`);
      }

      // Transform data for Excel
      const headers = ['Ticker', 'Exchange', 'Timestamp', ...fields];
      const data = response.data.map((point) => [
        point.security.ticker,
        point.security.exchange || '',
        point.timestamp.toISOString(),
        point.last,
        point.volume,
        point.bid,
        point.ask,
        point.open,
        point.high,
        point.low,
      ]);

      // Export to Excel
      return await this.exportToExcel({
        fileName,
        filePath,
        worksheets: [
          {
            name: 'Market Data',
            headers,
            data,
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
              dataStyle: {
                numberFormat: '#,##0.00',
                alignment: { horizontal: 'right' },
              },
              conditionalFormatting: [
                {
                  type: 'colorScale',
                  column: 'D:D',
                  rule: {},
                },
              ],
            },
          },
        ],
      });
    } catch (error) {
      this.logger.error('Failed to export market data', error);
      throw error;
    }
  }

  /**
   * Export historical data to Excel
   */
  async exportHistoricalData(
    security: SecurityIdentifier,
    fields: string[],
    startDate: Date,
    endDate: Date,
    fileName: string,
    filePath?: string
  ): Promise<string> {
    try {
      // Fetch historical data from Bloomberg
      const response = await this.bloombergClient.getHistoricalData({
        security,
        fields,
        startDate,
        endDate,
        period: 'DAILY',
      });

      if (!response.success || !response.data) {
        throw new Error(`Failed to fetch historical data: ${response.error?.message}`);
      }

      // Transform data for Excel
      const headers = ['Date', ...fields];
      const data = response.data.map((point) => {
        const row: any[] = [point.date];
        fields.forEach((field) => {
          row.push(point[field] || '');
        });
        return row;
      });

      // Export to Excel
      return await this.exportToExcel({
        fileName,
        filePath,
        worksheets: [
          {
            name: `${security.ticker} Historical`,
            headers,
            data,
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
            },
          },
        ],
      });
    } catch (error) {
      this.logger.error('Failed to export historical data', error);
      throw error;
    }
  }

  /**
   * Create live updating Excel workbook
   */
  async createLiveWorkbook(
    workbookId: string,
    ranges: LiveDataRange[],
    filePath: string
  ): Promise<void> {
    try {
      // Create initial workbook
      const workbook = new ExcelJS.Workbook();
      this.activeWorkbooks.set(workbookId, workbook);

      // Set up live ranges
      ranges.forEach((range) => {
        const rangeId = `${workbookId}_${range.worksheet}`;
        this.liveRanges.set(rangeId, range);

        // Subscribe to Bloomberg updates
        const subscription = this.bloombergClient.subscribe(
          range.securities.map((ticker) => ({
            ticker,
            securityType: 'EQUITY',
          })),
          range.fields,
          (data: MarketDataPoint) => {
            this.updateLiveRange(rangeId, data);
          }
        );

        // Set up periodic updates
        if (this.config.autoRefresh) {
          const interval = setInterval(async () => {
            await this.refreshLiveWorkbook(workbookId, filePath);
          }, range.updateFrequency);

          this.updateIntervals.set(rangeId, interval);
        }
      });

      this.logger.info(`Created live workbook: ${workbookId} with ${ranges.length} ranges`);
    } catch (error) {
      this.logger.error('Failed to create live workbook', error);
      throw error;
    }
  }

  /**
   * Update live data range
   */
  private updateLiveRange(rangeId: string, data: MarketDataPoint): void {
    this.emit('liveDataUpdate', { rangeId, data });
  }

  /**
   * Refresh live workbook and save to file
   */
  private async refreshLiveWorkbook(workbookId: string, filePath: string): Promise<void> {
    try {
      const workbook = this.activeWorkbooks.get(workbookId);
      if (!workbook) {
        throw new Error(`Workbook not found: ${workbookId}`);
      }

      await workbook.xlsx.writeFile(filePath);
      this.logger.debug(`Refreshed live workbook: ${workbookId}`);
      this.emit('workbookRefreshed', { workbookId, filePath });
    } catch (error) {
      this.logger.error('Failed to refresh live workbook', error);
    }
  }

  /**
   * Stop live workbook updates
   */
  closeLiveWorkbook(workbookId: string): void {
    // Clear all intervals for this workbook
    this.updateIntervals.forEach((interval, rangeId) => {
      if (rangeId.startsWith(workbookId)) {
        clearInterval(interval);
        this.updateIntervals.delete(rangeId);
      }
    });

    // Remove workbook
    this.activeWorkbooks.delete(workbookId);

    this.logger.info(`Closed live workbook: ${workbookId}`);
  }

  /**
   * Get connection status
   */
  getStatus(): ExcelConnectionStatus {
    return {
      connected: this.bloombergClient.isConnected(),
      lastUpdate: new Date(),
      activeRanges: this.liveRanges.size,
    };
  }

  /**
   * Close all connections
   */
  disconnect(): void {
    // Clear all intervals
    this.updateIntervals.forEach((interval) => clearInterval(interval));
    this.updateIntervals.clear();

    // Clear all workbooks
    this.activeWorkbooks.clear();
    this.liveRanges.clear();

    this.logger.info('Excel connector disconnected');
  }
}
