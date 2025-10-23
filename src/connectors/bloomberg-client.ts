/**
 * Bloomberg Terminal API Client
 * Handles connections to Bloomberg Terminal and data retrieval
 */

import axios, { AxiosInstance } from 'axios';
import WebSocket from 'ws';
import { EventEmitter } from 'events';
import {
  BloombergConfig,
  SecurityIdentifier,
  MarketDataPoint,
  HistoricalDataRequest,
  HistoricalDataPoint,
  ReferenceDataRequest,
  ReferenceDataResponse,
  BloombergSubscription,
  BloombergResponse,
  BloombergError,
} from '../types/bloomberg.types';
import { Logger } from '../utils/logger';

export class BloombergClient extends EventEmitter {
  private config: BloombergConfig;
  private httpClient: AxiosInstance;
  private wsClient?: WebSocket;
  private subscriptions: Map<string, BloombergSubscription>;
  private logger: Logger;
  private connected: boolean = false;

  constructor(config: BloombergConfig) {
    super();
    this.config = config;
    this.subscriptions = new Map();
    this.logger = new Logger('BloombergClient');

    // Initialize HTTP client
    this.httpClient = axios.create({
      baseURL: `https://${config.serverHost}:${config.serverPort}`,
      timeout: config.timeout || 30000,
      headers: {
        'X-API-Key': config.apiKey,
        'X-API-Secret': config.apiSecret,
        'Content-Type': 'application/json',
      },
    });

    // Initialize WebSocket for real-time data if enabled
    if (config.enableWebSocket) {
      this.initializeWebSocket();
    }
  }

  /**
   * Initialize WebSocket connection for real-time market data
   */
  private initializeWebSocket(): void {
    const wsUrl = `wss://${this.config.serverHost}:${this.config.serverPort}/stream`;

    this.wsClient = new WebSocket(wsUrl, {
      headers: {
        'X-API-Key': this.config.apiKey,
        'X-API-Secret': this.config.apiSecret,
      },
    });

    this.wsClient.on('open', () => {
      this.connected = true;
      this.logger.info('WebSocket connection established');
      this.emit('connected');
    });

    this.wsClient.on('message', (data: WebSocket.Data) => {
      try {
        const message = JSON.parse(data.toString());
        this.handleWebSocketMessage(message);
      } catch (error) {
        this.logger.error('Failed to parse WebSocket message', error);
      }
    });

    this.wsClient.on('error', (error) => {
      this.logger.error('WebSocket error', error);
      this.emit('error', error);
    });

    this.wsClient.on('close', () => {
      this.connected = false;
      this.logger.warn('WebSocket connection closed');
      this.emit('disconnected');

      // Attempt reconnection after 5 seconds
      setTimeout(() => this.initializeWebSocket(), 5000);
    });
  }

  /**
   * Handle incoming WebSocket messages
   */
  private handleWebSocketMessage(message: any): void {
    if (message.type === 'marketData') {
      const dataPoint: MarketDataPoint = {
        security: message.security,
        timestamp: new Date(message.timestamp),
        last: message.last,
        volume: message.volume,
        bid: message.bid,
        ask: message.ask,
        open: message.open,
        high: message.high,
        low: message.low,
        close: message.close,
        change: message.change,
        changePercent: message.changePercent,
      };

      // Trigger callbacks for subscribed securities
      this.subscriptions.forEach((subscription) => {
        if (this.matchesSecurity(dataPoint.security, subscription.securities)) {
          subscription.callback(dataPoint);
        }
      });

      this.emit('marketData', dataPoint);
    }
  }

  /**
   * Check if security matches subscription
   */
  private matchesSecurity(
    security: SecurityIdentifier,
    subscriptionSecurities: SecurityIdentifier[]
  ): boolean {
    return subscriptionSecurities.some(
      (sub) => sub.ticker === security.ticker && sub.exchange === security.exchange
    );
  }

  /**
   * Get real-time market data for securities
   */
  async getRealTimeData(
    securities: SecurityIdentifier[],
    fields: string[]
  ): Promise<BloombergResponse<MarketDataPoint[]>> {
    try {
      const response = await this.httpClient.post('/api/v1/market-data/realtime', {
        securities,
        fields,
      });

      return {
        success: true,
        data: response.data.data.map((point: any) => ({
          ...point,
          timestamp: new Date(point.timestamp),
        })),
      };
    } catch (error: any) {
      return this.handleError(error);
    }
  }

  /**
   * Get historical market data
   */
  async getHistoricalData(
    request: HistoricalDataRequest
  ): Promise<BloombergResponse<HistoricalDataPoint[]>> {
    try {
      const response = await this.httpClient.post('/api/v1/market-data/historical', {
        security: request.security,
        fields: request.fields,
        startDate: request.startDate.toISOString(),
        endDate: request.endDate.toISOString(),
        period: request.period || 'DAILY',
      });

      return {
        success: true,
        data: response.data.data.map((point: any) => ({
          ...point,
          date: new Date(point.date),
        })),
      };
    } catch (error: any) {
      return this.handleError(error);
    }
  }

  /**
   * Get reference data for securities
   */
  async getReferenceData(
    request: ReferenceDataRequest
  ): Promise<BloombergResponse<ReferenceDataResponse[]>> {
    try {
      const response = await this.httpClient.post('/api/v1/reference-data', {
        securities: request.securities,
        fields: request.fields,
      });

      return {
        success: true,
        data: response.data.data,
      };
    } catch (error: any) {
      return this.handleError(error);
    }
  }

  /**
   * Subscribe to real-time market data updates
   */
  subscribe(
    securities: SecurityIdentifier[],
    fields: string[],
    callback: (data: MarketDataPoint) => void
  ): string {
    const subscriptionId = `sub_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

    const subscription: BloombergSubscription = {
      id: subscriptionId,
      securities,
      fields,
      callback,
    };

    this.subscriptions.set(subscriptionId, subscription);

    // Send subscription request via WebSocket
    if (this.wsClient && this.connected) {
      this.wsClient.send(
        JSON.stringify({
          type: 'subscribe',
          id: subscriptionId,
          securities,
          fields,
        })
      );
    }

    this.logger.info(`Created subscription ${subscriptionId} for ${securities.length} securities`);
    return subscriptionId;
  }

  /**
   * Unsubscribe from market data updates
   */
  unsubscribe(subscriptionId: string): void {
    if (this.subscriptions.has(subscriptionId)) {
      // Send unsubscribe request via WebSocket
      if (this.wsClient && this.connected) {
        this.wsClient.send(
          JSON.stringify({
            type: 'unsubscribe',
            id: subscriptionId,
          })
        );
      }

      this.subscriptions.delete(subscriptionId);
      this.logger.info(`Removed subscription ${subscriptionId}`);
    }
  }

  /**
   * Get bulk security lookup
   */
  async searchSecurities(query: string, limit: number = 10): Promise<BloombergResponse<SecurityIdentifier[]>> {
    try {
      const response = await this.httpClient.get('/api/v1/securities/search', {
        params: { query, limit },
      });

      return {
        success: true,
        data: response.data.data,
      };
    } catch (error: any) {
      return this.handleError(error);
    }
  }

  /**
   * Handle API errors
   */
  private handleError(error: any): BloombergResponse<never> {
    const bloombergError: BloombergError = {
      code: error.response?.data?.code || 'UNKNOWN_ERROR',
      message: error.response?.data?.message || error.message,
      timestamp: new Date(),
    };

    this.logger.error('Bloomberg API error', bloombergError);

    return {
      success: false,
      error: bloombergError,
    };
  }

  /**
   * Close all connections
   */
  disconnect(): void {
    if (this.wsClient) {
      this.subscriptions.forEach((_, id) => this.unsubscribe(id));
      this.wsClient.close();
      this.wsClient = undefined;
    }
    this.connected = false;
    this.logger.info('Disconnected from Bloomberg Terminal');
  }

  /**
   * Check connection status
   */
  isConnected(): boolean {
    return this.connected;
  }
}
