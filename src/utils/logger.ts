/**
 * Logger utility for Bloomberg Terminal Integration
 */

import winston from 'winston';

export class Logger {
  private logger: winston.Logger;
  private context: string;

  constructor(context: string) {
    this.context = context;

    this.logger = winston.createLogger({
      level: process.env.LOG_LEVEL || 'info',
      format: winston.format.combine(
        winston.format.timestamp({
          format: 'YYYY-MM-DD HH:mm:ss',
        }),
        winston.format.errors({ stack: true }),
        winston.format.splat(),
        winston.format.json()
      ),
      defaultMeta: { service: 'bloomberg-excel-connector' },
      transports: [
        new winston.transports.File({ filename: 'logs/error.log', level: 'error' }),
        new winston.transports.File({ filename: 'logs/combined.log' }),
      ],
    });

    // Console logging in development
    if (process.env.NODE_ENV !== 'production') {
      this.logger.add(
        new winston.transports.Console({
          format: winston.format.combine(
            winston.format.colorize(),
            winston.format.simple()
          ),
        })
      );
    }
  }

  info(message: string, ...meta: any[]): void {
    this.logger.info(`[${this.context}] ${message}`, ...meta);
  }

  warn(message: string, ...meta: any[]): void {
    this.logger.warn(`[${this.context}] ${message}`, ...meta);
  }

  error(message: string, error?: any): void {
    this.logger.error(`[${this.context}] ${message}`, error);
  }

  debug(message: string, ...meta: any[]): void {
    this.logger.debug(`[${this.context}] ${message}`, ...meta);
  }
}
