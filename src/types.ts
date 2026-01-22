import { OSFDocument } from 'omniscript-parser';

export interface ConverterOptions {
  theme?: string;
  pageSize?: 'A4' | 'letter' | 'legal';
  orientation?: 'portrait' | 'landscape';
  margins?: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
  includeMetadata?: boolean;
  customStyles?: Record<string, string | number | boolean>;
  timeoutMs?: number;
  puppeteerArgs?: string[];
}

export interface ConversionResult {
  buffer: Buffer;
  mimeType: string;
  extension: string;
}

export interface Converter {
  // eslint-disable-next-line no-unused-vars -- type signature params are for clarity
  convert(document: OSFDocument, options?: ConverterOptions): Promise<ConversionResult>;
  getSupportedFormats(): string[];
}
