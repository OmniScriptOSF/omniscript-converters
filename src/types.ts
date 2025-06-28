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
  customStyles?: Record<string, any>;
}

export interface ConversionResult {
  buffer: Buffer;
  mimeType: string;
  extension: string;
}

export interface Converter {
  convert(document: OSFDocument, options?: ConverterOptions): Promise<ConversionResult>;
  getSupportedFormats(): string[];
}