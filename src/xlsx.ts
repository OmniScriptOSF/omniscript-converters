// File: src/xlsx.ts
// What: XLSX converter implementation using xlsx-populate.
// Why: Generate styled XLSX output from OSF documents without ExcelJS.
// Related: src/types/xlsx-populate.d.ts, tests/xlsx.test.ts

import XlsxPopulate from 'xlsx-populate';
import type { Workbook, Sheet, Borders } from 'xlsx-populate';
import {
  OSFDocument,
  MetaBlock,
  DocBlock,
  SlideBlock,
  SheetBlock,
  type TextRun as OSFTextRun,
} from 'omniscript-parser';
import { Converter, ConverterOptions, ConversionResult } from './types';

type ThemeColors = {
  primary: string;
  accent: string;
  background: string;
};

type SheetState = {
  usedDefault: boolean;
};

export class XLSXConverter implements Converter {
  getSupportedFormats(): string[] {
    return ['xlsx'];
  }

  async convert(document: OSFDocument, options: ConverterOptions = {}): Promise<ConversionResult> {
    const workbook = await XlsxPopulate.fromBlankAsync();

    // Set workbook metadata
    this.setWorkbookMetadata(workbook, document);

    // Process document blocks
    this.generateWorksheets(workbook, document, options);

    const output = await workbook.outputAsync({ type: 'nodebuffer' });
    const buffer = Buffer.isBuffer(output) ? output : Buffer.from(output as ArrayBuffer);

    return {
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      extension: 'xlsx',
    };
  }

  private setWorkbookMetadata(workbook: Workbook, document: OSFDocument): void {
    const meta = this.getMetadata(document);

    const properties: Record<string, string> = {
      title: meta.title || 'OSF Workbook',
      creator: meta.author || 'OmniScript OSF',
      description: 'Generated from OSF document',
    };

    workbook.property(properties);
  }

  private generateWorksheets(
    workbook: Workbook,
    document: OSFDocument,
    options: ConverterOptions
  ): void {
    let worksheetIndex = 1;
    let hasSheets = false;
    const sheetState: SheetState = { usedDefault: false };

    for (const block of document.blocks) {
      switch (block.type) {
        case 'sheet':
          this.createSheetWorksheet(workbook, block as SheetBlock, options, sheetState);
          hasSheets = true;
          break;
        case 'doc':
        case 'slide':
          this.createContentWorksheet(
            workbook,
            block,
            options,
            `Content_${worksheetIndex}`,
            sheetState
          );
          worksheetIndex++;
          break;
        case 'meta':
          // Metadata will be included in summary if no sheets exist
          break;
      }
    }

    // If no dedicated sheets exist, create a summary worksheet
    if (!hasSheets) {
      this.createSummaryWorksheet(workbook, document, options, sheetState);
    }
  }

  private getOrCreateSheet(workbook: Workbook, name: string, sheetState: SheetState): Sheet {
    if (!sheetState.usedDefault) {
      const defaultSheet = workbook.sheet(0);
      defaultSheet.name(name);
      sheetState.usedDefault = true;
      return defaultSheet;
    }

    return workbook.addSheet(name);
  }

  private createSheetWorksheet(
    workbook: Workbook,
    sheet: SheetBlock,
    options: ConverterOptions,
    sheetState: SheetState
  ): void {
    const worksheetName = this.sanitizeWorksheetName(sheet.name || 'Sheet');
    const worksheet = this.getOrCreateSheet(workbook, worksheetName, sheetState);

    // Configure worksheet styling
    this.applyWorksheetStyling(worksheet, options);

    let currentRow = 1;

    // Add sheet title
    if (sheet.name) {
      const titleCell = worksheet.cell(currentRow, 1);
      titleCell.value(sheet.name);
      titleCell.style({
        bold: true,
        fontSize: 16,
        fontColor: this.toRgb('FF2C3E50'),
        fill: this.toRgb('FFF8F9FA'),
      });

      // Merge cells for title if we have multiple columns
      const colCount = this.getColumnCount(sheet);
      if (colCount > 1) {
        worksheet.range(currentRow, 1, currentRow, colCount).merged(true);
      }

      currentRow += 2;
    }

    // Add column headers
    if (sheet.cols) {
      const cols = Array.isArray(sheet.cols)
        ? sheet.cols
        : String(sheet.cols)
            .replace(/[[\]]/g, '')
            .split(',')
            .map(s => s.trim());

      cols.forEach((col, index) => {
        const cell = worksheet.cell(currentRow, index + 1);
        cell.value(col);
        cell.style({
          bold: true,
          fontColor: this.toRgb('FFFFFFFF'),
          fill: this.toRgb('FF3498DB'),
          border: this.getThinBorder(),
        });
      });

      currentRow++;
    }

    // Add data rows
    if (sheet.data) {
      this.populateSheetData(worksheet, sheet, currentRow);
    }

    // Apply formulas if any
    if (sheet.formulas && sheet.formulas.length > 0) {
      this.applyFormulas(worksheet, sheet.formulas);
    }

    // Auto-size columns
    this.autoSizeColumns(worksheet);
  }

  private createContentWorksheet(
    workbook: Workbook,
    block: DocBlock | SlideBlock,
    options: ConverterOptions,
    name: string,
    sheetState: SheetState
  ): void {
    const worksheetName = this.sanitizeWorksheetName(name);
    const worksheet = this.getOrCreateSheet(workbook, worksheetName, sheetState);
    this.applyWorksheetStyling(worksheet, options);

    if (block.type === 'doc') {
      this.addDocContentToWorksheet(worksheet, block, 1);
    } else if (block.type === 'slide') {
      this.addSlideContentToWorksheet(worksheet, block, 1);
    }

    this.autoSizeColumns(worksheet);
  }

  private createSummaryWorksheet(
    workbook: Workbook,
    document: OSFDocument,
    options: ConverterOptions,
    sheetState: SheetState
  ): void {
    const worksheet = this.getOrCreateSheet(workbook, 'Summary', sheetState);
    this.applyWorksheetStyling(worksheet, options);

    let currentRow = 1;

    // Add document metadata
    const meta = this.getMetadata(document);
    if (meta.title || meta.author || meta.date) {
      const titleCell = worksheet.cell(currentRow, 1);
      titleCell.value('Document Information');
      titleCell.style({
        bold: true,
        fontSize: 16,
        fontColor: this.toRgb('FF2C3E50'),
      });
      currentRow += 2;

      if (meta.title) {
        const labelCell = worksheet.cell(currentRow, 1);
        labelCell.value('Title:');
        labelCell.style({ bold: true });
        worksheet.cell(currentRow, 2).value(meta.title);
        currentRow++;
      }

      if (meta.author) {
        const labelCell = worksheet.cell(currentRow, 1);
        labelCell.value('Author:');
        labelCell.style({ bold: true });
        worksheet.cell(currentRow, 2).value(meta.author);
        currentRow++;
      }

      if (meta.date) {
        const labelCell = worksheet.cell(currentRow, 1);
        labelCell.value('Date:');
        labelCell.style({ bold: true });
        worksheet.cell(currentRow, 2).value(meta.date);
        currentRow++;
      }

      currentRow += 2;
    }

    // Add content summary
    const contentBlocks = document.blocks.filter(b => b.type === 'doc' || b.type === 'slide');
    if (contentBlocks.length > 0) {
      const summaryTitleCell = worksheet.cell(currentRow, 1);
      summaryTitleCell.value('Content Summary');
      summaryTitleCell.style({
        bold: true,
        fontSize: 14,
        fontColor: this.toRgb('FF2C3E50'),
      });
      currentRow += 2;

      // Add headers
      const typeHeader = worksheet.cell(currentRow, 1);
      typeHeader.value('Type');
      typeHeader.style({ bold: true });

      const previewHeader = worksheet.cell(currentRow, 2);
      previewHeader.value('Title/Content Preview');
      previewHeader.style({ bold: true });
      currentRow++;

      for (const block of contentBlocks) {
        worksheet.cell(currentRow, 1).value(block.type.toUpperCase());

        if (block.type === 'slide') {
          const slide = block as SlideBlock;
          worksheet.cell(currentRow, 2).value(slide.title || 'Untitled Slide');
        } else if (block.type === 'doc') {
          const doc = block as DocBlock;
          const preview = this.getContentPreview(doc.content || '');
          worksheet.cell(currentRow, 2).value(preview);
        }

        currentRow++;
      }
    }

    this.autoSizeColumns(worksheet);
  }

  private populateSheetData(worksheet: Sheet, sheet: SheetBlock, startRow: number): void {
    if (!sheet.data) return;

    const coords = Object.keys(sheet.data).map(k => k.split(',').map(Number));
    const maxRow = Math.max(...coords.map(c => c[0]));
    const maxCol = Math.max(...coords.map(c => c[1]));

    // Populate data cells
    for (let r = 1; r <= maxRow; r++) {
      for (let c = 1; c <= maxCol; c++) {
        const key = `${r},${c}`;
        const value = sheet.data[key];

        if (value !== undefined) {
          const cell = worksheet.cell(startRow + r - 1, c);

          // Handle different data types
          if (typeof value === 'number') {
            cell.value(value);
            cell.style({ numberFormat: '#,##0.00' });
          } else if (typeof value === 'boolean') {
            cell.value(value);
          } else {
            cell.value(String(value));
          }

          // Apply basic styling
          cell.style({ border: this.getThinBorder() });
        }
      }
    }
  }

  private applyFormulas(
    worksheet: Sheet,
    formulas: Array<{ cell: [number, number]; expr: string }>
  ): void {
    for (const formula of formulas) {
      const [row, col] = formula.cell;
      const cell = worksheet.cell(row, col);

      // Convert OSF formula to Excel formula
      let excelFormula = formula.expr;
      if (!excelFormula.startsWith('=')) {
        excelFormula = '=' + excelFormula;
      }

      // Convert cell references if needed (basic conversion)
      excelFormula = this.convertToExcelFormula(excelFormula);

      cell.formula(excelFormula);

      // Style formula cells
      cell.style({
        fill: this.toRgb('FFF0F8FF'),
        italic: true,
      });
    }
  }

  private addDocContentToWorksheet(worksheet: Sheet, doc: DocBlock, startRow: number): number {
    let currentRow = startRow;
    const content = doc.content || '';

    // Add title
    const titleCell = worksheet.cell(currentRow, 1);
    titleCell.value('Document Content');
    titleCell.style({
      bold: true,
      fontSize: 14,
      fontColor: this.toRgb('FF2C3E50'),
    });
    currentRow += 2;

    // Split content into paragraphs and add to cells
    const paragraphs = content.split('\n\n');
    for (const paragraph of paragraphs) {
      if (paragraph.trim()) {
        const cell = worksheet.cell(currentRow, 1);
        cell.value(paragraph.trim());
        cell.style({ wrapText: true, verticalAlignment: 'top' });
        const height = Math.max(15, Math.min(100, paragraph.length / 10));
        worksheet.row(currentRow).height(height);
        currentRow++;
      }
    }

    return currentRow + 1;
  }

  private addSlideContentToWorksheet(worksheet: Sheet, slide: SlideBlock, startRow: number): number {
    let currentRow = startRow;

    // Add slide title
    if (slide.title) {
      const titleCell = worksheet.cell(currentRow, 1);
      titleCell.value(slide.title);
      titleCell.style({
        bold: true,
        fontSize: 14,
        fontColor: this.toRgb('FF2C3E50'),
      });
      currentRow += 2;
    }

    // Add slide content
    if (slide.content) {
      for (const contentBlock of slide.content) {
        if (contentBlock.type === 'unordered_list') {
          for (const item of contentBlock.items) {
            const itemText = item.content.map(this.extractText).join('');
            worksheet.cell(currentRow, 1).value(`• ${itemText}`);
            currentRow++;
          }
        } else if (contentBlock.type === 'paragraph') {
          const paragraphText = contentBlock.content.map(this.extractText).join('');
          const cell = worksheet.cell(currentRow, 1);
          cell.value(paragraphText);
          cell.style({ wrapText: true });
          currentRow++;
        }
      }
    }

    return currentRow + 1;
  }

  private applyWorksheetStyling(worksheet: Sheet, options: ConverterOptions): void {
    // Set default column width
    worksheet.column(1).width(20);
    worksheet.column(2).width(30);
    worksheet.column(3).width(15);
    worksheet.column(4).width(15);

    // Apply theme-based styling if specified
    const theme = options.theme || 'default';
    const themeColors = this.getThemeColors(theme);

    // Set worksheet tab color based on theme
    worksheet.tabColor(this.toRgb(themeColors.accent));
  }

  private autoSizeColumns(worksheet: Sheet): void {
    const usedRange = worksheet.usedRange();
    if (!usedRange) return;

    const values = usedRange.value();
    if (!values.length) return;

    const columnCount = Math.max(...values.map(row => row.length));
    if (!columnCount) return;

    const columnWidths = new Array<number>(columnCount).fill(10);

    values.forEach(row => {
      row.forEach((value, index) => {
        if (value === null || value === undefined) return;
        const length = String(value).length;
        const width = Math.min(50, length + 2);
        columnWidths[index] = Math.max(columnWidths[index], width);
      });
    });

    columnWidths.forEach((width, index) => {
      worksheet.column(index + 1).width(width);
    });
  }

  private getColumnCount(sheet: SheetBlock): number {
    if (sheet.cols) {
      const cols = Array.isArray(sheet.cols)
        ? sheet.cols
        : String(sheet.cols)
            .replace(/[[\]]/g, '')
            .split(',')
            .map(s => s.trim());
      return cols.length;
    }

    if (sheet.data) {
      const coords = Object.keys(sheet.data).map(k => k.split(',').map(Number));
      return Math.max(...coords.map(c => c[1]));
    }

    return 1;
  }

  private convertToExcelFormula(formula: string): string {
    // Basic conversion from OSF formula format to Excel format
    // This is a simplified conversion - you might want to enhance this
    return formula.replace(/\(\s*(\d+),\s*(\d+)\s*\)/g, (match, row, col) => {
      // Convert (row,col) to Excel cell reference
      const colLetter = this.numberToColumnLetter(parseInt(col));
      return `${colLetter}${row}`;
    });
  }

  private numberToColumnLetter(num: number): string {
    let result = '';
    while (num > 0) {
      num--;
      result = String.fromCharCode(65 + (num % 26)) + result;
      num = Math.floor(num / 26);
    }
    return result;
  }

  private sanitizeWorksheetName(name: string): string {
    // Excel worksheet names have restrictions
    return name.replace(/[\\/*?:\\[\\]]/g, '_').substring(0, 31); // Max 31 characters
  }

  private getContentPreview(content: string): string {
    const preview = content.replace(/[#*`]/g, '').replace(/\n+/g, ' ').trim();
    return preview.length > 50 ? preview.substring(0, 47) + '...' : preview;
  }

  private extractText(run: OSFTextRun): string {
    if (typeof run === 'string') return run;
    if (typeof run === 'object' && run !== null && 'type' in run) {
      if (run.type === 'link') return run.text;
      if (run.type === 'image') return run.alt || '';
    }
    if ('text' in run && typeof run.text === 'string') return run.text;
    return '';
  }

  private getMetadata(document: OSFDocument): { title?: string; author?: string; date?: string } {
    for (const block of document.blocks) {
      if (block.type === 'meta') {
        const meta = block as MetaBlock;
        return {
          title: meta.props.title ? String(meta.props.title) : undefined,
          author: meta.props.author ? String(meta.props.author) : undefined,
          date: meta.props.date ? String(meta.props.date) : undefined,
        };
      }
    }
    return {};
  }

  private getThemeColors(theme: string): ThemeColors {
    const themes: Record<string, ThemeColors> = {
      default: {
        primary: 'FF2C3E50',
        accent: 'FF3498DB',
        background: 'FFFFFFFF',
      },
      corporate: {
        primary: 'FF1A365D',
        accent: 'FF2B6CB0',
        background: 'FFFFFFFF',
      },
      academic: {
        primary: 'FF2D3748',
        accent: 'FF4A5568',
        background: 'FFFFFFFF',
      },
    };

    return themes[theme] || themes.default;
  }

  private getThinBorder(): Borders {
    return {
      top: 'thin',
      left: 'thin',
      bottom: 'thin',
      right: 'thin',
    };
  }

  private toRgb(color: string): string {
    const normalized = color.replace('#', '').toUpperCase();
    return normalized.length === 8 ? normalized.slice(2) : normalized;
  }
}
