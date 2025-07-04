import * as ExcelJS from 'exceljs';
import { OSFDocument, MetaBlock, DocBlock, SlideBlock, SheetBlock } from 'omniscript-parser';
import { Converter, ConverterOptions, ConversionResult } from './types';

export class XLSXConverter implements Converter {
  getSupportedFormats(): string[] {
    return ['xlsx'];
  }

  async convert(document: OSFDocument, options: ConverterOptions = {}): Promise<ConversionResult> {
    const workbook = new ExcelJS.Workbook();
    
    // Set workbook metadata
    this.setWorkbookMetadata(workbook, document);
    
    // Process document blocks
    await this.generateWorksheets(workbook, document, options);
    
    // Ensure at least one worksheet exists
    if (workbook.worksheets.length === 0) {
      this.createSummaryWorksheet(workbook, document, options);
    }
    
    const buffer = await workbook.xlsx.writeBuffer();
    
    return {
      buffer: Buffer.from(buffer),
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      extension: 'xlsx'
    };
  }

  private setWorkbookMetadata(workbook: ExcelJS.Workbook, document: OSFDocument): void {
    const meta = this.getMetadata(document);
    
    workbook.creator = meta.author || 'OmniScript OSF';
    workbook.created = meta.date ? new Date(meta.date) : new Date();
    workbook.modified = new Date();
    workbook.company = 'Generated by OmniScript';
    workbook.title = meta.title || 'OSF Workbook';
    workbook.description = 'Generated from OSF document';
  }

  private async generateWorksheets(workbook: ExcelJS.Workbook, document: OSFDocument, options: ConverterOptions): Promise<void> {
    let worksheetIndex = 1;
    let hasSheets = false;
    
    for (const block of document.blocks) {
      switch (block.type) {
        case 'sheet':
          this.createSheetWorksheet(workbook, block as SheetBlock, options);
          hasSheets = true;
          break;
        case 'doc':
        case 'slide':
          this.createContentWorksheet(workbook, block, options, `Content_${worksheetIndex}`);
          worksheetIndex++;
          break;
        case 'meta':
          // Metadata will be included in summary if no sheets exist
          break;
      }
    }
    
    // If no dedicated sheets exist, create a summary worksheet
    if (!hasSheets) {
      this.createSummaryWorksheet(workbook, document, options);
    }
  }

  private createSheetWorksheet(workbook: ExcelJS.Workbook, sheet: SheetBlock, options: ConverterOptions): void {
    const worksheetName = this.sanitizeWorksheetName(sheet.name || 'Sheet');
    const worksheet = workbook.addWorksheet(worksheetName);
    
    // Configure worksheet styling
    this.applyWorksheetStyling(worksheet, options);
    
    let currentRow = 1;
    
    // Add sheet title
    if (sheet.name) {
      const titleCell = worksheet.getCell(currentRow, 1);
      titleCell.value = sheet.name;
      titleCell.font = { bold: true, size: 16, color: { argb: 'FF2C3E50' } };
      titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF8F9FA' }
      };
      
      // Merge cells for title if we have multiple columns
      const colCount = this.getColumnCount(sheet);
      if (colCount > 1) {
        worksheet.mergeCells(currentRow, 1, currentRow, colCount);
      }
      
      currentRow += 2;
    }
    
    // Add column headers
    if (sheet.cols) {
      const cols = Array.isArray(sheet.cols) 
        ? sheet.cols 
        : String(sheet.cols).replace(/[[\]]/g, '').split(',').map(s => s.trim());
      
      cols.forEach((col, index) => {
        const cell = worksheet.getCell(currentRow, index + 1);
        cell.value = col;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF3498DB' }
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
      
      currentRow++;
    }
    
    // Add data rows
    if (sheet.data) {
      this.populateSheetData(worksheet, sheet, currentRow, options);
    }
    
    // Apply formulas if any
    if (sheet.formulas && sheet.formulas.length > 0) {
      this.applyFormulas(worksheet, sheet.formulas);
    }
    
    // Auto-size columns
    this.autoSizeColumns(worksheet);
  }

  private createContentWorksheet(workbook: ExcelJS.Workbook, block: any, options: ConverterOptions, name: string): void {
    const worksheet = workbook.addWorksheet(name);
    this.applyWorksheetStyling(worksheet, options);
    
    let currentRow = 1;
    
    if (block.type === 'doc') {
      const docBlock = block as DocBlock;
      currentRow = this.addDocContentToWorksheet(worksheet, docBlock, currentRow);
    } else if (block.type === 'slide') {
      const slideBlock = block as SlideBlock;
      currentRow = this.addSlideContentToWorksheet(worksheet, slideBlock, currentRow);
    }
    
    this.autoSizeColumns(worksheet);
  }

  private createSummaryWorksheet(workbook: ExcelJS.Workbook, document: OSFDocument, options: ConverterOptions): void {
    const worksheet = workbook.addWorksheet('Summary');
    this.applyWorksheetStyling(worksheet, options);
    
    let currentRow = 1;
    
    // Add document metadata
    const meta = this.getMetadata(document);
    if (meta.title || meta.author || meta.date) {
      const titleCell = worksheet.getCell(currentRow, 1);
      titleCell.value = 'Document Information';
      titleCell.font = { bold: true, size: 16, color: { argb: 'FF2C3E50' } };
      currentRow += 2;
      
      if (meta.title) {
        worksheet.getCell(currentRow, 1).value = 'Title:';
        worksheet.getCell(currentRow, 1).font = { bold: true };
        worksheet.getCell(currentRow, 2).value = meta.title;
        currentRow++;
      }
      
      if (meta.author) {
        worksheet.getCell(currentRow, 1).value = 'Author:';
        worksheet.getCell(currentRow, 1).font = { bold: true };
        worksheet.getCell(currentRow, 2).value = meta.author;
        currentRow++;
      }
      
      if (meta.date) {
        worksheet.getCell(currentRow, 1).value = 'Date:';
        worksheet.getCell(currentRow, 1).font = { bold: true };
        worksheet.getCell(currentRow, 2).value = meta.date;
        currentRow++;
      }
      
      currentRow += 2;
    }
    
    // Add content summary
    const contentBlocks = document.blocks.filter(b => b.type === 'doc' || b.type === 'slide');
    if (contentBlocks.length > 0) {
      const summaryTitleCell = worksheet.getCell(currentRow, 1);
      summaryTitleCell.value = 'Content Summary';
      summaryTitleCell.font = { bold: true, size: 14, color: { argb: 'FF2C3E50' } };
      currentRow += 2;
      
      // Add headers
      worksheet.getCell(currentRow, 1).value = 'Type';
      worksheet.getCell(currentRow, 2).value = 'Title/Content Preview';
      worksheet.getRow(currentRow).font = { bold: true };
      currentRow++;
      
      for (const block of contentBlocks) {
        worksheet.getCell(currentRow, 1).value = block.type.toUpperCase();
        
        if (block.type === 'slide') {
          const slide = block as SlideBlock;
          worksheet.getCell(currentRow, 2).value = slide.title || 'Untitled Slide';
        } else if (block.type === 'doc') {
          const doc = block as DocBlock;
          const preview = this.getContentPreview(doc.content || '');
          worksheet.getCell(currentRow, 2).value = preview;
        }
        
        currentRow++;
      }
    }
    
    this.autoSizeColumns(worksheet);
  }

  private populateSheetData(worksheet: ExcelJS.Worksheet, sheet: SheetBlock, startRow: number, options: ConverterOptions): void {
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
          const cell = worksheet.getCell(startRow + r - 1, c);
          
          // Handle different data types
          if (typeof value === 'number') {
            cell.value = value;
            cell.numFmt = '#,##0.00';
          } else if (typeof value === 'boolean') {
            cell.value = value;
          } else {
            cell.value = String(value);
          }
          
          // Apply basic styling
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        }
      }
    }
  }

  private applyFormulas(worksheet: ExcelJS.Worksheet, formulas: Array<{ cell: [number, number]; expr: string }>): void {
    for (const formula of formulas) {
      const [row, col] = formula.cell;
      const cell = worksheet.getCell(row, col);
      
      // Convert OSF formula to Excel formula
      let excelFormula = formula.expr;
      if (!excelFormula.startsWith('=')) {
        excelFormula = '=' + excelFormula;
      }
      
      // Convert cell references if needed (basic conversion)
      excelFormula = this.convertToExcelFormula(excelFormula);
      
      cell.value = { formula: excelFormula };
      
      // Style formula cells
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF0F8FF' }
      };
      cell.font = { italic: true };
    }
  }

  private addDocContentToWorksheet(worksheet: ExcelJS.Worksheet, doc: DocBlock, startRow: number): number {
    let currentRow = startRow;
    const content = doc.content || '';
    
    // Add title
    const titleCell = worksheet.getCell(currentRow, 1);
    titleCell.value = 'Document Content';
    titleCell.font = { bold: true, size: 14, color: { argb: 'FF2C3E50' } };
    currentRow += 2;
    
    // Split content into paragraphs and add to cells
    const paragraphs = content.split('\n\n');
    for (const paragraph of paragraphs) {
      if (paragraph.trim()) {
        const cell = worksheet.getCell(currentRow, 1);
        cell.value = paragraph.trim();
        cell.alignment = { wrapText: true, vertical: 'top' };
        worksheet.getRow(currentRow).height = Math.max(15, Math.min(100, paragraph.length / 10));
        currentRow++;
      }
    }
    
    return currentRow + 1;
  }

  private addSlideContentToWorksheet(worksheet: ExcelJS.Worksheet, slide: SlideBlock, startRow: number): number {
    let currentRow = startRow;
    
    // Add slide title
    if (slide.title) {
      const titleCell = worksheet.getCell(currentRow, 1);
      titleCell.value = slide.title;
      titleCell.font = { bold: true, size: 14, color: { argb: 'FF2C3E50' } };
      currentRow += 2;
    }
    
    // Add slide content
    if (slide.content) {
      for (const contentBlock of slide.content) {
        if (contentBlock.type === 'unordered_list') {
          for (const item of contentBlock.items) {
            const itemText = item.content.map(this.extractText).join('');
            const cell = worksheet.getCell(currentRow, 1);
            cell.value = `• ${itemText}`;
            currentRow++;
          }
        } else if (contentBlock.type === 'paragraph') {
          const paragraphText = contentBlock.content.map(this.extractText).join('');
          const cell = worksheet.getCell(currentRow, 1);
          cell.value = paragraphText;
          cell.alignment = { wrapText: true };
          currentRow++;
        }
      }
    }
    
    return currentRow + 1;
  }

  private applyWorksheetStyling(worksheet: ExcelJS.Worksheet, options: ConverterOptions): void {
    // Set default column width
    worksheet.columns = [
      { width: 20 },
      { width: 30 },
      { width: 15 },
      { width: 15 }
    ];
    
    // Apply theme-based styling if specified
    const theme = options.theme || 'default';
    const themeColors = this.getThemeColors(theme);
    
    // Set worksheet tab color based on theme
    worksheet.properties.tabColor = { argb: themeColors.accent };
  }

  private autoSizeColumns(worksheet: ExcelJS.Worksheet): void {
    worksheet.columns.forEach((column, index) => {
      let maxLength = 10; // Minimum width
      
      worksheet.eachRow({ includeEmpty: false }, (row) => {
        const cell = row.getCell(index + 1);
        if (cell.value) {
          const length = String(cell.value).length;
          maxLength = Math.max(maxLength, Math.min(50, length + 2));
        }
      });
      
      column.width = maxLength;
    });
  }

  private getColumnCount(sheet: SheetBlock): number {
    if (sheet.cols) {
      const cols = Array.isArray(sheet.cols) 
        ? sheet.cols 
        : String(sheet.cols).replace(/[[\]]/g, '').split(',').map(s => s.trim());
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
    return formula
      .replace(/\(\s*(\d+),\s*(\d+)\s*\)/g, (match, row, col) => {
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
    return name
      .replace(/[\\\/\*\?\[\]:]/g, '_')
      .substring(0, 31); // Max 31 characters
  }

  private getContentPreview(content: string): string {
    const preview = content
      .replace(/[#*`]/g, '')
      .replace(/\n+/g, ' ')
      .trim();
    return preview.length > 50 ? preview.substring(0, 47) + '...' : preview;
  }

  private extractText(run: any): string {
    if (typeof run === 'string') return run;
    if (run.type === 'link') return run.text;
    if (run.type === 'image') return run.alt || '';
    if (run.text) return run.text;
    return '';
  }

  private getMetadata(document: OSFDocument): { title?: string; author?: string; date?: string } {
    for (const block of document.blocks) {
      if (block.type === 'meta') {
        const meta = block as MetaBlock;
        return {
          title: meta.props.title ? String(meta.props.title) : undefined,
          author: meta.props.author ? String(meta.props.author) : undefined,
          date: meta.props.date ? String(meta.props.date) : undefined
        };
      }
    }
    return {};
  }

  private getThemeColors(theme: string): any {
    const themes: Record<string, any> = {
      default: {
        primary: 'FF2C3E50',
        accent: 'FF3498DB',
        background: 'FFFFFFFF'
      },
      corporate: {
        primary: 'FF1A365D',
        accent: 'FF2B6CB0',
        background: 'FFFFFFFF'
      },
      academic: {
        primary: 'FF2D3748',
        accent: 'FF4A5568',
        background: 'FFFFFFFF'
      }
    };
    
    return themes[theme] || themes.default;
  }
}