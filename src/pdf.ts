import puppeteer from 'puppeteer';
import { OSFDocument, MetaBlock, DocBlock, SlideBlock, SheetBlock } from 'omniscript-parser';
import { Converter, ConverterOptions, ConversionResult } from './types';

export class PDFConverter implements Converter {
  getSupportedFormats(): string[] {
    return ['pdf'];
  }

  async convert(document: OSFDocument, options: ConverterOptions = {}): Promise<ConversionResult> {
    const html = this.generateHTML(document, options);
    
    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
    });
    
    try {
      const page = await browser.newPage();
      await page.setContent(html, { waitUntil: 'networkidle0' });
      
      const pdfOptions = {
        format: options.pageSize || 'A4' as const,
        landscape: options.orientation === 'landscape',
        margin: options.margins || {
          top: '1in',
          right: '1in',
          bottom: '1in',
          left: '1in'
        },
        printBackground: true
      };
      
      const buffer = await page.pdf(pdfOptions);
      
      return {
        buffer: Buffer.from(buffer),
        mimeType: 'application/pdf',
        extension: 'pdf'
      };
    } finally {
      await browser.close();
    }
  }

  private generateHTML(document: OSFDocument, options: ConverterOptions): string {
    const theme = options.theme || 'default';
    const styles = this.getThemeStyles(theme);
    
    let html = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>OSF Document</title>
    <style>
        ${styles}
    </style>
</head>
<body>
`;

    for (const block of document.blocks) {
      switch (block.type) {
        case 'meta':
          html += this.renderMetaBlock(block as MetaBlock, options);
          break;
        case 'doc':
          html += this.renderDocBlock(block as DocBlock, options);
          break;
        case 'slide':
          html += this.renderSlideBlock(block as SlideBlock, options);
          break;
        case 'sheet':
          html += this.renderSheetBlock(block as SheetBlock, options);
          break;
      }
    }

    html += '</body></html>';
    return html;
  }

  private renderMetaBlock(meta: MetaBlock, options: ConverterOptions): string {
    if (!options.includeMetadata) return '';
    
    let html = '<div class="meta-block">';
    
    if (meta.props.title) {
      html += `<h1 class="document-title">${this.escapeHtml(String(meta.props.title))}</h1>`;
    }
    
    if (meta.props.author) {
      html += `<p class="document-author">By: ${this.escapeHtml(String(meta.props.author))}</p>`;
    }
    
    if (meta.props.date) {
      html += `<p class="document-date">Date: ${this.escapeHtml(String(meta.props.date))}</p>`;
    }
    
    html += '</div>';
    return html;
  }

  private renderDocBlock(doc: DocBlock, options: ConverterOptions): string {
    const content = doc.content || '';
    
    // Basic Markdown-to-HTML conversion
    let html = content
      .replace(/^# (.+)$/gm, '<h1>$1</h1>')
      .replace(/^## (.+)$/gm, '<h2>$1</h2>')
      .replace(/^### (.+)$/gm, '<h3>$1</h3>')
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.+?)\*/g, '<em>$1</em>')
      .replace(/`(.+?)`/g, '<code>$1</code>')
      .replace(/\n\n/g, '</p><p>')
      .replace(/\n/g, '<br>');
    
    return `<div class="doc-block"><p>${html}</p></div>`;
  }

  private renderSlideBlock(slide: SlideBlock, options: ConverterOptions): string {
    let html = '<div class="slide-block">';
    
    if (slide.title) {
      html += `<h2 class="slide-title">${this.escapeHtml(slide.title)}</h2>`;
    }
    
    if (slide.content) {
      html += '<div class="slide-content">';
      
      for (const contentBlock of slide.content) {
        if (contentBlock.type === 'unordered_list') {
          html += '<ul>';
          for (const item of contentBlock.items) {
            const itemText = item.content.map(this.extractText).join('');
            html += `<li>${this.escapeHtml(itemText)}</li>`;
          }
          html += '</ul>';
        } else if (contentBlock.type === 'paragraph') {
          const paragraphText = contentBlock.content.map(this.extractText).join('');
          html += `<p>${this.escapeHtml(paragraphText)}</p>`;
        }
      }
      
      html += '</div>';
    }
    
    html += '</div>';
    return html;
  }

  private renderSheetBlock(sheet: SheetBlock, options: ConverterOptions): string {
    let html = '<div class="sheet-block">';
    
    if (sheet.name) {
      html += `<h3 class="sheet-title">${this.escapeHtml(sheet.name)}</h3>`;
    }
    
    html += '<table class="sheet-table">';
    
    // Render column headers
    if (sheet.cols) {
      const cols = Array.isArray(sheet.cols) 
        ? sheet.cols 
        : String(sheet.cols).replace(/[[\]]/g, '').split(',').map(s => s.trim());
      
      html += '<thead><tr>';
      for (const col of cols) {
        html += `<th>${this.escapeHtml(col)}</th>`;
      }
      html += '</tr></thead>';
    }
    
    // Render data rows
    if (sheet.data) {
      html += '<tbody>';
      
      // Calculate table dimensions
      const coords = Object.keys(sheet.data).map(k => k.split(',').map(Number));
      const maxRow = Math.max(...coords.map(c => c[0]));
      const maxCol = Math.max(...coords.map(c => c[1]));
      
      for (let r = 1; r <= maxRow; r++) {
        html += '<tr>';
        for (let c = 1; c <= maxCol; c++) {
          const key = `${r},${c}`;
          const value = sheet.data[key] || '';
          html += `<td>${this.escapeHtml(String(value))}</td>`;
        }
        html += '</tr>';
      }
      
      html += '</tbody>';
    }
    
    html += '</table></div>';
    return html;
  }

  private extractText(run: any): string {
    if (typeof run === 'string') return run;
    if (run.type === 'link') return run.text;
    if (run.type === 'image') return run.alt || '';
    if (run.text) return run.text;
    return '';
  }

  private escapeHtml(text: string): string {
    const div = { innerHTML: '', textContent: text } as any;
    return div.innerHTML || text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  private getThemeStyles(theme: string): string {
    const baseStyles = `
      body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        line-height: 1.6;
        color: #333;
        max-width: 800px;
        margin: 0 auto;
        padding: 40px 20px;
      }
      
      .meta-block {
        text-align: center;
        margin-bottom: 40px;
        padding-bottom: 20px;
        border-bottom: 2px solid #eee;
      }
      
      .document-title {
        font-size: 2.5em;
        margin-bottom: 10px;
        color: #2c3e50;
      }
      
      .document-author, .document-date {
        color: #7f8c8d;
        margin: 5px 0;
      }
      
      .doc-block {
        margin: 30px 0;
      }
      
      .slide-block {
        margin: 40px 0;
        padding: 30px;
        border: 1px solid #ddd;
        border-radius: 8px;
        page-break-inside: avoid;
      }
      
      .slide-title {
        color: #2c3e50;
        border-bottom: 2px solid #3498db;
        padding-bottom: 10px;
        margin-bottom: 20px;
      }
      
      .sheet-block {
        margin: 30px 0;
      }
      
      .sheet-title {
        color: #2c3e50;
        margin-bottom: 15px;
      }
      
      .sheet-table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
      }
      
      .sheet-table th,
      .sheet-table td {
        border: 1px solid #ddd;
        padding: 8px 12px;
        text-align: left;
      }
      
      .sheet-table th {
        background-color: #f8f9fa;
        font-weight: 600;
      }
      
      h1, h2, h3 {
        color: #2c3e50;
      }
      
      ul {
        padding-left: 20px;
      }
      
      li {
        margin: 8px 0;
      }
      
      code {
        background-color: #f8f9fa;
        padding: 2px 4px;
        border-radius: 3px;
        font-family: 'Monaco', 'Menlo', monospace;
      }
    `;

    switch (theme) {
      case 'corporate':
        return baseStyles + `
          .document-title { color: #1a365d; }
          .slide-title { border-bottom-color: #2b6cb0; }
          .slide-block { border-color: #2b6cb0; }
        `;
      case 'academic':
        return baseStyles + `
          body { font-family: 'Times New Roman', serif; }
          .document-title { color: #2d3748; }
          .slide-title { border-bottom-color: #4a5568; }
        `;
      default:
        return baseStyles;
    }
  }
}