import { readFileSync } from 'fs';
import puppeteer from 'puppeteer';
import {
  OSFDocument,
  MetaBlock,
  DocBlock,
  SlideBlock,
  SheetBlock,
  ChartBlock,
  DiagramBlock,
  OSFCodeBlock,
  TableBlock,
} from 'omniscript-parser';
import { Converter, ConverterOptions, ConversionResult } from './types';

export class PDFConverter implements Converter {
  private chartIdCounter = 0;
  private static chartJsSource: string | null = null;
  private static mermaidSource: string | null = null;

  private static loadAsset(modulePath: string): string | null {
    try {
      const resolved = require.resolve(modulePath);
      return readFileSync(resolved, 'utf8');
    } catch {
      return null;
    }
  }

  private getChartJsSource(): string | null {
    if (PDFConverter.chartJsSource === null) {
      PDFConverter.chartJsSource = PDFConverter.loadAsset('chart.js/dist/chart.umd.min.js');
    }
    return PDFConverter.chartJsSource;
  }

  private getMermaidSource(): string | null {
    if (PDFConverter.mermaidSource === null) {
      PDFConverter.mermaidSource = PDFConverter.loadAsset('mermaid/dist/mermaid.min.js');
    }
    return PDFConverter.mermaidSource;
  }

  getSupportedFormats(): string[] {
    return ['pdf'];
  }

  async convert(document: OSFDocument, options: ConverterOptions = {}): Promise<ConversionResult> {
    const html = this.generateHTML(document, options);

    const timeoutMs = options.timeoutMs ?? Number(process.env.OSF_PDF_TIMEOUT_MS || 30000);
    const envNoSandbox = process.env.OSF_PUPPETEER_NO_SANDBOX;
    const disableSandbox = envNoSandbox
      ? envNoSandbox === 'true' || envNoSandbox === '1'
      : typeof process.getuid === 'function' && process.getuid() === 0;
    const args = ['--disable-dev-shm-usage', ...(options.puppeteerArgs || [])];
    if (disableSandbox) {
      args.push('--no-sandbox', '--disable-setuid-sandbox');
    }

    const browser = await puppeteer.launch({
      headless: true,
      args,
    });

    try {
      const page = await browser.newPage();
      const setContent = page.setContent(html, { waitUntil: 'load' });
      await this.withTimeout(setContent, timeoutMs, 'page.setContent');

      const pdfOptions = {
        format: options.pageSize || ('A4' as const),
        landscape: options.orientation === 'landscape',
        margin: options.margins || {
          top: '1in',
          right: '1in',
          bottom: '1in',
          left: '1in',
        },
        printBackground: true,
      };

      const buffer = await this.withTimeout(page.pdf(pdfOptions), timeoutMs, 'page.pdf');

      return {
        buffer: Buffer.from(buffer),
        mimeType: 'application/pdf',
        extension: 'pdf',
      };
    } finally {
      await browser.close();
    }
  }

  private generateHTML(document: OSFDocument, options: ConverterOptions): string {
    const theme = options.theme || 'default';
    const styles = this.getThemeStyles(theme);
    const hasChart = document.blocks.some(block => block.type === 'chart');
    const hasDiagram = document.blocks.some(block => block.type === 'diagram');
    const chartSource = hasChart ? this.getChartJsSource() : null;
    const mermaidSource = hasDiagram ? this.getMermaidSource() : null;

    let html = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>OSF Document</title>
    <style>
        ${styles}
    </style>
    ${chartSource ? `<script>${chartSource}</script>` : ''}
    ${
      mermaidSource
        ? `<script>${mermaidSource}</script>
    <script>
      window.addEventListener('load', () => {
        if (window.mermaid) {
          window.mermaid.initialize({ startOnLoad: true, theme: '${this.escapeHtml(theme)}' });
          window.mermaid.run();
        }
      });
    </script>`
        : ''
    }
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
        case 'chart':
          html += this.renderChartBlock(block as ChartBlock, options);
          break;
        case 'diagram':
          html += this.renderDiagramBlock(block as DiagramBlock, options);
          break;
        case 'osfcode':
          html += this.renderCodeBlock(block as OSFCodeBlock, options);
          break;
        case 'table':
          html += this.renderTableBlock(block as TableBlock, options);
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
    const html = this.renderMarkdown(content);
    return `<div class="doc-block">${html}</div>`;
  }

  private renderSlideBlock(slide: SlideBlock, options: ConverterOptions): string {
    let html = '<div class="slide-block">';

    if (slide.title) {
      html += `<h2 class="slide-title">${this.escapeHtml(slide.title)}</h2>`;
    }

    if (slide.content && slide.content.length > 0) {
      html += '<div class="slide-content">';

      for (const contentBlock of slide.content) {
        if (contentBlock.type === 'unordered_list') {
          html += '<ul>';
          for (const item of contentBlock.items) {
            const itemText = item.content.map(this.extractText).join('');
            html += `<li>${this.escapeHtml(itemText)}</li>`;
          }
          html += '</ul>';
        } else if (contentBlock.type === 'ordered_list') {
          html += '<ol>';
          for (const item of contentBlock.items) {
            const itemText = item.content.map(this.extractText).join('');
            html += `<li>${this.escapeHtml(itemText)}</li>`;
          }
          html += '</ol>';
        } else if (contentBlock.type === 'blockquote') {
          html += '<blockquote>';
          for (const paragraph of contentBlock.content) {
            const paragraphText = paragraph.content.map(this.extractText).join('');
            html += `<p>${this.escapeHtml(paragraphText)}</p>`;
          }
          html += '</blockquote>';
        } else if (contentBlock.type === 'paragraph') {
          const paragraphText = contentBlock.content.map(this.extractText).join('');
          html += `<p>${this.escapeHtml(paragraphText)}</p>`;
        } else if (contentBlock.type === 'code') {
          html += `<pre><code>${this.escapeHtml(contentBlock.content)}</code></pre>`;
        } else if (contentBlock.type === 'image') {
          html += `<img src="${this.escapeHtml(contentBlock.url)}" alt="${this.escapeHtml(
            contentBlock.alt
          )}" style="max-width: 100%; height: auto;" />`;
        }
      }

      html += '</div>';
    } else if (slide.bullets && slide.bullets.length > 0) {
      html += '<div class="slide-content"><ul>';
      for (const bullet of slide.bullets) {
        html += `<li>${this.escapeHtml(bullet)}</li>`;
      }
      html += '</ul></div>';
    }

    html += '</div>';
    return html;
  }

  private renderMarkdown(content: string): string {
    const lines = content.split(/\r?\n/);
    let html = '';
    let paragraph: string[] = [];
    let listItems: string[] = [];
    let blockquoteLines: string[] = [];

    const flushParagraph = () => {
      if (paragraph.length > 0) {
        html += `<p>${this.renderInlineMarkdown(paragraph.join(' '))}</p>`;
        paragraph = [];
      }
    };

    const flushList = () => {
      if (listItems.length > 0) {
        html += `<ul>${listItems.map(item => `<li>${this.renderInlineMarkdown(item)}</li>`).join('')}</ul>`;
        listItems = [];
      }
    };

    const flushBlockquote = () => {
      if (blockquoteLines.length > 0) {
        html += `<blockquote>${blockquoteLines
          .map(line => `<p>${this.renderInlineMarkdown(line)}</p>`)
          .join('')}</blockquote>`;
        blockquoteLines = [];
      }
    };

    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line) {
        flushParagraph();
        flushList();
        flushBlockquote();
        continue;
      }

      const headingMatch = /^(#{1,3})\s+(.+)$/.exec(line);
      if (headingMatch) {
        flushParagraph();
        flushList();
        flushBlockquote();
        const level = headingMatch[1].length;
        html += `<h${level}>${this.renderInlineMarkdown(headingMatch[2])}</h${level}>`;
        continue;
      }

      const listMatch = /^[-*]\s+(.+)$/.exec(line);
      if (listMatch) {
        flushParagraph();
        flushBlockquote();
        listItems.push(listMatch[1]);
        continue;
      }

      const quoteMatch = /^>\s?(.+)$/.exec(line);
      if (quoteMatch) {
        flushParagraph();
        flushList();
        blockquoteLines.push(quoteMatch[1]);
        continue;
      }

      if (listItems.length > 0) flushList();
      if (blockquoteLines.length > 0) flushBlockquote();
      paragraph.push(line);
    }

    flushParagraph();
    flushList();
    flushBlockquote();

    return html;
  }

  private renderInlineMarkdown(text: string): string {
    let html = this.escapeHtml(text);
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');
    html = html.replace(/`(.+?)`/g, '<code>$1</code>');
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
        : String(sheet.cols)
            .replace(/[[\]]/g, '')
            .split(',')
            .map(s => s.trim());

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

  private renderTableBlock(table: TableBlock, options: ConverterOptions): string {
    const tableStyle = table.style || 'bordered';
    let html = '<div class="table-block">';

    if (table.caption) {
      html += `<p class="table-caption">${this.escapeHtml(table.caption)}</p>`;
    }

    html += `<table class="osf-table ${this.escapeHtml(tableStyle)}">`;
    html += '<thead><tr>';

    table.headers.forEach((header, index) => {
      const align = table.alignment?.[index] || 'left';
      html += `<th style="text-align: ${align};">${this.escapeHtml(header)}</th>`;
    });

    html += '</tr></thead><tbody>';

    table.rows.forEach(row => {
      html += '<tr>';
      row.cells.forEach((cell, index) => {
        const align = table.alignment?.[index] || 'left';
        html += `<td style="text-align: ${align};">${this.escapeHtml(cell.text)}</td>`;
      });
      html += '</tr>';
    });

    html += '</tbody></table></div>';
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
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  private renderChartBlock(chart: ChartBlock, options: ConverterOptions): string {
    let html = '<div class="chart-block" style="margin: 40px 0; page-break-inside: avoid;">';

    if (chart.title) {
      html += `<h3 style="text-align: center; margin-bottom: 20px;">${this.escapeHtml(chart.title)}</h3>`;
    }

    const chartId = `chart-${++this.chartIdCounter}`;
    html += '<div class="chart-container" style="width: 100%; height: 400px; position: relative;">';
    html += `<canvas id="${chartId}"></canvas>`;
    html += '</div>';

    const chartPayload = {
      labels: chart.data.map((d: any) => d.label),
      datasets: chart.data.map((d: any, i: number) => ({
        label: d.label,
        data: d.values,
        backgroundColor: chart.options?.colors?.[i] || undefined,
      })),
    };

    html += `<script>(function(){\n`;
    html += `  const canvas = document.getElementById('${chartId}');\n`;
    html += `  if (!canvas || !window.Chart) return;\n`;
    html += `  const chartData = ${JSON.stringify(chartPayload)};\n`;
    html += `  const chartConfig = {\n`;
    html += `    type: '${chart.chartType}',\n`;
    html += `    data: chartData,\n`;
    html += `    options: {\n`;
    html += `      responsive: true,\n`;
    html += `      plugins: { legend: { display: ${chart.options?.legend !== false} } }\n`;
    html += `    }\n`;
    html += `  };\n`;
    html += `  new window.Chart(canvas, chartConfig);\n`;
    html += `})();</script>`;
    html += '</div>';

    return html;
  }

  private renderDiagramBlock(diagram: DiagramBlock, options: ConverterOptions): string {
    let html = '<div class="diagram-block" style="margin: 40px 0; page-break-inside: avoid;">';

    if (diagram.title) {
      html += `<h3 style="text-align: center; margin-bottom: 20px;">${this.escapeHtml(diagram.title)}</h3>`;
    }

    html += '<div class="mermaid" style="text-align: center;">';
    html += this.escapeHtml(diagram.code);
    html += '</div>';
    html += '</div>';

    return html;
  }

  private renderCodeBlock(code: OSFCodeBlock, options: ConverterOptions): string {
    let html = '<div class="code-block" style="margin: 40px 0; page-break-inside: avoid;">';

    if (code.caption) {
      html += `<p class="code-caption" style="font-style: italic; margin-bottom: 10px;">${this.escapeHtml(code.caption)}</p>`;
    }

    html +=
      '<pre style="background: #f6f8fa; padding: 16px; border-radius: 6px; overflow-x: auto; border: 1px solid #e1e4e8;">';
    html += `<code class="language-${code.language}">`;

    const lines = code.code.split('\n');
    lines.forEach((line: string, index: number) => {
      const lineNum = index + 1;
      const isHighlighted = code.highlight && code.highlight.includes(lineNum);
      const lineStyle = isHighlighted ? 'background: #fff3cd;' : '';

      if (code.lineNumbers) {
        html += `<span style="${lineStyle}"><span style="color: #6e7781; margin-right: 16px; user-select: none;">${lineNum.toString().padStart(3, ' ')}</span>${this.escapeHtml(line)}</span>\n`;
      } else {
        html += `<span style="${lineStyle}">${this.escapeHtml(line)}</span>\n`;
      }
    });

    html += '</code></pre>';
    html += '</div>';

    return html;
  }

  private async withTimeout<T>(promise: Promise<T>, timeoutMs: number, label: string): Promise<T> {
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) {
      return promise;
    }

    let timer: NodeJS.Timeout | undefined;
    const timeoutPromise = new Promise<T>((_, reject) => {
      timer = setTimeout(() => {
        reject(new Error(`PDF conversion timed out during ${label} after ${timeoutMs}ms`));
      }, timeoutMs);
    });

    try {
      return await Promise.race([promise, timeoutPromise]);
    } finally {
      if (timer) clearTimeout(timer);
    }
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

      .table-block {
        margin: 30px 0;
      }

      .table-caption {
        font-style: italic;
        color: #6b7280;
        margin-bottom: 8px;
      }

      .osf-table {
        width: 100%;
        border-collapse: collapse;
        margin: 16px 0;
      }

      .osf-table th,
      .osf-table td {
        border: 1px solid #ddd;
        padding: 8px 12px;
      }

      .osf-table.striped tbody tr:nth-child(odd) {
        background-color: #f8f9fa;
      }

      .osf-table.minimal th,
      .osf-table.minimal td {
        border: none;
        border-bottom: 1px solid #e5e7eb;
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

      blockquote {
        border-left: 4px solid #d1d5db;
        margin: 16px 0;
        padding: 8px 16px;
        color: #4b5563;
        font-style: italic;
        background: #f9fafb;
      }
    `;

    switch (theme) {
      case 'corporate':
        return (
          baseStyles +
          `
          .document-title { color: #1a365d; }
          .slide-title { border-bottom-color: #2b6cb0; }
          .slide-block { border-color: #2b6cb0; }
          .sheet-table th { background-color: #ebf4ff; }
        `
        );
      case 'academic':
        return (
          baseStyles +
          `
          body { font-family: 'Times New Roman', serif; }
          .document-title { color: #2d3748; }
          .slide-title { border-bottom-color: #4a5568; }
          .sheet-table th { background-color: #f7fafc; }
        `
        );
      case 'modern':
        return (
          baseStyles +
          `
          body { font-family: 'Segoe UI', system-ui, sans-serif; }
          .document-title { color: #6366f1; }
          .slide-title { border-bottom-color: #06b6d4; }
          .slide-block { border-color: #8b5cf6; }
          .sheet-table th { background-color: #f9fafb; }
        `
        );
      case 'dark':
        return (
          baseStyles +
          `
          body { background-color: #1f2937; color: #e5e7eb; }
          .document-title { color: #f59e0b; }
          .document-author, .document-date { color: #fcd34d; }
          .slide-title { border-bottom-color: #10b981; color: #f59e0b; }
          .slide-block { border-color: #4b5563; background-color: #374151; }
          h1, h2, h3 { color: #fcd34d; }
          .sheet-table { border-color: #4b5563; }
          .sheet-table th { background-color: #374151; color: #f59e0b; }
        `
        );
      case 'minimal':
        return (
          baseStyles +
          `
          body { font-family: 'Helvetica', 'Arial', sans-serif; }
          .document-title { color: #000; }
          .slide-title { border-bottom-color: #000; }
          .slide-block { border-color: #ccc; }
          .sheet-table th { background-color: #fafafa; }
        `
        );
      case 'vibrant':
        return (
          baseStyles +
          `
          .document-title { color: #ec4899; }
          .slide-title { border-bottom-color: #8b5cf6; }
          .slide-block { border-color: #f472b6; }
          .sheet-table th { background-color: #fdf4ff; }
        `
        );
      case 'ocean':
        return (
          baseStyles +
          `
          .document-title { color: #0ea5e9; }
          .slide-title { border-bottom-color: #3b82f6; }
          .slide-block { border-color: #06b6d4; }
          .sheet-table th { background-color: #f0f9ff; }
        `
        );
      case 'forest':
        return (
          baseStyles +
          `
          body { font-family: 'Georgia', serif; }
          .document-title { color: #059669; }
          .slide-title { border-bottom-color: #14b8a6; }
          .slide-block { border-color: #10b981; }
          .sheet-table th { background-color: #f0fdf4; }
        `
        );
      case 'sunset':
        return (
          baseStyles +
          `
          .document-title { color: #dc2626; }
          .slide-title { border-bottom-color: #fbbf24; }
          .slide-block { border-color: #f97316; }
          .sheet-table th { background-color: #fff7ed; }
        `
        );
      default:
        return baseStyles;
    }
  }
}
