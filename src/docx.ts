import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, HeadingLevel, AlignmentType, WidthType } from 'docx';
import { OSFDocument, MetaBlock, DocBlock, SlideBlock, SheetBlock } from 'omniscript-parser';
import { Converter, ConverterOptions, ConversionResult } from './types';

export class DOCXConverter implements Converter {
  getSupportedFormats(): string[] {
    return ['docx'];
  }

  async convert(document: OSFDocument, options: ConverterOptions = {}): Promise<ConversionResult> {
    const docElements = this.generateDocumentElements(document, options);
    
    const doc = new Document({
      creator: 'OmniScript OSF',
      title: this.getDocumentTitle(document),
      description: 'Generated from OSF document',
      sections: [{
        properties: {},
        children: docElements
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    
    return {
      buffer,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      extension: 'docx'
    };
  }

  private generateDocumentElements(document: OSFDocument, options: ConverterOptions): any[] {
    const elements: any[] = [];

    for (const block of document.blocks) {
      switch (block.type) {
        case 'meta':
          if (options.includeMetadata !== false) {
            elements.push(...this.renderMetaBlock(block as MetaBlock));
          }
          break;
        case 'doc':
          elements.push(...this.renderDocBlock(block as DocBlock));
          break;
        case 'slide':
          elements.push(...this.renderSlideBlock(block as SlideBlock));
          break;
        case 'sheet':
          elements.push(...this.renderSheetBlock(block as SheetBlock));
          break;
      }
    }

    return elements;
  }

  private renderMetaBlock(meta: MetaBlock): any[] {
    const elements: any[] = [];

    if (meta.props.title) {
      elements.push(
        new Paragraph({
          children: [
            new TextRun({
              text: String(meta.props.title),
              bold: true,
              size: 32,
            })
          ],
          heading: HeadingLevel.TITLE,
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 }
        })
      );
    }

    const metaInfo: string[] = [];
    if (meta.props.author) metaInfo.push(`Author: ${meta.props.author}`);
    if (meta.props.date) metaInfo.push(`Date: ${meta.props.date}`);
    if (meta.props.theme) metaInfo.push(`Theme: ${meta.props.theme}`);

    if (metaInfo.length > 0) {
      elements.push(
        new Paragraph({
          children: [
            new TextRun({
              text: metaInfo.join(' | '),
              italics: true,
              color: '666666'
            })
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 600 }
        })
      );
    }

    return elements;
  }

  private renderDocBlock(doc: DocBlock): any[] {
    const content = doc.content || '';
    const elements: any[] = [];

    // Split content into lines and process markdown-like syntax
    const lines = content.split('\n');
    let currentParagraph: string[] = [];

    for (const line of lines) {
      const trimmed = line.trim();
      
      if (trimmed.startsWith('# ')) {
        // Heading 1
        if (currentParagraph.length > 0) {
          elements.push(this.createParagraph(currentParagraph.join(' ')));
          currentParagraph = [];
        }
        elements.push(
          new Paragraph({
            children: [new TextRun({ text: trimmed.substring(2), bold: true, size: 28 })],
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 240, after: 120 }
          })
        );
      } else if (trimmed.startsWith('## ')) {
        // Heading 2
        if (currentParagraph.length > 0) {
          elements.push(this.createParagraph(currentParagraph.join(' ')));
          currentParagraph = [];
        }
        elements.push(
          new Paragraph({
            children: [new TextRun({ text: trimmed.substring(3), bold: true, size: 24 })],
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 100 }
          })
        );
      } else if (trimmed.startsWith('### ')) {
        // Heading 3
        if (currentParagraph.length > 0) {
          elements.push(this.createParagraph(currentParagraph.join(' ')));
          currentParagraph = [];
        }
        elements.push(
          new Paragraph({
            children: [new TextRun({ text: trimmed.substring(4), bold: true, size: 20 })],
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 160, after: 80 }
          })
        );
      } else if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
        // List item
        if (currentParagraph.length > 0) {
          elements.push(this.createParagraph(currentParagraph.join(' ')));
          currentParagraph = [];
        }
        const listText = trimmed.substring(2);
        elements.push(
          new Paragraph({
            children: [
              new TextRun({ text: '• ' }),
              ...this.parseInlineFormatting(listText)
            ],
            spacing: { after: 80 }
          })
        );
      } else if (trimmed === '') {
        // Empty line - end current paragraph
        if (currentParagraph.length > 0) {
          elements.push(this.createParagraph(currentParagraph.join(' ')));
          currentParagraph = [];
        }
      } else {
        // Regular text line
        currentParagraph.push(trimmed);
      }
    }

    // Handle any remaining paragraph content
    if (currentParagraph.length > 0) {
      elements.push(this.createParagraph(currentParagraph.join(' ')));
    }

    return elements;
  }

  private renderSlideBlock(slide: SlideBlock): any[] {
    const elements: any[] = [];

    // Add slide title
    if (slide.title) {
      elements.push(
        new Paragraph({
          children: [
            new TextRun({
              text: slide.title,
              bold: true,
              size: 24,
              color: '2E74B5'
            })
          ],
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 400, after: 200 }
        })
      );
    }

    // Add slide content
    if (slide.content) {
      for (const contentBlock of slide.content) {
        if (contentBlock.type === 'unordered_list') {
          for (const item of contentBlock.items) {
            const itemText = item.content.map(this.extractText).join('');
            elements.push(
              new Paragraph({
                children: [
                  new TextRun({ text: '• ' }),
                  ...this.parseInlineFormatting(itemText)
                ],
                spacing: { after: 100 }
              })
            );
          }
        } else if (contentBlock.type === 'paragraph') {
          const paragraphText = contentBlock.content.map(this.extractText).join('');
          elements.push(this.createParagraph(paragraphText));
        }
      }
    }

    return elements;
  }

  private renderSheetBlock(sheet: SheetBlock): any[] {
    const elements: any[] = [];

    // Add sheet title
    if (sheet.name) {
      elements.push(
        new Paragraph({
          children: [
            new TextRun({
              text: sheet.name,
              bold: true,
              size: 20,
              color: '2E74B5'
            })
          ],
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 300, after: 150 }
        })
      );
    }

    if (sheet.data) {
      // Calculate table dimensions
      const coords = Object.keys(sheet.data).map(k => k.split(',').map(Number));
      const maxRow = Math.max(...coords.map(c => c[0]));
      const maxCol = Math.max(...coords.map(c => c[1]));

      const tableRows: TableRow[] = [];

      // Add header row if columns are defined
      if (sheet.cols) {
        const cols = Array.isArray(sheet.cols) 
          ? sheet.cols 
          : String(sheet.cols).replace(/[[\]]/g, '').split(',').map(s => s.trim());
        
        const headerCells = cols.map(col => 
          new TableCell({
            children: [
              new Paragraph({
                children: [new TextRun({ text: col, bold: true })]
              })
            ],
            shading: { fill: 'E8E8E8' }
          })
        );
        
        tableRows.push(new TableRow({ children: headerCells }));
      }

      // Add data rows
      for (let r = 1; r <= maxRow; r++) {
        const cells: TableCell[] = [];
        for (let c = 1; c <= maxCol; c++) {
          const key = `${r},${c}`;
          const value = sheet.data[key] || '';
          cells.push(
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: String(value) })]
                })
              ]
            })
          );
        }
        tableRows.push(new TableRow({ children: cells }));
      }

      if (tableRows.length > 0) {
        const table = new Table({
          rows: tableRows,
          width: {
            size: 100,
            type: WidthType.PERCENTAGE
          }
        });
        elements.push(table);
      }
    }

    return elements;
  }

  private createParagraph(text: string): Paragraph {
    return new Paragraph({
      children: this.parseInlineFormatting(text),
      spacing: { after: 120 }
    });
  }

  private parseInlineFormatting(text: string): TextRun[] {
    const runs: TextRun[] = [];
    let currentIndex = 0;

    // Simple regex patterns for basic formatting
    const patterns = [
      { regex: /\*\*(.+?)\*\*/g, format: { bold: true } },
      { regex: /\*(.+?)\*/g, format: { italics: true } },
      { regex: /`(.+?)`/g, format: { font: 'Courier New' } }
    ];

    const matches: Array<{ index: number; length: number; text: string; format: any }> = [];

    // Find all formatting matches
    for (const pattern of patterns) {
      let match;
      while ((match = pattern.regex.exec(text)) !== null) {
        matches.push({
          index: match.index,
          length: match[0].length,
          text: match[1],
          format: pattern.format
        });
      }
    }

    // Sort matches by index
    matches.sort((a, b) => a.index - b.index);

    // Build text runs
    for (const match of matches) {
      // Add any plain text before this match
      if (match.index > currentIndex) {
        const plainText = text.substring(currentIndex, match.index);
        if (plainText) {
          runs.push(new TextRun({ text: plainText }));
        }
      }

      // Add formatted text
      runs.push(new TextRun({ text: match.text, ...match.format }));
      currentIndex = match.index + match.length;
    }

    // Add any remaining plain text
    if (currentIndex < text.length) {
      const remainingText = text.substring(currentIndex);
      if (remainingText) {
        runs.push(new TextRun({ text: remainingText }));
      }
    }

    // If no formatting was found, return the entire text as a single run
    if (runs.length === 0) {
      runs.push(new TextRun({ text }));
    }

    return runs;
  }

  private extractText(run: any): string {
    if (typeof run === 'string') return run;
    if (run.type === 'link') return run.text;
    if (run.type === 'image') return run.alt || '';
    if (run.text) return run.text;
    return '';
  }

  private getDocumentTitle(document: OSFDocument): string {
    for (const block of document.blocks) {
      if (block.type === 'meta') {
        const meta = block as MetaBlock;
        if (meta.props.title) {
          return String(meta.props.title);
        }
      }
    }
    return 'OSF Document';
  }
}