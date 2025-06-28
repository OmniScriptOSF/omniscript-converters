# omniscript-converters

Professional format converters for OmniScript Format (OSF) documents. Convert OSF documents to PDF, DOCX, PPTX, and XLSX with support for themes, styling, and advanced formatting.

## Features

- **PDF Generation** - High-quality PDF documents with custom themes
- **DOCX Export** - Microsoft Word documents with rich formatting
- **PPTX Creation** - PowerPoint presentations with automatic slide generation
- **XLSX Conversion** - Excel workbooks with formula support and data tables
- **Theme Support** - Multiple built-in themes (default, corporate, academic, modern)
- **TypeScript Support** - Full type definitions included

## Installation

```bash
npm install omniscript-converters
# or
pnpm add omniscript-converters
```

## Usage

### Basic Usage

```typescript
import { parse } from 'omniscript-parser';
import { PDFConverter, DOCXConverter, PPTXConverter, XLSXConverter } from 'omniscript-converters';

// Parse OSF document
const osfContent = '...'; // Your OSF content
const document = parse(osfContent);

// Convert to PDF
const pdfConverter = new PDFConverter();
const pdfResult = await pdfConverter.convert(document, {
  theme: 'corporate',
  pageSize: 'A4',
  orientation: 'portrait'
});

// Save the result
import { writeFileSync } from 'fs';
writeFileSync('output.pdf', pdfResult.buffer);
```

### Advanced Options

```typescript
const options = {
  theme: 'corporate',              // Theme: default, corporate, academic, modern
  pageSize: 'A4',                 // Page size: A4, letter, legal
  orientation: 'portrait',        // Orientation: portrait, landscape
  includeMetadata: true,          // Include document metadata
  margins: {                      // Custom margins (PDF only)
    top: 1,
    right: 1,
    bottom: 1,
    left: 1
  }
};
```

### CLI Integration

The converters are integrated into the OmniScript CLI:

```bash
# Generate PDF with corporate theme
osf render document.osf --format pdf --output report.pdf --theme corporate

# Create PowerPoint presentation
osf render slides.osf --format pptx --output presentation.pptx --theme modern

# Generate Excel workbook
osf render data.osf --format xlsx --output spreadsheet.xlsx
```

## Supported Formats

| Format | Extension | Features |
|--------|-----------|----------|
| PDF    | .pdf      | Themes, custom styling, vector graphics |
| DOCX   | .docx     | Rich text, tables, headers, formatting |
| PPTX   | .pptx     | Automatic slides, themes, bullet points |
| XLSX   | .xlsx     | Data tables, formulas, multiple sheets |

## Themes

### Default Theme
Clean, professional styling suitable for most documents.

### Corporate Theme
Business-focused design with blue accents and formal typography.

### Academic Theme
Traditional academic styling with serif fonts and conservative layout.

### Modern Theme
Contemporary design with vibrant colors and modern typography.

## API Reference

### PDF Converter
- High-quality PDF generation using Puppeteer
- Custom CSS styling and themes
- Vector graphics support
- Print-optimized layouts

### DOCX Converter
- Native Microsoft Word format
- Rich text formatting (bold, italic, headings)
- Tables with styling
- Document metadata

### PPTX Converter
- PowerPoint presentation generation
- Automatic slide creation from content blocks
- Theme-based styling
- Bullet point and text layouts

### XLSX Converter
- Excel workbook creation
- Multiple worksheet support
- Formula evaluation
- Data type preservation

## Development

```bash
# Install dependencies
pnpm install

# Build the package
pnpm run build

# Run tests
pnpm test

# Run test conversion
npx tsx src/test.ts
```

## License

MIT © 2025 Alphin Tom
