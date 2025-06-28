# OmniScript Converters

<div align="center">

<img src="https://raw.githubusercontent.com/OmniScriptOSF/omniscript-core/main/assets/osf-icon-512px.png" alt="OmniScript Logo" width="120" height="120" />

# 🔄 Professional Format Converters

**Convert OmniScript Format (OSF) documents to PDF, DOCX, PPTX, and XLSX with enterprise-grade quality**

[![npm version](https://badge.fury.io/js/omniscript-converters.svg)](https://badge.fury.io/js/omniscript-converters)
[![npm downloads](https://img.shields.io/npm/dm/omniscript-converters.svg)](https://www.npmjs.com/package/omniscript-converters)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![TypeScript](https://img.shields.io/badge/TypeScript-007ACC?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)

[🚀 Quick Start](#-quick-start) • [📖 API Reference](#-api-reference) • [🎨 Themes](#-themes) • [💡 Examples](#-examples) • [🛠️ CLI Integration](#-cli-integration)

</div>

---

## ✨ Features

<table>
<tr>
<td width="25%">

### 📄 **PDF Generation**
- High-quality PDF documents
- Custom themes & styling
- Vector graphics support
- Print-optimized layouts

</td>
<td width="25%">

### 📝 **DOCX Export**
- Native Microsoft Word format
- Rich text formatting
- Tables with styling
- Document metadata

</td>
<td width="25%">

### 🎯 **PPTX Creation**
- PowerPoint presentations
- Automatic slide generation
- Theme-based styling
- Interactive layouts

</td>
<td width="25%">

### 📊 **XLSX Conversion**
- Excel workbooks
- Formula evaluation
- Multiple worksheets
- Data type preservation

</td>
</tr>
</table>

---

## 🚀 Quick Start

### 📦 Installation

```bash
npm install omniscript-converters
# or
pnpm add omniscript-converters
# or
yarn add omniscript-converters
```

### ⚡ Basic Usage

```typescript
import { parse } from 'omniscript-parser';
import { 
  PDFConverter, 
  DOCXConverter, 
  PPTXConverter, 
  XLSXConverter 
} from 'omniscript-converters';

// Parse your OSF document
const osfContent = `
@meta {
  title: "Business Report";
  author: "Jane Smith";
  theme: "Corporate";
}

@doc {
  # Executive Summary
  Our Q2 performance exceeded expectations with **15% revenue growth**.
}
`;

const document = parse(osfContent);

// Convert to PDF
const pdfConverter = new PDFConverter();
const pdfResult = await pdfConverter.convert(document, {
  theme: 'corporate',
  pageSize: 'A4'
});

// Save the result
import { writeFileSync } from 'fs';
writeFileSync('report.pdf', pdfResult.buffer);
```

---

## 📖 API Reference

### 🔧 Converter Classes

#### PDFConverter

```typescript
const pdfConverter = new PDFConverter();
const result = await pdfConverter.convert(document, options);
```

**Options:**
- `theme`: `'default' | 'corporate' | 'academic' | 'modern'`
- `pageSize`: `'A4' | 'letter' | 'legal'`
- `orientation`: `'portrait' | 'landscape'`
- `margins`: `{ top: number, right: number, bottom: number, left: number }`

#### DOCXConverter

```typescript
const docxConverter = new DOCXConverter();
const result = await docxConverter.convert(document, options);
```

**Options:**
- `theme`: Theme selection for styling
- `includeMetadata`: Include document metadata
- `pageSize`: Page size configuration

#### PPTXConverter

```typescript
const pptxConverter = new PPTXConverter();
const result = await pptxConverter.convert(document, options);
```

**Options:**
- `theme`: Presentation theme
- `slideSize`: Slide dimensions
- `autoSlides`: Automatic slide generation

#### XLSXConverter

```typescript
const xlsxConverter = new XLSXConverter();
const result = await xlsxConverter.convert(document, options);
```

**Options:**
- `sheetNames`: Custom sheet naming
- `formulaEvaluation`: Enable/disable formula evaluation
- `formatting`: Apply cell formatting

---

## 🎨 Themes

### 🏢 Corporate Theme
Professional business styling with blue accents and formal typography.
```typescript
{ theme: 'corporate' }
```

### 🎓 Academic Theme
Traditional academic styling with serif fonts and conservative layout.
```typescript
{ theme: 'academic' }
```

### ✨ Modern Theme
Contemporary design with vibrant colors and modern typography.
```typescript
{ theme: 'modern' }
```

### 📄 Default Theme
Clean, versatile styling suitable for most documents.
```typescript
{ theme: 'default' }
```

---

## 💡 Examples

### 📊 Business Dashboard

```typescript
import { parse } from 'omniscript-parser';
import { PDFConverter, XLSXConverter } from 'omniscript-converters';

const dashboardOSF = `
@meta {
  title: "Q2 Sales Dashboard";
  author: "Analytics Team";
  theme: "Corporate";
}

@sheet {
  name: "Regional Performance";
  cols: [Region, Q1_Sales, Q2_Sales, Growth];
  data {
    (1,1)="North America"; (1,2)=850000; (1,3)=975000;
    (2,1)="Europe"; (2,2)=650000; (2,3)=748000;
    (3,1)="Asia Pacific"; (3,2)=400000; (3,3)=477000;
  }
  formula (1,4): "=(C1-B1)/B1*100";
  formula (2,4): "=(C2-B2)/B2*100";
  formula (3,4): "=(C3-B3)/B3*100";
}

@doc {
  # Sales Analysis Summary
  
  Our Q2 results show strong performance across all regions:
  
  - **North America**: 14.7% growth
  - **Europe**: 15.1% growth  
  - **Asia Pacific**: 19.3% growth
}
`;

const document = parse(dashboardOSF);

// Generate PDF report
const pdfConverter = new PDFConverter();
const pdfReport = await pdfConverter.convert(document, {
  theme: 'corporate',
  pageSize: 'A4'
});

// Generate Excel workbook
const xlsxConverter = new XLSXConverter();
const xlsxReport = await xlsxConverter.convert(document, {
  formulaEvaluation: true
});

// Save both formats
writeFileSync('dashboard.pdf', pdfReport.buffer);
writeFileSync('dashboard.xlsx', xlsxReport.buffer);
```

### 🎯 Presentation Creation

```typescript
const presentationOSF = `
@meta {
  title: "Product Launch 2025";
  author: "Product Team";
  theme: "Modern";
}

@slide {
  title: "Introducing OmniScript";
  layout: "TitleAndContent";
  content: "The future of document processing is here.";
}

@slide {
  title: "Key Features";
  layout: "TitleAndBullets";
  bullets {
    "🚀 Universal document format";
    "🤖 AI-native syntax design";
    "🔄 Git-friendly version control";
    "📊 Multi-format export capabilities";
  }
}
`;

const document = parse(presentationOSF);
const pptxConverter = new PPTXConverter();
const presentation = await pptxConverter.convert(document, {
  theme: 'modern',
  slideSize: 'widescreen'
});

writeFileSync('product-launch.pptx', presentation.buffer);
```

---

## 🛠️ CLI Integration

The converters are seamlessly integrated with the OmniScript CLI:

```bash
# Generate PDF with corporate theme
osf render document.osf --format pdf --output report.pdf --theme corporate

# Create PowerPoint presentation
osf render slides.osf --format pptx --output presentation.pptx --theme modern

# Generate Excel workbook with formulas
osf render data.osf --format xlsx --output spreadsheet.xlsx

# Convert to Word document
osf render document.osf --format docx --output document.docx --theme academic
```

### 📋 CLI Options

| Option | Description | Values |
|--------|-------------|--------|
| `--format` | Output format | `pdf`, `docx`, `pptx`, `xlsx` |
| `--theme` | Visual theme | `default`, `corporate`, `academic`, `modern` |
| `--output` | Output file path | Any valid file path |
| `--page-size` | Page size (PDF/DOCX) | `A4`, `letter`, `legal` |
| `--orientation` | Page orientation | `portrait`, `landscape` |

---

## 🔧 Advanced Configuration

### Custom Styling

```typescript
const customOptions = {
  theme: 'corporate',
  pageSize: 'A4',
  orientation: 'portrait',
  margins: {
    top: 1,      // inches
    right: 1,
    bottom: 1,
    left: 1
  },
  fonts: {
    heading: 'Arial Bold',
    body: 'Arial',
    monospace: 'Courier New'
  },
  colors: {
    primary: '#2563eb',
    secondary: '#64748b',
    accent: '#f59e0b'
  }
};

const result = await pdfConverter.convert(document, customOptions);
```

### Batch Processing

```typescript
import { glob } from 'glob';

// Convert all OSF files in a directory
const osfFiles = await glob('documents/*.osf');
const pdfConverter = new PDFConverter();

for (const file of osfFiles) {
  const content = readFileSync(file, 'utf-8');
  const document = parse(content);
  const result = await pdfConverter.convert(document, { theme: 'corporate' });
  
  const outputPath = file.replace('.osf', '.pdf');
  writeFileSync(outputPath, result.buffer);
  console.log(`✅ Converted ${file} → ${outputPath}`);
}
```

---

## 📊 Supported Features

<div align="center">

| Feature | PDF | DOCX | PPTX | XLSX |
|---------|-----|------|------|------|
| **Text Formatting** | ✅ | ✅ | ✅ | ✅ |
| **Tables** | ✅ | ✅ | ✅ | ✅ |
| **Formulas** | ❌ | ❌ | ❌ | ✅ |
| **Images** | ✅ | ✅ | ✅ | ✅ |
| **Themes** | ✅ | ✅ | ✅ | ⚠️ |
| **Metadata** | ✅ | ✅ | ✅ | ✅ |
| **Headers/Footers** | ✅ | ✅ | ❌ | ❌ |
| **Page Breaks** | ✅ | ✅ | N/A | N/A |

</div>

**Legend:** ✅ Full Support • ⚠️ Partial Support • ❌ Not Applicable

---

## 🛡️ Error Handling

```typescript
try {
  const result = await pdfConverter.convert(document, options);
  console.log('✅ Conversion successful');
} catch (error) {
  if (error instanceof ConversionError) {
    console.error('❌ Conversion failed:', error.message);
    console.error('Details:', error.details);
  } else {
    console.error('❌ Unexpected error:', error);
  }
}
```

### Common Error Types

- `ParseError`: Invalid OSF syntax
- `ConversionError`: Format conversion issues
- `ThemeError`: Invalid theme configuration
- `FileSystemError`: File I/O problems

---

## 🔧 Development

### Setup

```bash
# Clone the repository
git clone https://github.com/OmniScriptOSF/omniscript-converters.git
cd omniscript-converters

# Install dependencies
pnpm install

# Build the package
pnpm run build

# Run tests
pnpm test

# Run test conversion
pnpm run test:convert
```

### Testing

```bash
# Run all tests
pnpm test

# Run specific converter tests
pnpm test pdf
pnpm test docx
pnpm test pptx
pnpm test xlsx

# Generate test reports
pnpm run test:coverage
```

---

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guide](https://github.com/OmniScriptOSF/omniscript-core/blob/main/CONTRIBUTING.md).

### 🌟 Areas for Contribution

- 🎨 **New Themes** - Design beautiful new visual themes
- 📊 **Chart Support** - Add chart rendering capabilities
- 🔧 **Performance** - Optimize conversion speed and memory usage
- 🧪 **Testing** - Expand test coverage and edge cases
- 📖 **Documentation** - Improve examples and guides

---

## 📄 License

MIT License © 2025 [Alphin Tom](https://github.com/alpha912)

---

## 🔗 Related Packages

- **[omniscript-parser](https://www.npmjs.com/package/omniscript-parser)** - Core OSF parsing library
- **[omniscript-cli](https://www.npmjs.com/package/omniscript-cli)** - Command-line tools
- **[OmniScript Core](https://github.com/OmniScriptOSF/omniscript-core)** - Complete ecosystem

---

## 📞 Support

- 🐛 [Report Issues](https://github.com/OmniScriptOSF/omniscript-converters/issues)
- 💬 [Discussions](https://github.com/OmniScriptOSF/omniscript-core/discussions)
- 🏢 [Organization](https://github.com/OmniScriptOSF)
- 👤 [Maintainer](https://github.com/alpha912)

---

<div align="center">

### 🚀 Ready to convert your OSF documents?

**[📦 Install Now](https://www.npmjs.com/package/omniscript-converters)** • **[📖 View Examples](https://github.com/OmniScriptOSF/omniscript-examples)** • **[🤝 Get Support](https://github.com/OmniScriptOSF/omniscript-core/discussions)**

---

*Built with ❤️ for professional document workflows*

</div>