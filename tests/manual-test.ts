import { readFileSync, writeFileSync } from 'fs';
import { join } from 'path';
import { parse } from 'omniscript-parser';
import { PDFConverter, DOCXConverter, PPTXConverter, XLSXConverter } from './index';

async function testConverters() {
  // Load a sample OSF file
  const sampleOSF = `
@meta {
  title: "Test Document";
  author: "Test Author";
  date: "2025-01-01";
  theme: "corporate";
}

@doc {
  # Introduction
  This is a **test document** with *various* formatting elements.
  
  ## Key Features
  - Rich text support
  - Multiple output formats
  - Formula calculations
}

@slide {
  title: "Sample Slide";
  bullets {
    "First bullet point";
    "Second bullet point";
    "Third bullet point";
  }
}

@sheet {
  name: "Sales Data";
  cols: [Region, Q1, Q2, Growth];
  data {
    (1,1)="North"; (1,2)=100; (1,3)=120;
    (2,1)="South"; (2,2)=150; (2,3)=180;
  }
  formula (1,4): "=(C1-B1)/B1*100";
  formula (2,4): "=(C2-B2)/B2*100";
}
`;

  console.log('Parsing OSF document...');
  const document = parse(sampleOSF);
  
  const options = {
    theme: 'corporate',
    includeMetadata: true,
    pageSize: 'A4' as const,
    orientation: 'portrait' as const
  };

  try {
    // Test PDF Converter
    console.log('Testing PDF converter...');
    const pdfConverter = new PDFConverter();
    const pdfResult = await pdfConverter.convert(document, options);
    writeFileSync('./test-output.pdf', pdfResult.buffer);
    console.log('✅ PDF conversion successful');

    // Test DOCX Converter
    console.log('Testing DOCX converter...');
    const docxConverter = new DOCXConverter();
    const docxResult = await docxConverter.convert(document, options);
    writeFileSync('./test-output.docx', docxResult.buffer);
    console.log('✅ DOCX conversion successful');

    // Test PPTX Converter
    console.log('Testing PPTX converter...');
    const pptxConverter = new PPTXConverter();
    const pptxResult = await pptxConverter.convert(document, options);
    writeFileSync('./test-output.pptx', pptxResult.buffer);
    console.log('✅ PPTX conversion successful');

    // Test XLSX Converter
    console.log('Testing XLSX converter...');
    const xlsxConverter = new XLSXConverter();
    const xlsxResult = await xlsxConverter.convert(document, options);
    writeFileSync('./test-output.xlsx', xlsxResult.buffer);
    console.log('✅ XLSX conversion successful');

    console.log('\n🎉 All converters tested successfully!');
    console.log('Output files generated:');
    console.log('- test-output.pdf');
    console.log('- test-output.docx');
    console.log('- test-output.pptx');
    console.log('- test-output.xlsx');

  } catch (error) {
    console.error('❌ Test failed:', error);
    process.exit(1);
  }
}

// Run tests if this file is executed directly
if (require.main === module) {
  testConverters().catch(console.error);
}

export { testConverters };