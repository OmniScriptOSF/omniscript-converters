// File: tests/v1-features.test.ts
// What: Tests for v1.0 block rendering (@chart, @diagram, @code)
// Why: Ensure converters properly handle v1.0 features
// Related: pdf.ts, docx.ts, pptx.ts

import { describe, it, expect } from 'vitest';
import { PDFConverter } from '../src/pdf';
import { DOCXConverter } from '../src/docx';
import { PPTXConverter } from '../src/pptx';
import type { OSFDocument } from 'omniscript-parser';

describe('v1.0 Features - Chart Blocks', () => {
  const chartDocument: OSFDocument = {
    blocks: [
      {
        type: 'chart',
        chartType: 'bar',
        title: 'Sales Data',
        data: [
          { label: 'Q1', values: [100] },
          { label: 'Q2', values: [150] },
          { label: 'Q3', values: [200] }
        ],
        options: {
          xAxis: 'Quarter',
          yAxis: 'Revenue ($K)',
          legend: true
        }
      }
    ]
  };

  it('PDF should render chart block', async () => {
    const converter = new PDFConverter();
    const result = await converter.convert(chartDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(1000);
    expect(result.mimeType).toBe('application/pdf');
  });

  it('DOCX should render chart block', async () => {
    const converter = new DOCXConverter();
    const result = await converter.convert(chartDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(500);
    expect(result.mimeType).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  });

  it('PPTX should render chart block', async () => {
    const converter = new PPTXConverter();
    const result = await converter.convert(chartDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(1000);
    expect(result.mimeType).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation');
  });

  it('PDF should handle pie chart', async () => {
    const pieDocument: OSFDocument = {
      blocks: [
        {
          type: 'chart',
          chartType: 'pie',
          title: 'Market Share',
          data: [
            { label: 'Product A', values: [30] },
            { label: 'Product B', values: [45] },
            { label: 'Product C', values: [25] }
          ]
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(pieDocument);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('PDF should handle line chart', async () => {
    const lineDocument: OSFDocument = {
      blocks: [
        {
          type: 'chart',
          chartType: 'line',
          title: 'Growth Trend',
          data: [
            { label: 'Series 1', values: [10, 20, 30, 40, 50] }
          ]
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(lineDocument);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('should handle chart with custom colors', async () => {
    const coloredDocument: OSFDocument = {
      blocks: [
        {
          type: 'chart',
          chartType: 'bar',
          title: 'Colored Chart',
          data: [
            { label: 'A', values: [10] },
            { label: 'B', values: [20] }
          ],
          options: {
            colors: ['#FF0000', '#00FF00']
          }
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(coloredDocument);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });
});

describe('v1.0 Features - Diagram Blocks', () => {
  const flowchartDocument: OSFDocument = {
    blocks: [
      {
        type: 'diagram',
        diagramType: 'flowchart',
        engine: 'mermaid',
        title: 'Process Flow',
        code: 'graph TD\nA[Start] --> B[Process]\nB --> C[End]'
      }
    ]
  };

  it('PDF should render flowchart diagram', async () => {
    const converter = new PDFConverter();
    const result = await converter.convert(flowchartDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('DOCX should render diagram block', async () => {
    const converter = new DOCXConverter();
    const result = await converter.convert(flowchartDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(500);
  });

  it('PPTX should render diagram block', async () => {
    const converter = new PPTXConverter();
    const result = await converter.convert(flowchartDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('should handle sequence diagram', async () => {
    const seqDocument: OSFDocument = {
      blocks: [
        {
          type: 'diagram',
          diagramType: 'sequence',
          engine: 'mermaid',
          code: 'sequenceDiagram\nClient->>Server: Request\nServer-->>Client: Response'
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(seqDocument);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('should handle gantt diagram', async () => {
    const ganttDocument: OSFDocument = {
      blocks: [
        {
          type: 'diagram',
          diagramType: 'gantt',
          engine: 'mermaid',
          title: 'Project Timeline',
          code: 'gantt\ntitle Project Schedule\nsection Phase 1\nTask A: 2024-01-01, 30d'
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(ganttDocument);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });
});

describe('v1.0 Features - Code Blocks', () => {
  const codeDocument: OSFDocument = {
    blocks: [
      {
        type: 'osfcode',
        language: 'typescript',
        caption: 'Hello World',
        lineNumbers: false,
        code: 'function hello() {\n  console.log("Hello, World!");\n}'
      }
    ]
  };

  it('PDF should render code block', async () => {
    const converter = new PDFConverter();
    const result = await converter.convert(codeDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('DOCX should render code block', async () => {
    const converter = new DOCXConverter();
    const result = await converter.convert(codeDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(500);
  });

  it('PPTX should render code block', async () => {
    const converter = new PPTXConverter();
    const result = await converter.convert(codeDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('should handle code with line numbers', async () => {
    const numberedDocument: OSFDocument = {
      blocks: [
        {
          type: 'osfcode',
          language: 'python',
          lineNumbers: true,
          code: 'def main():\n    print("Hello")\n\nif __name__ == "__main__":\n    main()'
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(numberedDocument);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('should handle code with highlighting', async () => {
    const highlightDocument: OSFDocument = {
      blocks: [
        {
          type: 'osfcode',
          language: 'javascript',
          lineNumbers: true,
          highlight: [2, 3],
          code: 'function test() {\n  const x = 42;\n  return x * 2;\n}'
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(highlightDocument);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('should handle multiple programming languages', async () => {
    const languages = ['typescript', 'python', 'java', 'rust', 'go'];

    for (const lang of languages) {
      const doc: OSFDocument = {
        blocks: [
          {
            type: 'osfcode',
            language: lang,
            code: 'function example() { return 42; }'
          }
        ]
      };

      const converter = new PDFConverter();
      const result = await converter.convert(doc);
      
      expect(result.buffer.length).toBeGreaterThan(500);
    }
  });
});

describe('v1.0 Features - Mixed Documents', () => {
  const mixedDocument: OSFDocument = {
    blocks: [
      {
        type: 'meta',
        props: {
          title: 'v1.0 Feature Showcase',
          author: 'OmniScript Team',
          version: '1.0'
        }
      },
      {
        type: 'doc',
        content: '# Introduction\n\nThis document demonstrates v1.0 features.'
      },
      {
        type: 'chart',
        chartType: 'bar',
        title: 'Performance Metrics',
        data: [
          { label: 'Metric 1', values: [85] },
          { label: 'Metric 2', values: [92] }
        ]
      },
      {
        type: 'diagram',
        diagramType: 'flowchart',
        engine: 'mermaid',
        code: 'graph LR\nA --> B --> C'
      },
      {
        type: 'osfcode',
        language: 'typescript',
        caption: 'Example Code',
        lineNumbers: true,
        code: 'const x: number = 42;'
      }
    ]
  };

  it('PDF should render mixed v0.5 and v1.0 blocks', async () => {
    const converter = new PDFConverter();
    const result = await converter.convert(mixedDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(5000);
    expect(result.mimeType).toBe('application/pdf');
  });

  it('DOCX should render mixed v0.5 and v1.0 blocks', async () => {
    const converter = new DOCXConverter();
    const result = await converter.convert(mixedDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(2000);
  });

  it('PPTX should render mixed v0.5 and v1.0 blocks', async () => {
    const converter = new PPTXConverter();
    const result = await converter.convert(mixedDocument);
    
    expect(result).toBeDefined();
    expect(result.buffer).toBeInstanceOf(Buffer);
    expect(result.buffer.length).toBeGreaterThan(3000);
  });

  it('should handle document with all v1.0 block types', async () => {
    const allBlocksDocument: OSFDocument = {
      blocks: [
        {
          type: 'chart',
          chartType: 'line',
          title: 'Chart',
          data: [{ label: 'Data', values: [1, 2, 3] }]
        },
        {
          type: 'diagram',
          diagramType: 'flowchart',
          engine: 'mermaid',
          code: 'graph TD\nA --> B'
        },
        {
          type: 'osfcode',
          language: 'python',
          code: 'print("test")'
        }
      ]
    };

    const converters = [
      new PDFConverter(),
      new DOCXConverter(),
      new PPTXConverter()
    ];

    for (const converter of converters) {
      const result = await converter.convert(allBlocksDocument);
      expect(result.buffer.length).toBeGreaterThan(1000);
    }
  });
});

describe('v1.0 Features - Edge Cases', () => {
  it('should handle chart with empty data', async () => {
    const emptyChart: OSFDocument = {
      blocks: [
        {
          type: 'chart',
          chartType: 'bar',
          title: 'Empty Chart',
          data: []
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(emptyChart);
    
    expect(result.buffer.length).toBeGreaterThan(500);
  });

  it('should handle diagram with minimal code', async () => {
    const minimalDiagram: OSFDocument = {
      blocks: [
        {
          type: 'diagram',
          diagramType: 'flowchart',
          engine: 'mermaid',
          code: 'A'
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(minimalDiagram);
    
    expect(result.buffer.length).toBeGreaterThan(500);
  });

  it('should handle code block with empty code', async () => {
    const emptyCode: OSFDocument = {
      blocks: [
        {
          type: 'osfcode',
          language: 'text',
          code: ''
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(emptyCode);
    
    expect(result.buffer.length).toBeGreaterThan(500);
  });

  it('should handle chart without options', async () => {
    const noOptionsChart: OSFDocument = {
      blocks: [
        {
          type: 'chart',
          chartType: 'bar',
          title: 'Simple Chart',
          data: [{ label: 'A', values: [10] }]
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(noOptionsChart);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('should handle diagram without title', async () => {
    const noTitleDiagram: OSFDocument = {
      blocks: [
        {
          type: 'diagram',
          diagramType: 'sequence',
          engine: 'mermaid',
          code: 'sequenceDiagram\nA->>B: Message'
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(noTitleDiagram);
    
    expect(result.buffer.length).toBeGreaterThan(1000);
  });

  it('should handle code without caption', async () => {
    const noCaptionCode: OSFDocument = {
      blocks: [
        {
          type: 'osfcode',
          language: 'javascript',
          code: 'const x = 1;'
        }
      ]
    };

    const converter = new PDFConverter();
    const result = await converter.convert(noCaptionCode);
    
    expect(result.buffer.length).toBeGreaterThan(500);
  });
});
