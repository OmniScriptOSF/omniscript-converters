import { describe, it, expect, beforeEach } from 'vitest';
import { XLSXConverter } from '../src/xlsx';
import { parse } from 'omniscript-parser';

describe('XLSXConverter', () => {
  let converter: XLSXConverter;

  beforeEach(() => {
    converter = new XLSXConverter();
  });

  describe('getSupportedFormats', () => {
    it('should return xlsx as supported format', () => {
      const formats = converter.getSupportedFormats();
      expect(formats).toContain('xlsx');
      expect(formats.length).toBe(1);
    });
  });

  describe('convert', () => {
    it('should convert simple sheet to XLSX', async () => {
      const osfContent = `
@sheet {
  name: "Sales";
  cols: [Product, Price, Quantity];
  data {
    (2,1)="Widget"; (2,2)=19.99; (2,3)=100;
    (3,1)="Gadget"; (3,2)=29.99; (3,3)=150;
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result).toBeDefined();
      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.mimeType).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      expect(result.extension).toBe('xlsx');
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should handle formulas', async () => {
      const osfContent = `
@sheet {
  name: "Calculations";
  cols: [A, B, Sum];
  data {
    (2,1)=10; (2,2)=20;
    (3,1)=30; (3,2)=40;
  }
  formula (2,3): "=A2+B2";
  formula (3,3): "=A3+B3";
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { formulaEvaluation: true });

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should handle multiple sheets', async () => {
      const osfContent = `
@sheet {
  name: "Sheet1";
  cols: [Col1, Col2];
  data {
    (2,1)="A"; (2,2)="B";
  }
}

@sheet {
  name: "Sheet2";
  cols: [Col1, Col2];
  data {
    (2,1)="C"; (2,2)="D";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should create summary sheet for doc blocks', async () => {
      const osfContent = `
@meta {
  title: "Report";
  author: "John Doe";
}

@doc {
  # Introduction
  This is a document.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle slide blocks as content worksheets', async () => {
      const osfContent = `
@slide {
  title: "Slide 1";
  bullets {
    "Point 1";
    "Point 2";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle numeric data types', async () => {
      const osfContent = `
@sheet {
  name: "Numbers";
  cols: [Integer, Float, Negative];
  data {
    (2,1)=100; (2,2)=19.99; (2,3)=-5;
    (3,1)=200; (3,2)=29.99; (3,3)=-10;
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should apply worksheet styling', async () => {
      const osfContent = `
@sheet {
  name: "Styled Sheet";
  cols: [Name, Value];
  data {
    (2,1)="Item"; (2,2)=100;
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { theme: 'corporate' });

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle empty sheet', async () => {
      const osfContent = `
@sheet {
  name: "Empty";
  cols: [A, B, C];
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle complex formulas', async () => {
      const osfContent = `
@sheet {
  name: "Complex";
  cols: [Q1, Q2, Total, Average, Growth];
  data {
    (2,1)=1000; (2,2)=1200;
  }
  formula (2,3): "=A2+B2";
  formula (2,4): "=(A2+B2)/2";
  formula (2,5): "=(B2-A2)/A2*100";
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should set workbook metadata', async () => {
      const osfContent = `
@meta {
  title: "Annual Report 2025";
  author: "Finance Team";
  date: "2025-01-15";
}

@sheet {
  name: "Data";
  cols: [Month, Revenue];
  data {
    (2,1)="Jan"; (2,2)=10000;
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });
  });
});
