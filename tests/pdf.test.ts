import { describe, it, expect, beforeEach } from 'vitest';
import { PDFConverter } from '../src/pdf';
import { parse } from 'omniscript-parser';

describe('PDFConverter', () => {
  let converter: PDFConverter;

  beforeEach(() => {
    converter = new PDFConverter();
  });

  describe('getSupportedFormats', () => {
    it('should return pdf as supported format', () => {
      const formats = converter.getSupportedFormats();
      expect(formats).toContain('pdf');
      expect(formats.length).toBe(1);
    });
  });

  describe('convert', () => {
    it('should convert simple OSF document to PDF', async () => {
      const osfContent = `
@meta {
  title: "Test Document";
  author: "Test Author";
}

@doc {
  # Hello World
  This is a test document.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result).toBeDefined();
      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.mimeType).toBe('application/pdf');
      expect(result.extension).toBe('pdf');
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should convert document with slides', async () => {
      const osfContent = `
@meta {
  title: "Presentation";
}

@slide {
  title: "First Slide";
  bullets {
    "Point 1";
    "Point 2";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should convert document with sheets', async () => {
      const osfContent = `
@sheet {
  name: "Data";
  cols: [Name, Value];
  data {
    (2,1)="Item1"; (2,2)=100;
    (3,1)="Item2"; (3,2)=200;
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should apply corporate theme', async () => {
      const osfContent = `
@meta {
  title: "Corporate Doc";
  theme: "Corporate";
}

@doc {
  # Corporate Theme Test
  Testing theme support.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { theme: 'corporate' });

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should apply academic theme', async () => {
      const osfContent = `
@meta {
  title: "Academic Paper";
}

@doc {
  # Introduction
  This is an academic paper.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { theme: 'academic' });

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should apply modern theme', async () => {
      const osfContent = `
@meta {
  title: "Modern Doc";
}

@doc {
  # Modern Theme
  Testing modern styling.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { theme: 'modern' });

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle empty document', async () => {
      const osfContent = `@meta { title: "Empty"; }`;
      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should handle markdown formatting in doc blocks', async () => {
      const osfContent = `
@doc {
  # Heading 1
  ## Heading 2
  ### Heading 3

  This is **bold** and this is *italic*.

  This is \`code\`.

  - List item 1
  - List item 2
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should handle multiple blocks of different types', async () => {
      const osfContent = `
@meta {
  title: "Multi-Block Document";
  author: "Test";
}

@doc {
  # Introduction
  This document has multiple blocks.
}

@slide {
  title: "Key Points";
  bullets {
    "Point A";
    "Point B";
  }
}

@sheet {
  name: "Data";
  cols: [Item, Count];
  data {
    (2,1)="A"; (2,2)=10;
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
