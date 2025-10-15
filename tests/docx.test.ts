import { describe, it, expect, beforeEach } from 'vitest';
import { DOCXConverter } from '../src/docx';
import { parse } from 'omniscript-parser';

describe('DOCXConverter', () => {
  let converter: DOCXConverter;

  beforeEach(() => {
    converter = new DOCXConverter();
  });

  describe('getSupportedFormats', () => {
    it('should return docx as supported format', () => {
      const formats = converter.getSupportedFormats();
      expect(formats).toContain('docx');
      expect(formats.length).toBe(1);
    });
  });

  describe('convert', () => {
    it('should convert simple OSF document to DOCX', async () => {
      const osfContent = `
@meta {
  title: "Test Document";
  author: "Test Author";
  date: "2025-01-15";
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
      expect(result.mimeType).toBe('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      expect(result.extension).toBe('docx');
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should include metadata when requested', async () => {
      const osfContent = `
@meta {
  title: "Document with Metadata";
  author: "John Doe";
  date: "2025-01-15";
}

@doc {
  Content here.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { includeMetadata: true });

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should convert document with headings', async () => {
      const osfContent = `
@doc {
  # Heading 1
  Content under heading 1.

  ## Heading 2
  Content under heading 2.

  ### Heading 3
  Content under heading 3.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should convert document with lists', async () => {
      const osfContent = `
@doc {
  # Shopping List

  - Apples
  - Bananas
  - Oranges

  * Item 1
  * Item 2
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle inline formatting', async () => {
      const osfContent = `
@doc {
  This is **bold text**.
  This is *italic text*.
  This is \`code text\`.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should convert slides to sections', async () => {
      const osfContent = `
@slide {
  title: "Slide 1";
  bullets {
    "Point 1";
    "Point 2";
  }
}

@slide {
  title: "Slide 2";
  bullets {
    "Point A";
    "Point B";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should convert sheets to tables', async () => {
      const osfContent = `
@sheet {
  name: "Sales Data";
  cols: [Product, Revenue, Growth];
  data {
    (2,1)="Widget"; (2,2)=1000; (2,3)=10;
    (3,1)="Gadget"; (3,2)=2000; (3,3)=15;
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should handle empty document', async () => {
      const osfContent = `@meta { title: "Empty"; }`;
      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });
  });
});
