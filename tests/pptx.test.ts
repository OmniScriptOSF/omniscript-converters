import { describe, it, expect, beforeEach } from 'vitest';
import { PPTXConverter } from '../src/pptx';
import { parse } from 'omniscript-parser';

describe('PPTXConverter', () => {
  let converter: PPTXConverter;

  beforeEach(() => {
    converter = new PPTXConverter();
  });

  describe('getSupportedFormats', () => {
    it('should return pptx as supported format', () => {
      const formats = converter.getSupportedFormats();
      expect(formats).toContain('pptx');
      expect(formats.length).toBe(1);
    });
  });

  describe('convert', () => {
    it('should convert simple slides to PPTX', async () => {
      const osfContent = `
@meta {
  title: "Test Presentation";
  author: "Test Author";
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

      expect(result).toBeDefined();
      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.mimeType).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation');
      expect(result.extension).toBe('pptx');
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should create title slide from metadata', async () => {
      const osfContent = `
@meta {
  title: "My Presentation";
  author: "John Doe";
  date: "2025-01-15";
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { includeMetadata: true });

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle TitleOnly layout', async () => {
      const osfContent = `
@slide {
  title: "Title Only Slide";
  layout: "TitleOnly";
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle TitleAndContent layout', async () => {
      const osfContent = `
@slide {
  title: "Content Slide";
  layout: "TitleAndContent";
  bullets {
    "Item 1";
    "Item 2";
    "Item 3";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle TwoColumn layout', async () => {
      const osfContent = `
@slide {
  title: "Two Column Slide";
  layout: "TwoColumn";
  bullets {
    "Left 1";
    "Left 2";
    "Right 1";
    "Right 2";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle Blank layout', async () => {
      const osfContent = `
@slide {
  layout: "Blank";
  bullets {
    "Content";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should apply transition effects', async () => {
      const osfContent = `
@slide {
  title: "Animated Slide";
  transition: "FadeIn";
  bullets {
    "Point 1";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should include slide notes', async () => {
      const osfContent = `
@slide {
  title: "Slide with Notes";
  notes: "These are speaker notes";
  bullets {
    "Visible point";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle inline formatting', async () => {
      const osfContent = `
@slide {
  title: "Formatted Slide";
  bullets {
    "This is **bold**";
    "This is *italic*";
    "This is \`code\`";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should apply themes', async () => {
      const osfContent = `
@meta {
  title: "Themed Presentation";
}

@slide {
  title: "Corporate Theme";
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { theme: 'corporate' });

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should convert doc blocks to slides', async () => {
      const osfContent = `
@doc {
  # Section 1
  This is content for section 1.

  ## Subsection
  This is a subsection.
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should convert sheet blocks to table slides', async () => {
      const osfContent = `
@sheet {
  name: "Data Table";
  cols: [Product, Sales];
  data {
    (2,1)="A"; (2,2)=100;
    (3,1)="B"; (3,2)=200;
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should handle multiple slides with different layouts', async () => {
      const osfContent = `
@meta {
  title: "Multi-Layout Presentation";
}

@slide {
  title: "Title Slide";
  layout: "TitleOnly";
}

@slide {
  title: "Content Slide";
  layout: "TitleAndContent";
  bullets {
    "Point 1";
    "Point 2";
  }
}

@slide {
  title: "Two Column";
  layout: "TwoColumn";
  bullets {
    "Left";
    "Right";
  }
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document);

      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.buffer.length).toBeGreaterThan(0);
    });

    it('should apply modern theme', async () => {
      const osfContent = `
@slide {
  title: "Modern Theme";
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { theme: 'modern' });

      expect(result.buffer).toBeInstanceOf(Buffer);
    });

    it('should apply academic theme', async () => {
      const osfContent = `
@slide {
  title: "Academic Theme";
}
      `;

      const document = parse(osfContent);
      const result = await converter.convert(document, { theme: 'academic' });

      expect(result.buffer).toBeInstanceOf(Buffer);
    });
  });
});
