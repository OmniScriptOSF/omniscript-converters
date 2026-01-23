// File: src/types/xlsx-populate.d.ts
// What: Minimal TypeScript typings for xlsx-populate used by converters and tests.
// Why: Provide strict typing without relying on external @types packages.
// Related: src/xlsx.ts, tests/xlsx.test.ts

declare module 'xlsx-populate' {
  export type Color = {
    rgb?: string;
    theme?: number;
    tint?: number;
    indexed?: number;
    auto?: boolean;
  };

  export type ColorInput = Color | string | number;

  export type BorderStyle =
    | 'hair'
    | 'dotted'
    | 'dashDotDot'
    | 'dashed'
    | 'mediumDashDotDot'
    | 'thin'
    | 'slantDashDot'
    | 'mediumDashDot'
    | 'mediumDashed'
    | 'medium'
    | 'thick'
    | 'double';

  export type BorderInput =
    | BorderStyle
    | boolean
    | {
        style?: BorderStyle;
        color?: ColorInput;
        direction?: 'up' | 'down' | 'both';
      };

  export type Borders = {
    left?: BorderInput;
    right?: BorderInput;
    top?: BorderInput;
    bottom?: BorderInput;
    diagonal?: BorderInput;
  };

  export type Fill =
    | string
    | {
        type?: 'pattern' | 'gradient';
        pattern?: string;
        color?: ColorInput;
        foreground?: ColorInput;
        background?: ColorInput;
      };

  export type HorizontalAlignment =
    | 'left'
    | 'center'
    | 'right'
    | 'fill'
    | 'justify'
    | 'centerContinuous'
    | 'distributed';

  export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

  export interface Styles {
    bold?: boolean;
    italic?: boolean;
    fontSize?: number;
    fontFamily?: string;
    fontColor?: ColorInput;
    fill?: Fill;
    numberFormat?: string;
    border?: Borders | BorderInput;
    leftBorder?: BorderInput;
    rightBorder?: BorderInput;
    topBorder?: BorderInput;
    bottomBorder?: BorderInput;
    horizontalAlignment?: HorizontalAlignment;
    verticalAlignment?: VerticalAlignment;
    wrapText?: boolean;
  }

  export type CellValue = string | number | boolean | Date | null | undefined;

  export interface Cell {
    value(): CellValue;
    value(
      value:
        | CellValue
        | CellValue[][]
        | ((cell: Cell, ri: number, ci: number, range: Range) => CellValue)
    ): Cell | Range;
    style(name: string): unknown;
    style(names: string[]): Record<string, unknown>;
    style(name: string, value: unknown): Cell | Range;
    style(styles: Styles): Cell | Range;
    formula(): string | null;
    formula(formula: string): Cell;
    row(): Row;
    column(): Column;
  }

  export interface Range {
    value(): CellValue[][];
    value(
      values:
        | CellValue[][]
        | CellValue
        | ((cell: Cell, ri: number, ci: number, range: Range) => CellValue)
    ): Range;
    style(name: string, value: unknown): Range;
    style(styles: Styles): Range;
    merged(): boolean;
    merged(merged: boolean): Range;
    address(): string;
  }

  export interface Row {
    height(): number | undefined;
    height(value: number): Row;
    cell(column: number | string): Cell;
    style(name: string, value: unknown): Row;
    style(styles: Styles): Row;
  }

  export interface Column {
    width(): number | undefined;
    width(value: number): Column;
    style(name: string, value: unknown): Column;
    style(styles: Styles): Column;
  }

  export interface Sheet {
    name(): string;
    name(name: string): Sheet;
    cell(row: number, column: number): Cell;
    cell(address: string): Cell;
    range(address: string): Range;
    range(startRow: number, startCol: number, endRow: number, endCol: number): Range;
    row(rowNumber: number): Row;
    column(column: number | string): Column;
    usedRange(): Range | undefined;
    tabColor(): Color | undefined;
    tabColor(color: ColorInput): Color | undefined;
  }

  export interface Workbook {
    sheet(indexOrName: number | string): Sheet;
    sheets(): Sheet[];
    addSheet(name: string, indexOrBeforeSheet?: number | string | Sheet): Sheet;
    deleteSheet(sheet: number | string | Sheet): Workbook;
    property(name: string): unknown;
    property(names: string[]): Record<string, unknown>;
    property(name: string, value: unknown): Workbook;
    property(properties: Record<string, unknown>): Workbook;
    properties(): unknown;
    outputAsync(options?: {
      type?: 'nodebuffer' | 'base64' | 'binarystring' | 'uint8array' | 'arraybuffer';
      password?: string;
    }): Promise<Buffer | string | Uint8Array | ArrayBuffer>;
  }

  export interface XlsxPopulateStatic {
    fromBlankAsync(): Promise<Workbook>;
    fromFileAsync(
      path: string,
      opts?: { password?: string }
    ): Promise<Workbook>;
    fromDataAsync(
      data: Buffer | ArrayBuffer | Uint8Array | string,
      opts?: { password?: string }
    ): Promise<Workbook>;
  }

  const XlsxPopulate: XlsxPopulateStatic;
  export default XlsxPopulate;
}
