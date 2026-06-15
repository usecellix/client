/* eslint-disable @typescript-eslint/no-explicit-any */
/* global Office, Excel */

declare namespace Office {
  interface HostInfo {
    host: HostType;
    platform: PlatformType;
  }

  enum HostType {
    Excel = 'Excel',
    Word = 'Word',
    PowerPoint = 'PowerPoint',
    Outlook = 'Outlook',
    OneNote = 'OneNote',
    Project = 'Project',
    Access = 'Access',
  }

  enum PlatformType {
    PC = 'PC',
    Mac = 'Mac',
    OfficeOnline = 'OfficeOnline',
    iOS = 'iOS',
    Android = 'Android',
  }

  function onReady(callback: (info: HostInfo) => void): void;
}

declare namespace Excel {
  interface RequestContext {
    workbook: Workbook;
    sync(): Promise<void>;
  }

  interface Workbook {
    worksheets: WorksheetCollection;
    names: NamedItemCollection;
    tables: TableCollection;
    getSelectedRange(): Range;
  }

  interface NamedItemCollection {
    items: NamedItem[];
    load(propertyNames?: string | string[]): NamedItemCollection;
    add(name: string, formula: string, comment?: string): NamedItem;
    getItem(name: string): NamedItem;
  }

  interface NamedItem {
    name: string;
    formula: string;
    type: string;
    comment?: string;
    load(propertyNames?: string | string[]): NamedItem;
  }

  interface TableCollection {
    items: Table[];
    load(propertyNames?: string | string[]): TableCollection;
  }

  interface Table {
    name: string;
    showHeaders: boolean;
    style: string;
    worksheet: Worksheet;
    getHeaderRowRange(): Range;
    getRange(): Range;
    load(propertyNames?: string | string[]): Table;
  }

  interface TableScopedCollection {
    add(address: string, hasHeaders: boolean): Table;
  }

  enum RangeCopyType {
    all = 'All',
    formulas = 'Formulas',
    values = 'Values',
    formats = 'Formats',
  }

  interface WorksheetCollection {
    items: Worksheet[];
    count: number;
    getActiveWorksheet(): Worksheet;
    getItem(name: string): Worksheet;
    getItemAt(index: number): Worksheet;
    add(name?: string): Worksheet;
    load(propertyNames?: string | string[]): WorksheetCollection;
  }

  enum SheetVisibility {
    visible = 'Visible',
    hidden = 'Hidden',
    veryHidden = 'VeryHidden',
  }

  interface Worksheet {
    tables: TableScopedCollection;
    getUsedRange(): Range | null;
    getRange(address: string): Range;
    getCell(row: number, column: number): Range;
    getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number): Range;
    name: string;
    position: number;
    visibility: SheetVisibility;
    tabColor: string;
    freezePanes: WorksheetFreezePanes;
    load(propertyNames?: string | string[]): Worksheet;
    delete(): void;
    copy(): Worksheet;
    activate(): void;
  }

  interface WorksheetFreezePanes {
    freezeRows(count: number): void;
    freezeColumns(count: number): void;
    unfreeze(): void;
  }

  enum ClearApplyTo {
    all = 'All',
    formats = 'Formats',
    contents = 'Contents',
    hyperlinks = 'Hyperlinks',
    removeHyperlinks = 'RemoveHyperlinks',
  }

  enum SortOrientation {
    rows = 'Rows',
    columns = 'Columns',
  }

  interface SortField {
    key: number;
    ascending?: boolean;
    sortOn?: SortOn;
  }

  enum SortOn {
    value = 'Value',
    cellColor = 'CellColor',
    fontColor = 'FontColor',
    icon = 'Icon',
  }

  interface RangeSort {
    apply(
      fields: SortField[],
      matchCase?: boolean,
      hasHeaders?: boolean,
      orientation?: SortOrientation | 'Rows' | 'Columns',
    ): void;
  }

  enum AutoFillType {
    fillDefault = 'FillDefault',
    fillCopy = 'FillCopy',
    fillSeries = 'FillSeries',
    fillFormats = 'FillFormats',
    fillValues = 'FillValues',
    fillDays = 'FillDays',
    fillWeekdays = 'FillWeekdays',
    fillMonths = 'FillMonths',
    fillYears = 'FillYears',
    linearTrend = 'LinearTrend',
    growthTrend = 'GrowthTrend',
  }

  type HorizontalAlignment =
    | 'General'
    | 'Left'
    | 'Center'
    | 'Right'
    | 'Fill'
    | 'Justify'
    | 'CenterAcrossSelection'
    | 'Distributed';

  type VerticalAlignment = 'Top' | 'Center' | 'Bottom' | 'Justify' | 'Distributed';

  interface Range {
    values: any[][];
    formulas: any[][];
    numberFormat: any[][];
    format: RangeFormat;
    borders: RangeBorderCollection;
    rowHidden: boolean;
    columnHidden: boolean;
    load(propertyNames?: string | string[]): Range;
    delete(shift: DeleteShiftDirection): void;
    getEntireRow(): Range;
    getEntireColumn(): Range;
    insert(shift: InsertShiftDirection): void;
    merge(across?: boolean): void;
    unmerge(): void;
    clear(applyTo?: ClearApplyTo): void;
    autoFill(destinationRange: Range, autoFillType: AutoFillType): void;
    copyFrom(
      sourceRange: Range,
      copyType: RangeCopyType,
      skipBlanks?: boolean,
      transpose?: boolean,
    ): void;
    getResizedRange(deltaRows: number, deltaColumns: number): Range;
    getComment(): Comment;
    rowCount: number;
    columnCount: number;
    row: number;
    column: number;
    address: string;
    sort: RangeSort;
    select(): void;
  }

  interface Comment {
    add(content: string): void;
    delete(): void;
  }

  interface RangeFormat {
    fill: RangeFill;
    borders: RangeBorderCollection;
    font: RangeFont;
    rowHeight: number;
    columnWidth: number;
    horizontalAlignment: HorizontalAlignment;
    verticalAlignment: VerticalAlignment;
    wrapText: boolean;
    autofitColumns(): void;
  }

  interface RangeFont {
    name: string;
    bold: boolean;
    italic: boolean;
    underline: string;
    size: number;
    color: string;
  }

  interface RangeFill {
    color: string | null;
    pattern: FillPattern;
  }

  type FillPattern =
    | 'None'
    | 'Solid'
    | 'Gray50'
    | 'Gray75'
    | 'Gray25'
    | 'Horizontal'
    | 'Vertical'
    | 'Down'
    | 'Up'
    | 'Checker'
    | 'SemiGray75'
    | 'LightHorizontal'
    | 'LightVertical'
    | 'LightDown'
    | 'LightUp'
    | 'Grid'
    | 'CrissCross'
    | 'Gray16'
    | 'Gray8';

  interface RangeBorderCollection {
    getItem(side: BorderSide): RangeBorder;
  }

  interface RangeBorder {
    color: string;
    style: BorderStyle;
    weight: BorderWeight;
  }

  type BorderSide = 'EdgeTop' | 'EdgeBottom' | 'EdgeLeft' | 'EdgeRight';

  type BorderStyle = 'None' | 'Continuous' | 'Dash' | 'DashDot' | 'DashDotDot' | 'Dot' | 'Double' | 'SlantDashDot';

  type BorderWeight = 'Thin' | 'Medium' | 'Thick';

  type DeleteShiftDirection = 'Up' | 'Left';

  type InsertShiftDirection = 'Down' | 'Up' | 'Right' | 'Left';

  function run<T>(
    callback: (context: RequestContext) => Promise<T> | T
  ): Promise<T>;
}
