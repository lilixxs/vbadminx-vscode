/**
 * Excel 宿主对象的**内建近似模型**。
 *
 * 目的：导出的 VBA 工程通常以 Excel 为宿主，但没有附带 Excel 类型库索引。
 * 未加载索引时，`Dim ws As Worksheet` 会得到“类型信息缺失”。本模块提供一组
 * 覆盖常见对象（Application / Workbook / Worksheet / Range / ListObject / Chart…）
 * 的声明，默认参与语义分析，使这些类型及其常用成员可解析。
 *
 * 约束：
 * - 这是**近似模型**，不是 Microsoft Excel 类型库的完整拷贝；签名以常见用法为准。
 * - 工程显式加载的同名类型（例如真实 Excel 索引）优先：这些类型不会重复注入。
 * - 仅补充类型与宿主全局符号，不改变 VBA 语法、条件编译或诊断保守性规则。
 * - 非 Excel 宿主（Word/Outlook 等）不注入；未配置 host 时按 Excel 处理。
 */
import type { PortableTypeIndex, VbaParameter, VbaSymbol } from './model';

export const EXCEL_HOST_LIBRARY_NAME = 'Excel 16.0 Host Model (built-in)';
const EXCEL_HOST_GUID = '{00000000-0000-4000-8000-0000E0CE1101}';

type MemberKind = 'method' | 'property' | 'constant' | 'event';

interface MemberDef {
  name: string;
  kind: MemberKind;
  returnType?: string;
  parameters?: VbaParameter[];
  documentation?: string;
}

interface TypeDef {
  name: string;
  members: MemberDef[];
  documentation?: string;
}

const param = (name: string, typeName = 'Variant', optional = false): VbaParameter => ({ name, typeName, optional, byRef: false });
const method = (name: string, returnType = 'void', parameters: VbaParameter[] = [], documentation?: string): MemberDef =>
  ({ name, kind: 'method', returnType, parameters, documentation });
const prop = (name: string, returnType = 'Variant', parameters?: VbaParameter[], documentation?: string): MemberDef =>
  ({ name, kind: 'property', returnType, parameters, documentation });

const RANGE_PARAMS = [param('Cell1'), param('Cell2', 'Variant', true)];

const TYPES: TypeDef[] = [
  {
    name: 'Application',
    documentation: 'Excel 应用程序对象（内建近似模型）。',
    members: [
      prop('Application', 'Application'),
      prop('ActiveCell', 'Range'),
      prop('ActiveChart', 'Chart'),
      prop('ActiveSheet', 'Worksheet'),
      prop('ActiveWindow', 'Window'),
      prop('ActiveWorkbook', 'Workbook'),
      prop('ThisWorkbook', 'Workbook'),
      prop('Selection', 'Range'),
      prop('Sheets', 'Sheets'),
      prop('Worksheets', 'Worksheets'),
      prop('Workbooks', 'Workbooks'),
      prop('Names', 'Names', [param('Index', 'Variant', true)]),
      prop('Charts', 'Charts', [param('Index', 'Variant', true)]),
      prop('Windows', 'Windows', [param('Index', 'Variant', true)]),
      prop('Range', 'Range', RANGE_PARAMS),
      prop('Cells', 'Range', [param('RowIndex', 'Long'), param('ColumnIndex', 'Long', true)]),
      prop('Rows', 'Range', [param('Index', 'Variant', true)]),
      prop('Columns', 'Range', [param('Index', 'Variant', true)]),
      prop('ScreenUpdating', 'Boolean'),
      prop('DisplayAlerts', 'Boolean'),
      prop('EnableEvents', 'Boolean'),
      prop('Calculation', 'Long'),
      prop('StatusBar', 'Variant'),
      prop('CutCopyMode', 'Boolean'),
      prop('AutomationSecurity', 'Long'),
      method('Union', 'Range', [param('Arg1', 'Range'), param('Arg2', 'Range', true), param('Arg3', 'Range', true)]),
      method('Intersect', 'Range', [param('Arg1', 'Range'), param('Arg2', 'Range'), param('Arg3', 'Range', true)]),
      method('Evaluate', 'Variant', [param('Name')]),
      method('Run', 'Variant', [param('Macro'), param('Arg1', 'Variant', true)]),
      method('Calculate'),
      method('Quit'),
      method('Volatile', 'void', [param('Volatile', 'Variant', true)]),
      method('Wait', 'void', [param('Time')]),
    ],
  },
  {
    name: 'Workbook',
    documentation: 'Excel 工作簿对象（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('FullName', 'String'),
      prop('Path', 'String'),
      prop('CodeName', 'String'),
      prop('ActiveSheet', 'Worksheet'),
      prop('Worksheets', 'Worksheets'),
      prop('Sheets', 'Sheets'),
      prop('Charts', 'Charts'),
      prop('Names', 'Names'),
      prop('Windows', 'Windows'),
      prop('ReadOnly', 'Boolean'),
      prop('Saved', 'Boolean'),
      prop('HasPassword', 'Boolean'),
      prop('Parent', 'Application'),
      method('Activate'),
      method('Save'),
      method('Close', 'void', [param('SaveChanges', 'Variant', true), param('Filename', 'Variant', true), param('RouteWorkbook', 'Variant', true)]),
      method('SaveAs', 'void', [param('Filename'), param('FileFormat', 'Variant', true)]),
      method('Protect', 'void', [param('Password', 'Variant', true)]),
      method('Unprotect', 'void', [param('Password', 'Variant', true)]),
      method('PrintOut', 'void', [param('From', 'Variant', true), param('To', 'Variant', true), param('Copies', 'Variant', true)]),
      method('ExportAsFixedFormat', 'void', [param('Type', 'Long'), param('Filename', 'Variant', true)]),
    ],
  },
  {
    name: 'Workbooks',
    documentation: '打开的工作簿集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'Workbook', [param('Index')]),
      method('Add', 'Workbook', [param('Template', 'Variant', true)]),
      method('Open', 'Workbook', [param('Filename'), param('UpdateLinks', 'Variant', true), param('ReadOnly', 'Variant', true)]),
      method('Close'),
    ],
  },
  {
    name: 'Worksheet',
    documentation: 'Excel 工作表对象（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('CodeName', 'String'),
      prop('Index', 'Long'),
      prop('Visible', 'Long'),
      prop('StandardWidth', 'Double'),
      prop('UsedRange', 'Range'),
      prop('Range', 'Range', RANGE_PARAMS),
      prop('Cells', 'Range', [param('Row', 'Long'), param('Column', 'Long', true)]),
      prop('Rows', 'Range', [param('Index', 'Variant', true)]),
      prop('Columns', 'Range', [param('Index', 'Variant', true)]),
      prop('ListObjects', 'ListObjects', [param('Index', 'Variant', true)]),
      prop('Shapes', 'Shapes', [param('Index', 'Variant', true)]),
      prop('Charts', 'Charts', [param('Index', 'Variant', true)]),
      prop('PageSetup', 'PageSetup'),
      prop('Sort', 'Sort'),
      prop('AutoFilter', 'AutoFilter'),
      prop('Parent', 'Workbook'),
      method('Activate'),
      method('Select', 'void', [param('Replace', 'Variant', true)]),
      method('Copy', 'void', [param('Before', 'Variant', true), param('After', 'Variant', true)]),
      method('Delete'),
      method('Move', 'void', [param('Before', 'Variant', true), param('After', 'Variant', true)]),
      method('Calculate'),
      method('Evaluate', 'Variant', [param('Name')]),
      method('Protect', 'void', [param('Password', 'Variant', true)]),
      method('Unprotect', 'void', [param('Password', 'Variant', true)]),
    ],
  },
  {
    name: 'Worksheets',
    documentation: '工作表集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Visible', 'Long'),
      prop('Item', 'Worksheet', [param('Index')]),
      // 默认成员近似：Worksheets("x").Range(...) 在真实 VBA 中经由默认成员返回 Worksheet。
      prop('Range', 'Range', RANGE_PARAMS),
      prop('Cells', 'Range', [param('Row', 'Long'), param('Column', 'Long', true)]),
      prop('UsedRange', 'Range'),
      method('Add', 'Worksheet', [param('Before', 'Variant', true), param('After', 'Variant', true), param('Count', 'Variant', true), param('Type', 'Variant', true)]),
      method('Delete'),
      method('Select', 'void', [param('Replace', 'Variant', true)]),
      method('Copy', 'void', [param('Before', 'Variant', true), param('After', 'Variant', true)]),
      method('Move', 'void', [param('Before', 'Variant', true), param('After', 'Variant', true)]),
    ],
  },
  {
    name: 'Sheets',
    documentation: '工作表与图表工作表集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Visible', 'Long'),
      prop('Item', 'Worksheet', [param('Index')]),
      prop('Range', 'Range', RANGE_PARAMS),
      prop('Cells', 'Range', [param('Row', 'Long'), param('Column', 'Long', true)]),
      method('Add', 'Worksheet', [param('Before', 'Variant', true), param('After', 'Variant', true), param('Count', 'Variant', true), param('Type', 'Variant', true)]),
      method('Delete'),
      method('Select', 'void', [param('Replace', 'Variant', true)]),
      method('Copy', 'void', [param('Before', 'Variant', true), param('After', 'Variant', true)]),
      method('Move', 'void', [param('Before', 'Variant', true), param('After', 'Variant', true)]),
    ],
  },
  {
    name: 'Range',
    documentation: '单元格区域对象（内建近似模型）。',
    members: [
      prop('Value', 'Variant'),
      prop('Value2', 'Variant'),
      prop('Text', 'String'),
      prop('Formula', 'Variant'),
      prop('FormulaR1C1', 'Variant'),
      prop('FormulaLocal', 'Variant'),
      prop('FormulaArray', 'Variant'),
      prop('Address', 'String', [param('RowAbsolute', 'Variant', true), param('ColumnAbsolute', 'Variant', true), param('ReferenceStyle', 'Variant', true), param('External', 'Variant', true), param('RelativeTo', 'Variant', true)]),
      prop('Row', 'Long'),
      prop('Column', 'Long'),
      prop('Count', 'Long'),
      prop('CountLarge', 'Variant'),
      prop('Rows', 'Range'),
      prop('Columns', 'Range'),
      prop('Cells', 'Range', [param('Row', 'Long'), param('Column', 'Long', true)]),
      prop('Areas', 'Range'),
      prop('EntireRow', 'Range'),
      prop('EntireColumn', 'Range'),
      prop('CurrentRegion', 'Range'),
      prop('Worksheet', 'Worksheet'),
      prop('Parent', 'Worksheet'),
      prop('Font', 'Font'),
      prop('Interior', 'Interior'),
      prop('Borders', 'Borders'),
      prop('NumberFormat', 'Variant'),
      prop('HorizontalAlignment', 'Long'),
      prop('VerticalAlignment', 'Long'),
      prop('WrapText', 'Boolean'),
      prop('MergeCells', 'Boolean'),
      prop('Locked', 'Boolean'),
      prop('Hidden', 'Boolean'),
      prop('Comment', 'Object'),
      prop('Validation', 'Validation'),
      method('Offset', 'Range', [param('RowOffset', 'Long', true), param('ColumnOffset', 'Long', true)]),
      method('Resize', 'Range', [param('RowSize', 'Long', true), param('ColumnSize', 'Long', true)]),
      method('End', 'Range', [param('Direction', 'Long')]),
      method('Item', 'Range', [param('Row', 'Long'), param('Column', 'Long', true)]),
      method('Clear'),
      method('ClearContents'),
      method('ClearFormats'),
      method('Copy', 'void', [param('Destination', 'Variant', true)]),
      method('Cut', 'void', [param('Destination', 'Variant', true)]),
      method('PasteSpecial', 'void', [param('Paste', 'Long'), param('Operation', 'Variant', true), param('SkipBlanks', 'Variant', true), param('Transpose', 'Variant', true)]),
      method('Select'),
      method('Activate'),
      method('Delete', 'Variant', [param('Shift', 'Variant', true)]),
      method('Insert', 'Variant', [param('Shift', 'Variant', true), param('CopyOrigin', 'Variant', true)]),
      method('Merge', 'void', [param('Across', 'Variant', true)]),
      method('UnMerge'),
      method('AutoFit'),
      method('Find', 'Range', [param('What'), param('After', 'Variant', true), param('LookIn', 'Variant', true), param('LookAt', 'Variant', true)]),
      method('Replace', 'Boolean', [param('What'), param('Replacement'), param('LookAt', 'Variant', true), param('SearchOrder', 'Variant', true)]),
      method('Sort', 'Variant', [param('Key1', 'Variant', true), param('Order1', 'Variant', true)]),
      method('AutoFilter', 'Variant', [param('Field', 'Variant', true), param('Criteria1', 'Variant', true), param('Operator', 'Variant', true), param('Criteria2', 'Variant', true)]),
      method('SpecialCells', 'Range', [param('Type', 'Long'), param('Value', 'Variant', true)]),
      method('CopyFromRecordset', 'Long', [param('Data', 'Object'), param('MaxRows', 'Variant', true), param('MaxColumns', 'Variant', true)]),
    ],
  },
  {
    name: 'Font',
    documentation: '字体对象（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('Size', 'Variant'),
      prop('Bold', 'Variant'),
      prop('Italic', 'Variant'),
      prop('Underline', 'Variant'),
      prop('Strikethrough', 'Variant'),
      prop('Color', 'Variant'),
      prop('ColorIndex', 'Variant'),
      prop('Parent', 'Object'),
    ],
  },
  {
    name: 'Interior',
    documentation: '单元格内部填充（内建近似模型）。',
    members: [
      prop('Color', 'Variant'),
      prop('ColorIndex', 'Variant'),
      prop('Pattern', 'Variant'),
      prop('PatternColorIndex', 'Variant'),
      prop('Parent', 'Object'),
    ],
  },
  {
    name: 'Borders',
    documentation: '区域边框集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'Border', [param('Index', 'Long')]),
      prop('LineStyle', 'Variant'),
      prop('Weight', 'Variant'),
      prop('Color', 'Variant'),
      prop('ColorIndex', 'Variant'),
    ],
  },
  {
    name: 'Border',
    documentation: '单条边框（内建近似模型）。',
    members: [
      prop('LineStyle', 'Variant'),
      prop('Weight', 'Variant'),
      prop('Color', 'Variant'),
      prop('ColorIndex', 'Variant'),
      prop('Parent', 'Object'),
    ],
  },
  {
    name: 'Name',
    documentation: '定义名称（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('RefersTo', 'String'),
      prop('RefersToR1C1', 'String'),
      prop('Value', 'String'),
      method('Delete'),
      prop('Parent', 'Object'),
    ],
  },
  {
    name: 'Names',
    documentation: '定义名称集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'Name', [param('Index', 'Variant')]),
      method('Add', 'Name', [param('Name'), param('RefersTo'), param('Visible', 'Variant', true)]),
      method('Delete'),
    ],
  },
  {
    name: 'ListObjects',
    documentation: '表格（ListObject）集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'ListObject', [param('Index', 'Variant')]),
      method('Add', 'ListObject', [param('SourceType', 'Long'), param('Source', 'Variant'), param('Destination', 'Variant', true)]),
    ],
  },
  {
    name: 'ListObject',
    documentation: 'Excel 表格对象（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('Range', 'Range'),
      prop('DataBodyRange', 'Range'),
      prop('HeaderRowRange', 'Range'),
      prop('ListColumns', 'ListColumns'),
      prop('ListRows', 'Object'),
      method('Delete'),
      method('Refresh'),
    ],
  },
  {
    name: 'ListColumns',
    documentation: '表格列集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'ListColumn', [param('Index', 'Variant')]),
      method('Add', 'ListColumn', [param('Position', 'Variant', true)]),
    ],
  },
  {
    name: 'ListColumn',
    documentation: '表格列（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('Range', 'Range'),
      prop('DataBodyRange', 'Range'),
      prop('Parent', 'Object'),
    ],
  },
  {
    name: 'Shapes',
    documentation: '图形集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'Shape', [param('Index', 'Variant')]),
      method('AddShape', 'Shape', [param('Type', 'Long'), param('Left', 'Single'), param('Top', 'Single'), param('Width', 'Single'), param('Height', 'Single')]),
      method('AddTextbox', 'Shape', [param('Orientation', 'Long'), param('Left', 'Single'), param('Top', 'Single'), param('Width', 'Single'), param('Height', 'Single')]),
    ],
  },
  {
    name: 'Shape',
    documentation: '图形对象（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('Visible', 'Long'),
      prop('Left', 'Single'),
      prop('Top', 'Single'),
      prop('Width', 'Single'),
      prop('Height', 'Single'),
      prop('Fill', 'Object'),
      prop('Line', 'Object'),
      prop('TextFrame', 'Object'),
      prop('OnAction', 'String'),
      method('Delete'),
      method('Select', 'void', [param('Replace', 'Variant', true)]),
    ],
  },
  {
    name: 'Charts',
    documentation: '图表工作表集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'Chart', [param('Index', 'Variant')]),
      method('Add', 'Chart', [param('Before', 'Variant', true), param('After', 'Variant', true), param('Count', 'Variant', true)]),
    ],
  },
  {
    name: 'Chart',
    documentation: '图表对象（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('ChartType', 'Long'),
      prop('HasTitle', 'Boolean'),
      prop('ChartTitle', 'Object'),
      prop('SeriesCollection', 'SeriesCollection', [param('Index', 'Variant', true)]),
      method('SetSourceData', 'void', [param('Source', 'Range'), param('PlotBy', 'Variant', true)]),
      method('Activate'),
      method('Delete'),
      prop('Parent', 'Object'),
    ],
  },
  {
    name: 'SeriesCollection',
    documentation: '图表数据系列集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'Series', [param('Index', 'Variant')]),
      method('NewSeries', 'Series'),
      method('Extend', 'Series', [param('Source', 'Range'), param('Rowcol', 'Variant', true)]),
    ],
  },
  {
    name: 'Series',
    documentation: '图表数据系列（内建近似模型）。',
    members: [
      prop('Name', 'String'),
      prop('Values', 'Variant'),
      prop('XValues', 'Variant'),
      prop('Formula', 'String'),
      method('Delete'),
    ],
  },
  {
    name: 'PageSetup',
    documentation: '页面设置（内建近似模型）。',
    members: [
      prop('Orientation', 'Long'),
      prop('PrintArea', 'String'),
      prop('PrintTitleRows', 'String'),
      prop('Zoom', 'Variant'),
      prop('FitToPagesWide', 'Variant'),
      prop('FitToPagesTall', 'Variant'),
      prop('PaperSize', 'Long'),
      prop('LeftMargin', 'Double'),
      prop('RightMargin', 'Double'),
      prop('TopMargin', 'Double'),
      prop('BottomMargin', 'Double'),
    ],
  },
  {
    name: 'Window',
    documentation: '窗口对象（内建近似模型）。',
    members: [
      prop('Caption', 'String'),
      prop('Visible', 'Boolean'),
      prop('WindowState', 'Long'),
      prop('ActiveCell', 'Range'),
      method('Activate'),
    ],
  },
  {
    name: 'Windows',
    documentation: '窗口集合（内建近似模型）。',
    members: [
      prop('Count', 'Long'),
      prop('Item', 'Window', [param('Index', 'Variant')]),
      method('Arrange', 'Variant', [param('ArrangeStyle', 'Variant', true)]),
    ],
  },
  {
    name: 'Validation',
    documentation: '数据验证（内建近似模型）。',
    members: [
      prop('Type', 'Long'),
      prop('Formula1', 'String'),
      prop('Formula2', 'String'),
      prop('Value', 'Boolean'),
      method('Add', 'void', [param('Type', 'Long'), param('AlertStyle', 'Variant', true), param('Operator', 'Variant', true), param('Formula1', 'Variant', true), param('Formula2', 'Variant', true)]),
      method('Delete'),
    ],
  },
  {
    name: 'Sort',
    documentation: '排序设置（内建近似模型）。',
    members: [
      prop('SortFields', 'Object'),
      method('Apply'),
      method('Clear'),
    ],
  },
  {
    name: 'AutoFilter',
    documentation: '自动筛选（内建近似模型）。',
    members: [
      prop('Range', 'Range'),
      prop('Field', 'Long'),
      prop('Criteria1', 'String'),
      prop('Criteria2', 'String'),
      prop('Operator', 'Long'),
      method('ApplyFilter'),
      method('ShowAllData'),
    ],
  },
];

/** 宿主全局符号：可直接使用、无需声明的 Excel 对象与常量。 */
interface GlobalDef {
  name: string;
  kind: 'variable' | 'function' | 'constant';
  typeName: string;
  parameters?: VbaParameter[];
  value?: string;
}

const GLOBALS: GlobalDef[] = [
  { name: 'Application', kind: 'variable', typeName: 'Application' },
  { name: 'ThisWorkbook', kind: 'variable', typeName: 'Workbook' },
  { name: 'ActiveWorkbook', kind: 'variable', typeName: 'Workbook' },
  { name: 'ActiveSheet', kind: 'variable', typeName: 'Worksheet' },
  { name: 'ActiveCell', kind: 'variable', typeName: 'Range' },
  { name: 'Selection', kind: 'variable', typeName: 'Range' },
  { name: 'ActiveWindow', kind: 'variable', typeName: 'Window' },
  { name: 'Range', kind: 'function', typeName: 'Range', parameters: RANGE_PARAMS },
  { name: 'Cells', kind: 'function', typeName: 'Range', parameters: [param('RowIndex', 'Long'), param('ColumnIndex', 'Long', true)] },
  { name: 'Rows', kind: 'function', typeName: 'Range', parameters: [param('Index', 'Variant', true)] },
  { name: 'Columns', kind: 'function', typeName: 'Range', parameters: [param('Index', 'Variant', true)] },
  { name: 'Sheets', kind: 'function', typeName: 'Worksheet', parameters: [param('Index', 'Variant')] },
  { name: 'Worksheets', kind: 'function', typeName: 'Worksheet', parameters: [param('Index', 'Variant')] },
  { name: 'Workbooks', kind: 'function', typeName: 'Workbook', parameters: [param('Index', 'Variant')] },
  { name: 'Names', kind: 'function', typeName: 'Name', parameters: [param('Index', 'Variant')] },
  { name: 'Charts', kind: 'function', typeName: 'Chart', parameters: [param('Index', 'Variant')] },
  { name: 'Windows', kind: 'function', typeName: 'Window', parameters: [param('Index', 'Variant')] },
  ...([
    ['xlUp', -4162], ['xlDown', -4121], ['xlToLeft', -4159], ['xlToRight', -4161],
    ['xlCalculationAutomatic', -4105], ['xlCalculationManual', -4135], ['xlCalculationSemiautomatic', 2],
    ['xlYes', 1], ['xlNo', 2], ['xlGuess', 0],
    ['xlWhole', 1], ['xlPart', 2],
    ['xlByRows', 1], ['xlByColumns', 2],
    ['xlAscending', 1], ['xlDescending', 2],
    ['xlNone', -4142], ['xlContinuous', 1], ['xlThin', 2], ['xlMedium', -4138],
    ['xlCenter', -4108], ['xlLeft', -4131], ['xlRight', -4152],
    ['xlTop', -4160], ['xlBottom', -4107],
  ] as Array<[string, number]>).map(([name, value]): GlobalDef => ({ name, kind: 'constant', typeName: 'Long', value: String(value) })),
];

let cachedIndex: PortableTypeIndex | undefined;
let cachedGlobals: VbaSymbol[] | undefined;

/** 生成内建 Excel 宿主类型索引；`bitness` 供工程位数显示使用。 */
export function excelHostIndex(bitness: 32 | 64 = 64): PortableTypeIndex {
  if (cachedIndex && cachedIndex.library.bitness === bitness) { return cachedIndex; }
  cachedIndex = {
    formatVersion: 1,
    library: {
      name: EXCEL_HOST_LIBRARY_NAME,
      guid: EXCEL_HOST_GUID,
      major: 16,
      minor: 0,
      lcid: 0,
      bitness,
    },
    types: TYPES.map(type => ({
      name: type.name,
      kind: 'class' as const,
      documentation: type.documentation,
      members: type.members.map(member => ({
        name: member.name,
        kind: member.kind,
        returnType: member.returnType,
        parameters: member.parameters,
        documentation: member.documentation,
      })),
    })),
  };
  return cachedIndex;
}

function hostSymbolId(name: string): string {
  return `vba-host:/excel/${encodeURIComponent(name)}`;
}

function baseSymbol(name: string): VbaSymbol {
  const uri = hostSymbolId(name);
  return {
    id: uri,
    name,
    kind: 'variable',
    uri,
    range: { start: { line: 0, character: 0 }, end: { line: 0, character: 0 } },
    selectionRange: { start: { line: 0, character: 0 }, end: { line: 0, character: 0 } },
    access: 'public',
    signature: name,
    external: true,
  };
}

/** 宿主全局对象与常量；返回的符号已标记 external，不参与重命名。 */
export function excelHostGlobals(): VbaSymbol[] {
  if (cachedGlobals) { return cachedGlobals; }
  cachedGlobals = GLOBALS.map(definition => {
    const symbol = baseSymbol(definition.name);
    if (definition.kind === 'constant') {
      symbol.kind = 'constant';
      symbol.typeName = definition.typeName;
      symbol.signature = `Const ${definition.name} As ${definition.typeName} = ${definition.value}`;
      symbol.documentation = 'Excel 内建常量（近似模型）。';
      return symbol;
    }
    symbol.kind = definition.kind === 'function' ? 'function' : 'variable';
    symbol.typeName = definition.typeName;
    const parameters = definition.parameters ?? [];
    symbol.parameters = parameters.length ? parameters : undefined;
    symbol.signature = definition.kind === 'function'
      ? `Function ${definition.name}(${parameters.map(item => `${item.optional ? 'Optional ' : ''}${item.name} As ${item.typeName}`).join(', ')}) As ${definition.typeName}`
      : `${definition.name} As ${definition.typeName}`;
    symbol.documentation = 'Excel 宿主对象（内建近似模型）。';
    return symbol;
  });
  return cachedGlobals;
}

/** 是否应按 Excel 宿主处理：仅 Excel，或未配置宿主（默认假定 Excel）。 */
export function usesExcelHostModel(host: string | undefined): boolean {
  if (!host || !host.trim()) { return true; }
  return /excel/i.test(host);
}

/** 判断某个索引是否就是本内建模型。 */
export function isExcelHostIndex(index: PortableTypeIndex): boolean {
  return index.library.name === EXCEL_HOST_LIBRARY_NAME;
}
