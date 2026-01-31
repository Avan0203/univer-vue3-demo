/**
 * 纯前端 ExcelJS ↔ Univer IWorkbookData 转换
 * 支持：文本、数字、布尔、日期、公式、富文本（双向 p ↔ richText）
 */
import type { IWorkbookData } from "@univerjs/core";
import { BooleanNumber, LocaleType } from "@univerjs/core";
import { CellValueType } from "@univerjs/core";
import type { IDocumentBody, IDocumentData, ITextRun, ITextStyle } from "@univerjs/core";
import ExcelJS from "exceljs";
import type { Workbook, Worksheet } from "exceljs";
import { ValueType } from "exceljs";

const DEFAULT_ROW_COUNT = 1000;
const DEFAULT_COLUMN_COUNT = 20;

/** Univer ICellData 最小表示（v/t/f/p） */
interface UniverCellLike {
  v?: string | number | boolean;
  t?: CellValueType;
  f?: string;
  p?: IDocumentData;
}

/**
 * ExcelJS Cell → Univer ICellData（支持 Number/String/Boolean/Date/Formula/RichText/Hyperlink/Error）
 */
function excelCellToUniverCell(
  value: ExcelJS.CellValue,
  cellType: ValueType,
  formula?: string
): UniverCellLike | undefined {
  if (value === null || value === undefined) return undefined;

  // 公式：value 可能为 { formula, result } 或仅为 result；优先从 value 取 formula/result
  const formulaVal = value as { formula?: string; result?: unknown };
  if (
    typeof formulaVal === "object" &&
    "formula" in formulaVal &&
    typeof formulaVal.formula === "string" &&
    formulaVal.formula.length > 0
  ) {
    const result = formulaVal.result !== undefined ? formulaVal.result : value;
    const v = cellValueToCellValue(result);
    if (v !== undefined)
      return { v, f: formulaVal.formula, t: inferCellValueType(v) };
    return undefined;
  }
  if (formula && formula.length > 0) {
    const v = cellValueToCellValue(value);
    if (v !== undefined) return { v, f: formula, t: inferCellValueType(v) };
  }

  switch (cellType) {
    case ValueType.Number:
      return typeof value === "number" ? { v: value, t: CellValueType.NUMBER } : undefined;
    case ValueType.Boolean:
      return typeof value === "boolean" ? { v: value, t: CellValueType.BOOLEAN } : undefined;
    case ValueType.Date:
      if (value instanceof Date) return { v: value.getTime(), t: CellValueType.NUMBER };
      return undefined;
    case ValueType.String:
    case ValueType.SharedString:
      return { v: String(value), t: CellValueType.STRING };
    case ValueType.RichText: {
      const rich = value as { richText?: ExcelJS.RichText[] };
      const segments = rich.richText;
      if (!segments?.length) return { v: "", t: CellValueType.STRING };
      const plainText = segments.map((r) => r.text).join("");
      const p = excelRichTextToUniverDocument(segments);
      return { v: plainText, t: CellValueType.STRING, p };
    }
    case ValueType.Hyperlink: {
      const link = value as { text?: string };
      return { v: link.text ?? String(value), t: CellValueType.STRING };
    }
    case ValueType.Formula:
      const r = cellValueToCellValue(value);
      if (r !== undefined && formula)
        return { v: r, f: formula, t: inferCellValueType(r) };
      return r !== undefined ? { v: r, t: inferCellValueType(r) } : undefined;
    case ValueType.Error:
      return { v: String(value), t: CellValueType.STRING };
    default:
      return fallbackCellValue(value);
  }
}

function cellValueToCellValue(val: unknown): string | number | boolean | undefined {
  if (val === null || val === undefined) return undefined;
  if (typeof val === "string" || typeof val === "number" || typeof val === "boolean") return val;
  if (val instanceof Date) return val.getTime();
  if (typeof val === "object" && "text" in (val as { text?: string }))
    return (val as { text: string }).text ?? "";
  return String(val);
}

function inferCellValueType(v: string | number | boolean): CellValueType {
  if (typeof v === "number") return CellValueType.NUMBER;
  if (typeof v === "boolean") return CellValueType.BOOLEAN;
  return CellValueType.STRING;
}

function fallbackCellValue(value: unknown): UniverCellLike | undefined {
  if (value === null || value === undefined) return undefined;
  if (typeof value === "string") return { v: value, t: CellValueType.STRING };
  if (typeof value === "number") return { v: value, t: CellValueType.NUMBER };
  if (typeof value === "boolean") return { v: value, t: CellValueType.BOOLEAN };
  if (value instanceof Date) return { v: value.getTime(), t: CellValueType.NUMBER };
  if (typeof value === "object" && "richText" in (value as { richText?: unknown })) {
    const rich = (value as { richText: ExcelJS.RichText[] }).richText;
    if (rich.length) {
      const plainText = rich.map((r) => r.text).join("");
      const p = excelRichTextToUniverDocument(rich);
      return { v: plainText, t: CellValueType.STRING, p };
    }
    return { v: "", t: CellValueType.STRING };
  }
  return { v: String(value), t: CellValueType.STRING };
}

/** ExcelJS RichText[] → Univer IDocumentData（用于 ICellData.p） */
function excelRichTextToUniverDocument(richText: ExcelJS.RichText[]): IDocumentData {
  const dataStreamParts: string[] = [];
  const textRuns: ITextRun[] = [];
  let offset = 0;
  for (const seg of richText) {
    const text = seg.text ?? "";
    dataStreamParts.push(text);
    const st = offset;
    offset += text.length;
    const ed = offset;
    const ts = excelFontToUniverTS(seg.font);
    if (text.length > 0) {
      textRuns.push(ts ? { st, ed, ts } : { st, ed });
    }
  }
  const dataStream = dataStreamParts.join("") + "\r";
  const body: IDocumentBody = {
    dataStream,
    textRuns: textRuns.length ? textRuns : undefined,
    paragraphs: [{ startIndex: 0 }],
  };
  return {
    id: `doc-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`,
    documentStyle: {},
    body,
  };
}

/** ExcelJS Font → Univer ITextStyle（仅常用：ff/fs/bl/it/ul/cl） */
function excelFontToUniverTS(
  font?: Partial<{ name: string; size: number; bold: boolean; italic: boolean; underline: unknown; color: { argb?: string } }>
): ITextStyle | undefined {
  if (!font || Object.keys(font).length === 0) return undefined;
  const ts: ITextStyle = {};
  if (font.name != null) ts.ff = font.name;
  if (font.size != null) ts.fs = font.size;
  if (font.bold != null) ts.bl = font.bold ? BooleanNumber.TRUE : BooleanNumber.FALSE;
  if (font.italic != null) ts.it = font.italic ? BooleanNumber.TRUE : BooleanNumber.FALSE;
  if (font.underline != null && font.underline !== false && font.underline !== "none") {
    ts.ul = { s: BooleanNumber.TRUE };
  }
  if (font.color?.argb != null) {
    const hex = font.color.argb;
    ts.cl = { rgb: hex.length === 8 ? "#" + hex.slice(2) : hex.startsWith("#") ? hex : "#" + hex };
  }
  return Object.keys(ts).length ? ts : undefined;
}

/**
 * ExcelJS Workbook → IWorkbookData（文本/数字/布尔/日期/公式/富文本等）
 */
export function excelJSToWorkbookData(workbook: Workbook): IWorkbookData {
  const sheetOrder: string[] = [];
  const sheets: IWorkbookData["sheets"] = {};
  const ts = Date.now();
  const workbookId = `workbook-${ts}`;

  workbook.eachSheet((worksheet, index) => {
    const sheetId = `sheet-${index}-${ts}`;
    sheetOrder.push(sheetId);
    const { cellData, rowCount, columnCount } = worksheetToCellData(worksheet);
    sheets[sheetId] = buildMinimalSheet(sheetId, worksheet.name || `Sheet${index + 1}`, cellData, rowCount, columnCount);
  });

  if (sheetOrder.length === 0) {
    const sheetId = `sheet-0-${ts}`;
    sheetOrder.push(sheetId);
    sheets[sheetId] = buildMinimalSheet(sheetId, "Sheet1", {}, 0, 0);
  }

  return {
    id: workbookId,
    name: "workbook",
    locale: LocaleType.ZH_CN,
    sheetOrder,
    appVersion: "3.0.0-alpha",
    styles: {},
    sheets,
  };
}

function buildMinimalSheet(
  sheetId: string,
  name: string,
  cellData: IWorkbookData["sheets"][string]["cellData"],
  rowCount: number,
  columnCount: number
): IWorkbookData["sheets"][string] {
  return {
    id: sheetId,
    name,
    cellData,
    rowCount: Math.max(rowCount, DEFAULT_ROW_COUNT),
    columnCount: Math.max(columnCount, DEFAULT_COLUMN_COUNT),
    hidden: BooleanNumber.FALSE,
    rowHeader: { width: 46, hidden: BooleanNumber.FALSE },
    columnHeader: { height: 20, hidden: BooleanNumber.FALSE },
    rightToLeft: BooleanNumber.FALSE,
  };
}

function worksheetToCellData(
  worksheet: Worksheet
): { cellData: IWorkbookData["sheets"][string]["cellData"]; rowCount: number; columnCount: number } {
  const cellData: IWorkbookData["sheets"][string]["cellData"] = {};
  let maxRow = 0;
  let maxCol = 0;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const r = rowNumber - 1;
    if (!cellData[r]) cellData[r] = {};
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const c = colNumber - 1;
      const type = (cell as { type?: ValueType }).type ?? inferValueType(cell.value);
      const formula = (cell as { formula?: string }).formula;
      const univerCell = excelCellToUniverCell(cell.value, type, formula);
      if (univerCell) {
        (cellData[r] as Record<number, UniverCellLike>)[c] = univerCell;
        maxRow = Math.max(maxRow, r + 1);
        maxCol = Math.max(maxCol, c + 1);
      }
    });
  });

  return {
    cellData,
    rowCount: maxRow || 1,
    columnCount: maxCol || 1,
  };
}

function inferValueType(value: unknown): ValueType {
  if (value === null || value === undefined) return ValueType.Null;
  if (typeof value === "number") return ValueType.Number;
  if (typeof value === "boolean") return ValueType.Boolean;
  if (typeof value === "string") return ValueType.String;
  if (value instanceof Date) return ValueType.Date;
  if (typeof value === "object") {
    if ("formula" in (value as object)) return ValueType.Formula;
    if ("richText" in (value as object)) return ValueType.RichText;
    if ("hyperlink" in (value as object)) return ValueType.Hyperlink;
  }
  return ValueType.String;
}

/** Univer ICellData 最小表示（含 v/t/f/p） */
interface UniverCellExport {
  v?: string | number | boolean;
  t?: CellValueType;
  f?: string;
  p?: IDocumentData;
}

/**
 * IWorkbookData → ExcelJS Workbook（文本/数字/布尔/公式；日期以数字写入）
 */
export function workbookDataToExcelJS(data: IWorkbookData): Workbook {
  const workbook = new ExcelJS.Workbook();
  const order = data.sheetOrder || Object.keys(data.sheets || {});

  for (const sheetId of order) {
    const sheet = data.sheets?.[sheetId];
    if (!sheet) continue;
    const worksheet = workbook.addWorksheet(sheet.name || sheetId, {
      views: [{ state: "normal" }],
    });
    const cellData = sheet.cellData || {};
    const cellDataMap = cellData as Record<string, Record<string, UniverCellExport>>;
    for (const rowKey of Object.keys(cellDataMap)) {
      const rowIndex = Number(rowKey);
      if (Number.isNaN(rowIndex)) continue;
      const row = cellDataMap[rowKey];
      for (const colKey of Object.keys(row)) {
        const colIndex = Number(colKey);
        if (Number.isNaN(colIndex)) continue;
        const cell = row[colKey];
        const excelValue = univerCellToExcelValue(cell);
        if (excelValue !== undefined) {
          worksheet.getRow(rowIndex + 1).getCell(colIndex + 1).value = excelValue;
        }
      }
    }
  }

  return workbook;
}

/**
 * Univer ICellData → ExcelJS cell value（支持公式、布尔、数字、字符串、富文本 p）
 */
function univerCellToExcelValue(cell: UniverCellExport): ExcelJS.CellValue | undefined {
  if (cell.p?.body) {
    const rich = univerDocumentToExcelRichText(cell.p);
    if (rich?.richText?.length) return rich;
  }

  const v = cell?.v;
  if (v === undefined || v === null) return undefined;
  if (v === "" && !cell.f) return undefined;

  if (cell.f && cell.f.length > 0) {
    return { formula: cell.f, result: v } as ExcelJS.CellFormulaValue;
  }

  const t = cell.t;
  if (t === CellValueType.BOOLEAN) return typeof v === "boolean" ? v : Boolean(v);
  if (t === CellValueType.NUMBER) return typeof v === "number" ? v : Number(v);
  if (t === CellValueType.STRING || t === CellValueType.FORCE_STRING) return String(v);
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v;
  return String(v);
}

/** Univer IDocumentData（ICellData.p）→ ExcelJS CellRichTextValue */
function univerDocumentToExcelRichText(doc: IDocumentData): ExcelJS.CellRichTextValue | undefined {
  const body = doc.body;
  if (!body?.dataStream) return undefined;
  const stream = body.dataStream.replace(/\r$/, "");
  if (!stream.length) return { richText: [] };

  const textRuns = body.textRuns ?? [];
  if (textRuns.length === 0) {
    return { richText: [{ text: stream }] };
  }

  const sorted = [...textRuns].sort((a, b) => a.st - b.st);
  const richText: ExcelJS.RichText[] = [];
  let lastEd = 0;
  for (const run of sorted) {
    const st = Math.max(run.st, 0);
    const ed = Math.min(run.ed, stream.length);
    if (st > lastEd) {
      richText.push({ text: stream.slice(lastEd, st) });
    }
    if (ed > st) {
      const font = univerTSToExcelFont(run.ts);
      richText.push(font ? { text: stream.slice(st, ed), font } : { text: stream.slice(st, ed) });
    }
    lastEd = Math.max(lastEd, ed);
  }
  if (lastEd < stream.length) {
    richText.push({ text: stream.slice(lastEd) });
  }
  return { richText };
}

/** Univer ITextStyle → ExcelJS Partial<Font> */
function univerTSToExcelFont(
  ts?: ITextStyle
): Partial<ExcelJS.Font> | undefined {
  if (!ts || Object.keys(ts).length === 0) return undefined;
  const font: Partial<ExcelJS.Font> = {};
  if (ts.ff != null) font.name = ts.ff;
  if (ts.fs != null) font.size = ts.fs;
  if (ts.bl != null) font.bold = ts.bl === BooleanNumber.TRUE;
  if (ts.it != null) font.italic = ts.it === BooleanNumber.TRUE;
  if (ts.ul?.s === BooleanNumber.TRUE) font.underline = "single";
  if (ts.cl?.rgb != null) {
    const rgb = ts.cl.rgb!.replace(/^#/, "");
    font.color = { argb: rgb.length === 6 ? "FF" + rgb : rgb };
  }
  return Object.keys(font).length ? font : undefined;
}
