/**
 * 模板渲染引擎。
 *
 * 三件事：
 *  1. 展开明细行区块（arrays）——按数组长度复制模板行，支持 {{arr[0].x}} 形式的槽位占位符
 *  2. 替换标量占位符 {{name}}
 *  3. 把 ='[5]LINK DATA'!D2 这类外部工作簿引用替换为真实值（externalRefs 开启时）
 *
 * 不做任何公式计算：模板里的 =TODAY() 等公式原样保留，交给 Excel 自己算。
 */

import * as ExcelJS from 'exceljs';
import type { BitableRecord, FieldMeta } from './bitable';
import { formatValue } from './format';
import type { ArrayConfig, MappingConfig } from './mapping';

/** ='[5]LINK DATA'!D2 / ='[4]AE JAN'!$AB$10 —— 外部工作簿引用（`=` 可缺省，便于匹配公式文本） */
const EXTERNAL_REF = /(?:=\s*)?'\[[^\]]*\]([^']*)'\s*!\s*(\$?[A-Z]{1,3}\$?\d+)/g;
/** {{name}} —— 标量占位符；name 允许中文、下划线、点、连字符 */
const SCALAR_PLACEHOLDER = /\{\{\s*([^{}]+?)\s*\}\}/g;
/** {{arr[0].key}} —— 数组槽位占位符 */
const ARRAY_PLACEHOLDER = /\{\{\s*([A-Za-z_一-龥][\w一-龥.-]*?)\s*\[\s*(\d+)\s*\]\s*\.\s*([\w一-龥.-]+)\s*\}\}/g;

export interface RenderInput {
  templateBuffer: Buffer;
  record: BitableRecord;
  fieldsIndex: Map<string, FieldMeta>;
  mapping: MappingConfig;
  /** 预先取好的明细数据：数组名 -> 记录列表 */
  arrays: Record<string, BitableRecord[]>;
}

export interface RenderOutput {
  buffer: Buffer;
  /** 渲染后仍未被替换掉的占位符，用于如实告知用户哪些没配 */
  unresolved: string[];
  sheets: string[];
}

/** 取单元格的「可替换文本」：普通字符串或富文本；公式单元格返回 null（公式不参与标量替换） */
function readCellText(cell: ExcelJS.Cell): string | null {
  const v = cell.value as any;
  if (typeof v === 'string') return v;
  if (v && typeof v === 'object' && Array.isArray(v.richText)) {
    return v.richText.map((part: any) => (typeof part === 'string' ? part : part?.text || '')).join('');
  }
  return null;
}

/** 取单元格的公式文本（含 sharedFormula 展开后的形式），非公式返回 null */
function readCellFormula(cell: ExcelJS.Cell): string | null {
  const v = cell.value as any;
  if (v && typeof v === 'object' && typeof v.formula === 'string') return v.formula;
  return null;
}

function writeCellText(cell: ExcelJS.Cell, text: string): void {
  cell.value = text;
}

/** 去掉 $ 符号并大写，便于把 D2 / $d$2 归一 */
function normalizeRef(ref: string): string {
  return ref.replace(/\$/g, '').toUpperCase();
}

/**
 * 构造标量占位符解析器：占位符名 -> 单元格文本。
 * 映射里没配的占位符返回 undefined，由调用方决定如何汇报。
 */
export function createScalarResolver(
  record: BitableRecord,
  fieldsIndex: Map<string, FieldMeta>,
  mapping: MappingConfig,
): (placeholder: string) => string | undefined {
  return (placeholder: string): string | undefined => {
    const key = placeholder.trim();
    const fieldName = mapping.fields[key];
    if (fieldName === undefined) return undefined;
    const field = fieldsIndex.get(fieldName);
    if (!field) return undefined;
    const raw = record.fields ? record.fields[fieldName] : undefined;
    return formatValue(raw, field, mapping);
  };
}

/** 用给定解析器替换文本里的所有标量占位符；解析不到的保持原样 */
function applyScalar(text: string, resolve: (key: string) => string | undefined, unresolved: Set<string>): string {
  return text.replace(SCALAR_PLACEHOLDER, (whole, key: string) => {
    const value = resolve(key);
    if (value === undefined) {
      unresolved.add(key);
      return whole;
    }
    return value;
  });
}

/** 把文本里的外部工作簿引用替换成真实值：把单元格地址（如 D2）作为映射键 */
function applyExternalRefs(
  text: string,
  resolve: (key: string) => string | undefined,
  mapping: MappingConfig,
  unresolved: Set<string>,
): string {
  if (!mapping.externalRefs) return text;
  return text.replace(EXTERNAL_REF, (whole, _sheet: string, ref: string) => {
    const key = normalizeRef(ref);
    const value = resolve(key);
    if (value === undefined) {
      unresolved.add(key);
      return whole;
    }
    return value;
  });
}

/**
 * 展开一个明细行区块。
 *
 * 语义：模板中 [startRow, endRow] 是一个「槽位块」，块内 k = endRow - startRow + 1 行。
 * 数组有 n 项时，块被复制 ceil(n / k) 份，第 c 份的第 s 个槽位对应数组第 c*k + s 项。
 * 因此模板里写 {{arr[0].x}} / {{arr[1].x}} 即可，块会随数组长度自动扩展。
 */
function expandArrayRegion(
  ws: ExcelJS.Worksheet,
  config: ArrayConfig,
  items: BitableRecord[],
  fieldsIndex: Map<string, FieldMeta>,
  mapping: MappingConfig,
  unresolved: Set<string>,
): void {
  const startRow = config.startRow;
  const endRow = config.endRow;
  const slotCount = endRow - startRow + 1;

  // 行号超出模板实际范围时直接跳过，避免把整张表撑坏
  if (startRow < 1 || endRow > ws.rowCount) return;

  const copies = Math.max(1, Math.ceil(items.length / slotCount));
  if (copies > 1) {
    // ExcelJS 的 duplicateRow 只复制单元格模型，不复制合并单元格；明细区一般不合并，故可接受
    ws.duplicateRow(endRow, copies - 1, true);
  }

  const itemFieldsByKey: Record<string, FieldMeta | undefined> = {};
  for (const [key, fieldName] of Object.entries(config.items)) {
    itemFieldsByKey[key] = fieldsIndex.get(fieldName);
  }

  for (let copy = 0; copy < copies; copy += 1) {
    for (let slot = 0; slot < slotCount; slot += 1) {
      const rowNumber = startRow + copy * slotCount + slot;
      const row = ws.getRow(rowNumber);
      if (!row) continue;
      const item = items[copy * slotCount + slot];

      row.eachCell({ includeEmpty: false }, (cell) => {
        const formula = readCellFormula(cell);
        if (formula !== null) {
          // 明细区块内如果写了外部引用，一并替换
          const replaced = applyExternalRefs(formula, resolveArrayKey, mapping, unresolved);
          if (replaced !== formula) writeCellText(cell, replaced);
          return;
        }
        const text = readCellText(cell);
        if (text === null) return;
        if (!text.includes('{{')) return;

        const replaced = text.replace(ARRAY_PLACEHOLDER, (whole, name: string, idxText: string, key: string) => {
          if (name !== config.name) return whole;
          const slotIndex = Number(idxText);
          if (slotIndex !== slot) return whole;
          if (!item) return '';
          const field = itemFieldsByKey[key];
          const raw = item.fields ? item.fields[config.items[key]] : undefined;
          return formatValue(raw, field, mapping);
        });
        if (replaced !== text) writeCellText(cell, replaced);
      });

      function resolveArrayKey(key: string): string | undefined {
        if (!item) return undefined;
        const fieldName = config.items[key];
        if (fieldName === undefined) return undefined;
        const field = itemFieldsByKey[key];
        const raw = item.fields ? item.fields[fieldName] : undefined;
        return formatValue(raw, field, mapping);
      }
    }
  }
}

/** 收集渲染过程中没能替换掉的占位符（区分数组槽位与标量两类） */
function scanUnresolved(ws: ExcelJS.Worksheet, set: Set<string>): void {
  ws.eachRow((row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const formula = readCellFormula(cell);
      const text = formula ?? readCellText(cell);
      if (!text) return;
      let match: RegExpExecArray | null;
      const re = new RegExp(ARRAY_PLACEHOLDER.source, 'g');
      while ((match = re.exec(text)) !== null) set.add(match[0]);
      const re2 = new RegExp(SCALAR_PLACEHOLDER.source, 'g');
      while ((match = re2.exec(text)) !== null) set.add(match[0]);
    });
  });
}

export async function renderTemplate(input: RenderInput): Promise<RenderOutput> {
  const { templateBuffer, record, fieldsIndex, mapping, arrays } = input;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer as unknown as ExcelJS.Buffer);

  const targetSheets: ExcelJS.Worksheet[] = mapping.sheet
    ? [workbook.getWorksheet(mapping.sheet)].filter((ws): ws is ExcelJS.Worksheet => Boolean(ws))
    : workbook.worksheets.filter((ws): ws is ExcelJS.Worksheet => Boolean(ws));
  if (targetSheets.length === 0) {
    throw new Error(mapping.sheet ? `模板里找不到名为「${mapping.sheet}」的工作表` : '模板里没有任何工作表');
  }

  const resolve = createScalarResolver(record, fieldsIndex, mapping);
  const unresolved = new Set<string>();

  // 先展开明细区块（按起始行倒序，避免前面的展开影响后面区间的行号），再做全表标量替换
  const arrayConfigs = [...mapping.arrays].sort((a, b) => b.startRow - a.startRow);
  for (const config of arrayConfigs) {
    const items = arrays[config.name] || [];
    const fallbackSheet = config.sheet || mapping.sheet || '';
    let ws: ExcelJS.Worksheet | undefined = fallbackSheet ? workbook.getWorksheet(fallbackSheet) : undefined;
    if (!ws) ws = targetSheets[0];
    if (!ws) {
      unresolved.add(`arrays.${config.name}(缺少工作表 ${config.sheet || mapping.sheet || '默认'})`);
      continue;
    }
    expandArrayRegion(ws, config, items, fieldsIndex, mapping, unresolved);
  }

  for (const ws of targetSheets) {
    ws.eachRow((row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const formula = readCellFormula(cell);
        if (formula !== null) {
          const replaced = applyExternalRefs(formula, resolve, mapping, unresolved);
          if (replaced !== formula) writeCellText(cell, replaced);
          return;
        }
        const text = readCellText(cell);
        if (text === null || !text.includes('{{')) return;
        const afterScalar = applyScalar(text, resolve, unresolved);
        const withRefs = applyExternalRefs(afterScalar, resolve, mapping, unresolved);
        if (withRefs !== text) writeCellText(cell, withRefs);
      });
    });
  }

  for (const ws of targetSheets) scanUnresolved(ws, unresolved);

  const buffer = (await workbook.xlsx.writeBuffer()) as unknown as Buffer;
  return {
    buffer: Buffer.from(buffer),
    unresolved: [...unresolved],
    sheets: targetSheets.map((ws) => ws.name),
  };
}

/** 依据映射生成输出文件名，去掉文件系统不允许的字符 */
export function buildFileName(
  mapping: MappingConfig,
  record: BitableRecord,
  fieldsIndex: Map<string, FieldMeta>,
  templateName: string,
  index: number,
): string {
  const strip = (s: string) => s.replace(/[\\/:*?"<>|]/g, '_').trim();
  const baseName = (templateName || 'export').replace(/\.[xX][lL][sS][xX]$/, '') || 'export';

  if (mapping.fileNamePattern) {
    const resolve = createScalarResolver(record, fieldsIndex, mapping);
    const rendered = mapping.fileNamePattern.replace(SCALAR_PLACEHOLDER, (whole, key: string) => {
      const value = resolve(key);
      if (value === undefined) return whole;
      return value;
    });
    // 未被替换的占位符会留下 {{...}}，直接清掉避免文件名非法
    const cleaned = strip(rendered.replace(/\{\{[^{}]*\}\}/g, ''));
    if (cleaned) return `${cleaned}.xlsx`;
  }

  return `${strip(baseName)}-${String(index + 1).padStart(4, '0')}.xlsx`;
}