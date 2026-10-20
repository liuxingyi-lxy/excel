/**
 * 字段映射配置解析与校验。
 *
 * 配置由用户在插件表单里以 JSON 形式填入，插件本身不含任何业务字段名——
 * 占位符与多维表格字段的对应关系完全由这份配置决定。
 *
 * 配置结构（所有键都可省略，除 fields 外建议至少给一个）：
 * {
 *   "fields":    { "<占位符名>": "<多维表格字段名>" },
 *   "externalRefs": true,                  // 把 ='[5]LINK DATA'!D2 这类外部工作簿引用也当作占位符处理
 *   "sheet":     "工作表名",                 // 留空表示导出模板中所有工作表
 *   "dateFormat":"yyyy/MM/dd HH:mm",       // 日期字段的输出格式
 *   "booleanTrue": "是", "booleanFalse": "否",
 *   "arrayJoin": "、",                      // 多选 / 关联字段的拼接分隔符
 *   "emptyValue": "",                      // 空值统一填成什么
 *   "fileNamePattern": "{{hawb}}-BOOKING",  // 输出文件名，缺省用模板名 + 记录标识
 *   "arrays": [
 *     {
 *       "name": "schedule",
 *       "sheet": "BKCF (TRANSIT)",         // 缺省跟随全局 sheet
 *       "startRow": 21, "endRow": 22,      // 明细区间的行号（含两端）
 *       "source": {
 *         "table": "数据表",
 *         "matchField": "HAWB",
 *         "matchValue": "{{hawb}}",        // 支持引用当前记录的占位符
 *         "maxItems": 10
 *       },
 *       "items": { "flight": "航班号", "date": "ETD" }
 *     }
 *   ]
 * }
 */

export interface ArraySourceConfig {
  table: string;
  matchField?: string;
  /** 固定值，或用 {{占位符}} 引用当前记录的值 */
  matchValue?: string;
  filter?: string;
  sort?: string;
  maxItems?: number;
}

export interface ArrayConfig {
  name: string;
  sheet?: string;
  startRow: number;
  endRow: number;
  source: ArraySourceConfig;
  items: Record<string, string>;
}

export interface MappingConfig {
  fields: Record<string, string>;
  /** 是否启用外部工作簿引用（如 ='[5]LINK DATA'!D2）作为占位符来替换 */
  externalRefs: boolean;
  sheet?: string;
  dateFormat: string;
  booleanTrue: string;
  booleanFalse: string;
  arrayJoin: string;
  emptyValue: string;
  fileNamePattern?: string;
  arrays: ArrayConfig[];
}

export const DEFAULT_MAPPING: MappingConfig = {
  fields: {},
  externalRefs: false,
  dateFormat: 'yyyy/MM/dd HH:mm',
  booleanTrue: '是',
  booleanFalse: '否',
  arrayJoin: '、',
  emptyValue: '',
  arrays: [],
};

class MappingError extends Error {}

function fail(message: string): never {
  throw new MappingError(`映射配置有误：${message}`);
}

function asString(value: unknown, where: string): string {
  if (typeof value === 'string') return value;
  if (typeof value === 'number') return String(value);
  return fail(`${where} 必须是字符串，当前是 ${typeof value}`);
}

function asOptionalString(value: unknown, where: string): string | undefined {
  if (value === undefined || value === null || value === '') return undefined;
  return asString(value, where);
}

function asPositiveInt(value: unknown, where: string): number {
  const n = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(n) || !Number.isInteger(n) || n <= 0) {
    return fail(`${where} 必须是正整数，当前是 ${JSON.stringify(value)}`);
  }
  return n;
}

function parseArrays(raw: unknown): ArrayConfig[] {
  if (raw === undefined || raw === null) return [];
  if (!Array.isArray(raw)) return fail('arrays 必须是数组');

  return raw.map((item: any, index: number) => {
    const at = `arrays[${index}]`;
    if (!item || typeof item !== 'object') return fail(`${at} 必须是对象`);

    const name = asString(item.name, `${at}.name`);
    if (!name) return fail(`${at}.name 不能为空`);

    const source = item.source;
    if (!source || typeof source !== 'object') return fail(`${at}.source 必须是对象`);
    const table = asString(source.table, `${at}.source.table`);
    if (!table) return fail(`${at}.source.table 不能为空`);

    if (!item.items || typeof item.items !== 'object' || Array.isArray(item.items)) {
      return fail(`${at}.items 必须是「键名 -> 多维表格字段名」的对象`);
    }
    const items: Record<string, string> = {};
    for (const [key, value] of Object.entries(item.items)) {
      items[key] = asString(value, `${at}.items.${key}`);
    }

    const startRow = asPositiveInt(item.startRow, `${at}.startRow`);
    const endRow = asPositiveInt(item.endRow, `${at}.endRow`);
    if (endRow < startRow) return fail(`${at}.endRow 不能小于 startRow`);

    return {
      name,
      sheet: asOptionalString(item.sheet, `${at}.sheet`),
      startRow,
      endRow,
      source: {
        table,
        matchField: asOptionalString(source.matchField, `${at}.source.matchField`),
        matchValue: asOptionalString(source.matchValue, `${at}.source.matchValue`),
        filter: asOptionalString(source.filter, `${at}.source.filter`),
        sort: asOptionalString(source.sort, `${at}.source.sort`),
        maxItems: source.maxItems === undefined ? undefined : asPositiveInt(source.maxItems, `${at}.source.maxItems`),
      },
      items,
    };
  });
}

export function parseMapping(input: string): MappingConfig {
  const rawText = (input || '').trim();
  if (!rawText) return { ...DEFAULT_MAPPING, fields: {} };

  let parsed: any;
  try {
    parsed = JSON.parse(rawText);
  } catch (e) {
    fail(`不是合法 JSON（${(e as Error).message}）`);
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    fail('顶层必须是一个 JSON 对象');
  }

  const fields: Record<string, string> = {};
  if (parsed.fields !== undefined) {
    if (!parsed.fields || typeof parsed.fields !== 'object' || Array.isArray(parsed.fields)) {
      fail('fields 必须是「占位符名 -> 多维表格字段名」的对象');
    }
    for (const [key, value] of Object.entries(parsed.fields)) {
      fields[key] = asString(value, `fields.${key}`);
    }
  }

  const externalRefs = parsed.externalRefs === undefined ? DEFAULT_MAPPING.externalRefs : Boolean(parsed.externalRefs);

  return {
    fields,
    externalRefs,
    sheet: asOptionalString(parsed.sheet, 'sheet'),
    dateFormat: asOptionalString(parsed.dateFormat, 'dateFormat') || DEFAULT_MAPPING.dateFormat,
    booleanTrue: asOptionalString(parsed.booleanTrue, 'booleanTrue') ?? DEFAULT_MAPPING.booleanTrue,
    booleanFalse: asOptionalString(parsed.booleanFalse, 'booleanFalse') ?? DEFAULT_MAPPING.booleanFalse,
    arrayJoin: parsed.arrayJoin === undefined ? DEFAULT_MAPPING.arrayJoin : asString(parsed.arrayJoin, 'arrayJoin'),
    emptyValue: parsed.emptyValue === undefined ? DEFAULT_MAPPING.emptyValue : String(parsed.emptyValue ?? ''),
    fileNamePattern: asOptionalString(parsed.fileNamePattern, 'fileNamePattern'),
    arrays: parseArrays(parsed.arrays),
  };
}

export { MappingError };