/**
 * 按多维表格字段类型把单元格值格式化成写入 Excel 的文本。
 *
 * 字段类型编号取自多维表格开放接口；未识别的类型走通用兜底序列化，
 * 保证任何情况下都能产出可读文本，而不是抛错或写入 [object Object]。
 */

import type { FieldMeta } from './bitable';
import type { MappingConfig } from './mapping';

const FIELD_TYPE = {
  TEXT: 1,
  NUMBER: 2,
  SINGLE_SELECT: 3,
  MULTI_SELECT: 4,
  DATE: 5,
  CHECKBOX: 7,
  USER: 11,
  PHONE: 13,
  URL: 15,
  ATTACHMENT: 17,
  LINK: 18,
  LOOKUP: 19,
  FORMULA: 20,
  DOUBLE_LINK: 21,
  LOCATION: 22,
  GROUP_CHAT: 23,
  CREATED_TIME: 1001,
  MODIFIED_TIME: 1002,
  CREATED_BY: 1003,
  MODIFIED_BY: 1004,
  AUTO_NUMBER: 1005,
} as const;

/** 支持的日期格式占位符：yyyy YY MM M dd D HH H mm M ss S，其余字符原样保留 */
function formatTimestamp(ms: number, pattern: string): string {
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return '';
  const pad = (n: number, len = 2) => String(n).padStart(len, '0');
  const map: Record<string, string> = {
    yyyy: String(d.getFullYear()),
    YY: pad(d.getFullYear() % 100),
    MM: pad(d.getMonth() + 1),
    dd: pad(d.getDate()),
    HH: pad(d.getHours()),
    mm: pad(d.getMinutes()),
    ss: pad(d.getSeconds()),
    S: pad(d.getMilliseconds(), 3),
  };
  // 单字母形式（不区分大小写地补零）：M -> 月、D -> 日、H -> 时
  map.M = pad(d.getMonth() + 1);
  map.D = pad(d.getDate());
  map.H = pad(d.getHours());
  return pattern.replace(/yyyy|YY|MM|dd|HH|mm|ss|S|M|D|H/g, (token) => map[token] ?? token);
}

function isEmptyValue(value: any): boolean {
  if (value === null || value === undefined || value === '') return true;
  if (Array.isArray(value)) return value.length === 0;
  return false;
}

/** 多维表格的人员 / 创建人字段只保证有 id，name 视接口返回而定 */
function personToText(person: any): string {
  if (person === null || person === undefined) return '';
  if (typeof person === 'string') return person;
  if (typeof person !== 'object') return String(person);
  return person.name || person.en_name || person.email || person.id || '';
}

function attachmentToText(item: any): string {
  if (typeof item === 'string') return item;
  if (!item || typeof item !== 'object') return '';
  return item.name || item.url || item.token || '';
}

/** 通用兜底：尽力把任意值转成可读文本 */
function stringifyUnknown(value: any, join: string): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (Array.isArray(value)) return value.map((v) => stringifyUnknown(v, join)).filter((s) => s !== '').join(join);
  if (typeof value === 'object') {
    const obj = value as Record<string, any>;
    if (typeof obj.text === 'string') return obj.text; // 富文本
    if (typeof obj.name === 'string') return obj.name;
    if (typeof obj.value !== 'undefined') return stringifyUnknown(obj.value, join);
    try {
      return JSON.stringify(value);
    } catch {
      return '';
    }
  }
  return String(value);
}

export function formatValue(value: any, field: FieldMeta | undefined, mapping: MappingConfig): string {
  if (isEmptyValue(value)) return mapping.emptyValue;

  const type = field?.type;

  switch (type) {
    case FIELD_TYPE.DATE:
    case FIELD_TYPE.CREATED_TIME:
    case FIELD_TYPE.MODIFIED_TIME: {
      // 多维表格日期字段返回毫秒时间戳；个别场景可能是格式化好的字符串，原样透传
      if (typeof value === 'number') return formatTimestamp(value, mapping.dateFormat);
      if (typeof value === 'string' && /^\d+$/.test(value)) return formatTimestamp(Number(value), mapping.dateFormat);
      return typeof value === 'string' ? value : stringifyUnknown(value, mapping.arrayJoin);
    }

    case FIELD_TYPE.NUMBER:
      // 数字不做千分位等二次加工，保持接口返回的精度，避免和源数据对不上
      return typeof value === 'number' ? String(value) : stringifyUnknown(value, mapping.arrayJoin);

    case FIELD_TYPE.CHECKBOX:
      return value ? mapping.booleanTrue : mapping.booleanFalse;

    case FIELD_TYPE.SINGLE_SELECT:
    case FIELD_TYPE.MULTI_SELECT:
      return (Array.isArray(value) ? value : [value])
        .map((v) => stringifyUnknown(v, mapping.arrayJoin))
        .filter((s) => s !== '')
        .join(mapping.arrayJoin);

    case FIELD_TYPE.USER:
    case FIELD_TYPE.CREATED_BY:
    case FIELD_TYPE.MODIFIED_BY:
      return (Array.isArray(value) ? value : [value])
        .map(personToText)
        .filter((s) => s !== '')
        .join(mapping.arrayJoin);

    case FIELD_TYPE.ATTACHMENT:
      return (Array.isArray(value) ? value : [value])
        .map(attachmentToText)
        .filter((s) => s !== '')
        .join(mapping.arrayJoin);

    case FIELD_TYPE.LOCATION:
      return stringifyUnknown(value, mapping.arrayJoin);

    case FIELD_TYPE.GROUP_CHAT:
      return (Array.isArray(value) ? value : [value])
        .map((g: any) => (typeof g === 'object' && g ? g.name || g.id : g))
        .filter((s: any) => s !== '')
        .join(mapping.arrayJoin);

    case FIELD_TYPE.LINK:
    case FIELD_TYPE.DOUBLE_LINK:
      // 关联字段只有 record_id，这里原样输出，便于人工回查
      return (Array.isArray(value) ? value : [value])
        .map((v: any) => (typeof v === 'object' && v ? v.record_id || v.id || v.text || v.name : v))
        .filter((s: any) => s !== '')
        .join(mapping.arrayJoin);

    default:
      return stringifyUnknown(value, mapping.arrayJoin);
  }
}

/** 把字段名 -> 显示名 建索引，records.fields 的 key 就是字段名 */
export function indexFields(fields: FieldMeta[]): Map<string, FieldMeta> {
  const map = new Map<string, FieldMeta>();
  for (const f of fields) map.set(f.field_name, f);
  return map;
}