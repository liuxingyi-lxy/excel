/**
 * 渲染链路验证脚本：在沙箱里把 BOOKING.xlsx + main_records.ndjson 跑一遍，
 * 输出一份填充好的 .xlsx，再用 openpyxl 回读核对。
 *
 * 不用走 basekit 框架；直接调用 renderTemplate + 序列化报告。
 */

import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import { indexFields } from '../src/lib/format';
import { renderTemplate, buildFileName } from '../src/lib/render';
import type { MappingConfig } from '../src/lib/mapping';
import type { BitableRecord, FieldMeta } from '../src/lib/bitable';

const ART = path.resolve(__dirname, '..', '..', '..');
const TPL_PATH = path.join(ART, 'BOOKING.xlsx');
const FIELDS_PATH = path.join(ART, 'fields_main.json');
const RECORDS_PATH = path.join(ART, 'main_records.ndjson');
const AIR_FIELDS_PATH = path.join(ART, 'fields_air.json');
const OUT_DIR = path.join(__dirname, 'output');

function loadJson(p: string): any {
  return JSON.parse(fs.readFileSync(p, 'utf-8'));
}

/** 接口返回会被外面包一层 {ok, identity, data: {...}}，这里剥到真正的字段数组上 */
function unwrapFields(raw: any): FieldMeta[] {
  let arr = raw;
  if (raw && Array.isArray(raw.data?.items)) arr = raw.data.items;
  else if (raw && Array.isArray(raw.data?.fields)) arr = raw.data.fields;
  else if (raw && Array.isArray(raw.fields)) arr = raw.fields;
  else if (raw && Array.isArray(raw.items)) arr = raw.items;
  if (!Array.isArray(arr)) {
    throw new Error('字段 JSON 既不是数组也找不到 data.items / data.fields');
  }
  return arr as FieldMeta[];
}

/** 把 basekit 的字段 shape (id/name/type) 转成 render.ts 期望的 {field_id, field_name, type} */
function toFieldMeta(arr: any[]): FieldMeta[] {
  return arr.map((f: any) => ({
    field_id: f.id ?? f.field_id,
    field_name: f.name ?? f.field_name,
    type: typeof f.type === 'string' ? typeNameToNumber(f.type) : f.type,
    property: f.property ?? f.style ?? f.options,
  }));
}

function typeNameToNumber(name: string): number {
  const m: Record<string, number> = {
    text: 1,
    number: 2,
    select: 3,
    multi_select: 4,
    date: 5,
    datetime: 5,
    checkbox: 7,
    user: 11,
    phone: 13,
    url: 15,
    attachment: 17,
    link: 18,
    lookup: 19,
    formula: 20,
    double_link: 21,
    location: 22,
    group_chat: 23,
    created_time: 1001,
    modified_time: 1002,
    created_by: 1003,
    modified_by: 1004,
    auto_number: 1005,
  };
  return m[name] ?? 1;
}

function loadNdjson(p: string): any[] {
  const text = fs.readFileSync(p, 'utf-8');
  return text.split('\n').filter(Boolean).map((line) => JSON.parse(line));
}

function ensureDir(p: string): void {
  if (!fs.existsSync(p)) fs.mkdirSync(p, { recursive: true });
}

async function main(): Promise<void> {
  ensureDir(OUT_DIR);

  const templateBuffer = fs.readFileSync(TPL_PATH);
  const mainFields: FieldMeta[] = toFieldMeta(unwrapFields(loadJson(FIELDS_PATH)));
  const airFields: FieldMeta[] = toFieldMeta(unwrapFields(loadJson(AIR_FIELDS_PATH)));
  void airFields;
  const recordsFromNdjson = loadNdjson(RECORDS_PATH).map((r, i) => ({
    record_id: r.record_id || `rec_${i}`,
    fields: r,
  }));

  // 占位符（以公式里出现的单元格地址或 template 里的 {{name}} 作为 key）→ 多维表格字段名
  // BOOKING.xlsx 的公式长这样：="[5]LINK DATA'!D2，渲染器把 D2 归一后做 key 查询
  const mapping: MappingConfig = {
    fields: {
      // 标量占位符示例（如果模板里有 {{hawb}} 这种字面文本，可这样配）
      hawb: 'HAWB',
      mawb: 'MAWB',
      ori: 'ORI',
      dest: 'DEST',
      commodity: 'COMMODITY',
      qty: 'CTNS ON HAWB',
      gw: 'GW (KGS)',
      cw: 'CW (KGS)',
      remark: 'SHIPPER',
      cut_off_time: 'CUT OFF TIME',
      warehouse: 'WAREHOUSE',
      // 外部引用：用单元地址作 key
      D2: 'HAWB',     // B5  : TO 名（实际放 HAWB 行对齐查看者）
      H2: 'MAWB',     // B10 : MAWB
      I2: 'HAWB',     // E10 : HAWB
      J2: 'ORI',      // B11 : ORIGIN
      K2: 'DEST',     // B12 : DEST
      AD2: 'COMMODITY',
      W2: 'CTNS ON HAWB',
      Y2: 'GW (KGS)',
      G2: 'SHIPPER',  // B17 : REMARK (取发货人做默认备注)
      AF2: 'CUT OFF TIME',
      AG2: 'WAREHOUSE',
      // SCHEDULE 多行的外部引用
      L2: 'ORI',
      M2: 'FLIGHT 1 NO',
      N2: 'FLIGHT 1 DATE',
      O2: 'CTNS ON HAWB',
      P2: 'CUT OFF TIME',
      Q2: 'GW (KGS)',
      R2: 'DEST',
      S2: 'FLIGHT 2 NO',
      T2: 'FLIGHT 2 DATE',
      U2: 'CTNS ON HAWB',
      V2: 'CUT OFF TIME',
    },
    externalRefs: true,
    dateFormat: 'yyyy/MM/dd HH:mm',
    booleanTrue: '是',
    booleanFalse: '否',
    arrayJoin: '、',
    emptyValue: '',
    fileNamePattern: '{{hawb}}-BOOKING',
    arrays: [],
  };

  const fieldsIndex = indexFields(mainFields);

  console.log(`准备渲染 ${recordsFromNdjson.length} 条记录`);
  const results: any[] = [];
  for (let index = 0; index < recordsFromNdjson.length; index += 1) {
    const record = recordsFromNdjson[index] as BitableRecord;
    const out = await renderTemplate({
      templateBuffer,
      record,
      fieldsIndex,
      mapping,
      arrays: {},
    });
    const name = buildFileName(mapping, record, fieldsIndex, 'BOOKING.xlsx', index);
    const target = path.join(OUT_DIR, name);
    fs.writeFileSync(target, out.buffer);
    console.log(`  [${index + 1}] ${name}  unresolved=${out.unresolved.length}`);
    for (const u of out.unresolved) console.log(`     - ${u}`);
    results.push({ name, target, unresolved: out.unresolved, sheets: out.sheets });
  }

  fs.writeFileSync(
    path.join(OUT_DIR, 'verify-report.json'),
    JSON.stringify({ results, mapping: { ...mapping, fields: Object.keys(mapping.fields) } }, null, 2),
  );

  // 抽样回读第一个输出
  if (results.length) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(results[0].target);
    const sheetName = wb.worksheets[0].name;
    const ws = wb.worksheets[0];
    const cells: Record<string, any> = {};
    for (const ref of ['B5', 'B10', 'B11', 'B12', 'B13', 'B14', 'B15', 'B17', 'D10', 'E10']) {
      const cell = ws.getCell(ref);
      cells[ref] = { value: cell.value, text: cell.text };
    }
    console.log('\n第一个文件抽样回读：');
    console.log(JSON.stringify(cells, null, 2));
  }
}

main().catch((e) => {
  console.error('verify 失败：', e);
  process.exit(1);
});
