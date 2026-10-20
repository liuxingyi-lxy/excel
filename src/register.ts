/**
 * 模板导出插件 —— basekit.addAction 入口文件。
 *
 * 三件输入：
 *   1) templateFiles  模板 .xlsx 附件
 *   2) mappingJson    字段映射 JSON
 *   3) sourceTable    数据表选择（autocomplete 动态拉取当前多维表格下的全部数据表）
 *
 * 一件选择：exportMode —— 单条导出（按当前打开记录的 record_id）或 批量导出（拉整张表）。
 */

import {
  basekit,
  Component,
  ParamType,
  StructureType,
} from '@lark-opdev/block-basekit-server-api';
import { downloadAttachment, uploadFiles, type AttachmentRef } from './lib/attachment';
import { listFields, listRecords, listTables, type BitContext, type FieldMeta, type BitableRecord } from './lib/bitable';
import { formatValue, indexFields } from './lib/format';
import { parseMapping, MappingError, type ArrayConfig, type MappingConfig } from './lib/mapping';
import { renderTemplate, buildFileName } from './lib/render';

const { SingleSelect, Input, Attachment } = Component;

const MAX_BATCH_RECORDS = 200;

function getBitContext(_args: any, context: any): BitContext {
  const url = context?.app?.trigger?.url || context?.app?.url || '';
  const m = url.match(/bitable\/([A-Za-z0-9_-]+)/);
  const appToken = m ? m[1] : '';
  if (!appToken) {
    throw new Error('无法识别当前多维表格 app_token，请在多维表格中打开此插件');
  }
  return {
    ...context,
    appToken,
    tableId: '',
    recordId: '',
    triggerRecordId: context?.app?.trigger?.recordID,
  } as BitContext;
}

interface ArraySource {
  table: string;
  matchField?: string;
  matchValue?: string;
  sort?: string;
  maxItems?: number;
}

/**
 * 从明细数据表里把与当前记录关联的若干条记录拉出来。
 * 这里假定明细表的 matchField 形如 "<当前表字段名>.<明细表字段名>" 时一进来就拿当前记录的字段值；
 * 也可以是单个字段名。简单实现：手动按字段名对当前记录求值，再传给 listRecords 的 filter。
 */
async function loadArrayItems(
  ctx: BitContext,
  tableId: string,
  source: ArraySource,
  recordFields: Record<string, any>,
): Promise<BitableRecord[]> {
  if (!source.table || !tableId) return [];
  const result = await listRecords(ctx, ctx.appToken, tableId, {} as any);
  let items = result.records;
  if (source.matchField) {
    const want = recordFields ? recordFields[source.matchField] : undefined;
    items = items.filter((r) => {
      const v = r.fields ? r.fields[source.matchField!] : undefined;
      if (Array.isArray(v)) return v.includes(want);
      if (want === undefined || want === null) return v === undefined || v === null;
      return String(v) === String(want);
    });
  }
  if (source.maxItems && source.maxItems > 0) {
    items = items.slice(0, source.maxItems);
  }
  return items;
}

/** 把 mapping.fields 里 {{占位符名}} 替换为当前记录字段值；用于 matchValue 这类动态条件 */
function renderMatchString(pattern: string, recordFields: Record<string, any>, fieldsIndex: Map<string, FieldMeta>, mapping: MappingConfig): string {
  return pattern.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_: string, key: string) => {
    const fieldName = mapping.fields[key.trim()];
    if (!fieldName) return '';
    const field = fieldsIndex.get(fieldName);
    if (!field) return '';
    const raw = recordFields ? recordFields[fieldName] : undefined;
    const v = formatValue(raw, field, mapping);
    return v ?? '';
  });
}

// —— formItems ——

const formItems = [
  {
    key: 'templateFiles',
    label: '模板文件（.xlsx）',
    component: Attachment,
    componentProps: {
      mode: 'single',
      fileType: { accept: { xlsx: ['.xlsx', '.XLSX'] } },
    },
    required: true,
    tooltip: '上传一份含占位符（{{name}} / {{arr[0].key}} / 外部工作簿引用）的 Excel 模板',
  } as any,
  {
    key: 'mappingJson',
    label: '字段映射 JSON',
    component: Input,
    componentProps: {
      mode: 'textarea',
      rows: 14,
      placeholder:
        '{\n  "fields": { "hawb": "HAWB", "D2": "MAWB" },\n  "externalRefs": true,\n  "arrays": [ { "name":"schedule", "startRow":21, "endRow":22, "source": {"table":"空运表","matchField":"HAWB"}, "items": {"flight":"航班号","date":"ETD"} } ]\n}',
    },
    required: true,
    tooltip:
      'fields：{占位符 -> 多维表格字段名}；externalRefs=true 把 BOOKING 那类外部工作簿引用也参与替换；arrays：明细行区块（多个航班等多行）。',
  } as any,
  {
    key: 'sourceTable',
    label: '数据表',
    component: SingleSelect,
    componentProps: {
      placeholder: '请选择要把哪些记录导出',
      options: [],
      autocomplete: true,
    },
    required: true,
    tooltip: '选择当前文档中的某个数据表作为「主表」（待导出的每一行记录就从这里来）',
  } as any,
  {
    key: 'exportMode',
    label: '导出范围',
    component: SingleSelect,
    componentProps: {
      options: [
        { label: '单条导出（当前记录）', value: 'single' },
        { label: '批量导出（整张表，最多 200 条）', value: 'batch' },
      ],
      defaultValue: 'single',
    },
    required: true,
    tooltip: '单条 = 当前打开的那一行；批量 = 把数据表里所有记录都生成一份',
  } as any,
];

basekit.addAction({
  formItems: formItems as any,
  formItemAutocomplete: {
    sourceTable: async (_args: any, context: any) => {
      try {
        const ctx = getBitContext(_args, context);
        const tables = await listTables(ctx, ctx.appToken);
        return tables.map((t) => ({ label: t.name, value: t.table_id }));
      } catch (e) {
        return [{ label: `无法读取数据表：${(e as Error).message}`, value: '', disabled: true }];
      }
    },
  } as any,
  execute: async (args: any, context: any): Promise<any> => {
    try {
      const ctx = getBitContext(args, context);

      // 1) 模板
      const tplRaw: any = args.templateFiles;
      const templateFile: AttachmentRef | undefined = Array.isArray(tplRaw) ? tplRaw[0] : tplRaw;
      if (!templateFile || !templateFile.attachmentToken) {
        throw new Error('请先上传模板文件');
      }
      const templateBuffer = await downloadAttachment(templateFile, context);

      // 2) 映射
      let mapping: MappingConfig;
      try {
        mapping = parseMapping(typeof args.mappingJson === 'string' ? args.mappingJson : '');
      } catch (e) {
        if (e instanceof MappingError) throw new Error(e.message);
        throw e;
      }

      // 3) 数据表
      const tables = await listTables(ctx, ctx.appToken);
      const tablesById: Record<string, string> = {};
      const tablesByName: Record<string, string> = {};
      for (const t of tables) {
        tablesById[t.table_id] = t.name;
        tablesByName[t.name] = t.table_id;
      }
      const wanted = args.sourceTable;
      if (!wanted) throw new Error('请选择数据表');
      const tableId = tablesById[wanted] ? wanted : tablesByName[wanted];
      if (!tableId) throw new Error(`未找到数据表「${wanted}」`);
      ctx.tableId = tableId;
      const tableName = tablesById[tableId];

      // 4) 主表字段索引
      const mainFields = await listFields(ctx, ctx.appToken, tableId);
      const fieldsIndex = indexFields(mainFields);

      // 5) 各明细表的字段索引
      const arrayTableIds: Record<string, string> = {};
      for (const arr of mapping.arrays) {
        const tid = tablesByName[arr.source.table];
        if (!tid) {
          throw new Error(`明细数据表「${arr.source.table}」不存在；请检查映射 JSON 里 arrays[].source.table`);
        }
        arrayTableIds[arr.source.table] = tid;
      }
      const arrayFieldsByTable: Record<string, { list: FieldMeta[]; index: Map<string, FieldMeta> }> = {};
      for (const [name, tid] of Object.entries(arrayTableIds)) {
        const list = await listFields(ctx, ctx.appToken, tid);
        arrayFieldsByTable[name] = { list, index: indexFields(list) };
      }

      // 6) 取主表记录
      const exportMode: 'single' | 'batch' = args.exportMode === 'batch' ? 'batch' : 'single';
      let records: BitableRecord[];
      if (exportMode === 'single') {
        const recordId: string | undefined = ctx.triggerRecordId;
        if (!recordId) {
          throw new Error('请先在数据表里打开一条具体记录，再用「单条导出」模式运行');
        }
        const result = await listRecords(ctx, ctx.appToken, tableId, {} as any);
        records = result.records.filter((r: any) => r.record_id === recordId);
        if (!records.length) throw new Error(`未找到记录 ${recordId}`);
      } else {
        const all = await listRecords(ctx, ctx.appToken, tableId, {} as any);
        records = all.records.slice(0, MAX_BATCH_RECORDS);
      }

      // 7) 逐条渲染
      const buffers: Buffer[] = [];
      const names: string[] = [];
      const unresolvedAll: Set<string> = new Set();
      const failed: string[] = [];
      const truncated: string[] = [];

      for (let index = 0; index < records.length; index += 1) {
        const record = records[index];
        const recordFields = (record as any).fields || {};

        // 收集每条明细记录（可能引用主表字段动态条件）
        const arrays: Record<string, BitableRecord[]> = {};
        for (const arr of mapping.arrays as ArrayConfig[]) {
          const tid = arrayTableIds[arr.source.table];
          if (!tid) continue;
          const meta = arrayFieldsByTable[arr.source.table];
          if (!meta) continue;

          let items: BitableRecord[] = await listRecords(ctx, ctx.appToken, tid, {} as any).then(
            (r) => r.records,
          );

          // 按当前主表字段过滤
          if (arr.source.matchField) {
            const want = (() => {
              if (arr.source.matchValue !== undefined) {
                const v = renderMatchString(arr.source.matchValue, recordFields, fieldsIndex, mapping);
                return v;
              }
              const raw = recordFields[arr.source.matchField];
              if (raw === undefined || raw === null) return raw;
              const field = fieldsIndex.get(arr.source.matchField);
              if (!field) return String(raw);
              return formatValue(raw, field, mapping);
            })();
            items = items.filter((r: any) => {
              const v = r.fields ? r.fields[arr.source.matchField!] : undefined;
              if (Array.isArray(v)) return v.includes(want);
              if (want === undefined || want === null || want === '') return v === undefined || v === null || v === '';
              return String(v) === String(want);
            });
          }
          if (arr.source.maxItems && arr.source.maxItems > 0) {
            items = items.slice(0, arr.source.maxItems);
          }
          arrays[arr.name] = items;

          // 把当前记录的字段注入明细字段索引：明细里若有占位符键名等于主表字段名，依然能解析
          // （简化：让两套字段索引都可用：占位符键名 -> 明细字段名 -> 主表 fallback）
          for (const [key, fName] of Object.entries(arr.items)) {
            const f = meta.index.get(fName);
            if (f) {
              fieldsIndex.set(fName, f);
            } else if (fieldsIndex.get(fName) === undefined) {
              // 留作占位以供 renderTemplate 内部解析
            }
          }
        }

        try {
          const out = await renderTemplate({
            templateBuffer,
            record,
            fieldsIndex,
            mapping,
            arrays,
          });
          const name = buildFileName(mapping, record, fieldsIndex, templateFile.name || 'template.xlsx', index);
          buffers.push(out.buffer);
          names.push(name);
          for (const u of out.unresolved) unresolvedAll.add(u);
        } catch (e) {
          failed.push(`第 ${index + 1} 条：${(e as Error).message}`);
        }
      }

      if (exportMode === 'batch' && records.length >= MAX_BATCH_RECORDS) {
        truncated.push(`已达批量上限 ${MAX_BATCH_RECORDS} 条`);
      }

      if (buffers.length === 0) {
        throw new Error('没有任何记录被渲染；' + (failed.length ? `失败原因：${failed.join('；')}` : '请检查数据表与映射配置'));
      }

      const files = buffers.map((buf, i) => ({
        name: names[i],
        file: buf,
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      } as any));

      const uploaded = await uploadFiles(files, context);

      const summary =
        `数据表：${tableName}（渲染 ${records.length} 条）` +
        `\n模式：${exportMode === 'single' ? '单条导出' : '批量导出'}` +
        `\n生成文件：${files.length} 个` +
        (unresolvedAll.size ? `\n未替换占位符（${[...unresolvedAll].slice(0, 20).join('、')}${unresolvedAll.size > 20 ? ' 等' : ''}）` : '') +
        (failed.length ? `\n失败：${failed.join('；')}` : '') +
        (truncated.length ? `\n备注：${truncated.join('；')}` : '');

      return {
        attachments: uploaded,
        summary,
      };
    } catch (err) {
      const message = (err as Error).message || '执行失败';
      return { attachments: [], summary: `导出失败：${message}` };
    }
  },
  resultType: {
    type: ParamType.Object,
    properties: {
      attachments: {
        type: ParamType.Array,
        structure: { type: StructureType.Attachment, field: 'attachments' },
      },
      summary: { type: ParamType.String },
    },
  },
  useTenantAccessToken: true,
  permission: { type: 2 } as any,
  nextSteps: [
    { key: 'AddRecord', label: '把渲染结果作为新记录写入当前表' },
    { key: 'SetRecord', label: '把渲染结果覆盖到当前记录' },
  ],
} as any);

