/**
 * 多维表格 OpenAPI 封装。
 *
 * 全部请求走 context.fetch（node-fetch 语法）——飞书沙箱禁止 axios / got 等三方库。
 * 鉴权统一用 context.tenantAccessToken（需在 addAction 中开启 useTenantAccessToken）。
 */

const API_BASE = 'https://open.feishu.cn/open-apis/bitable/v1';

/** 自动化插件 Context 的最小子集，便于本地测试时注入桩实现 */
export interface FetchLike {
  (url: string, init?: { method?: string; headers?: Record<string, string>; body?: string }): Promise<{
    ok: boolean;
    status: number;
    statusText: string;
    json(): Promise<any>;
    buffer(): Promise<Buffer>;
  }>;
}

export interface BitContext {
  fetch: FetchLike;
  tenantAccessToken?: string;
  [key: string]: any;
}

export interface TableItem {
  table_id: string;
  name: string;
}

/** 来自多维表格 API 的字段条目；使用宽松字段以适配不同 OpenAPI 版本 */
export interface FieldMeta {
  field_id?: string;
  field_name: string;
  type: number;
  property?: Record<string, any>;
}

export interface BitableRecord {
  record_id: string;
  fields: Record<string, any>;
}

async function callApi<T>(ctx: BitContext, path: string): Promise<T> {
  const token = ctx.tenantAccessToken;
  if (!token) {
    throw new Error('缺少 tenantAccessToken：需在 addAction 中设置 useTenantAccessToken: true');
  }
  const url = `${API_BASE}${path}`;
  const res = await ctx.fetch(url, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json; charset=utf-8',
    },
  });
  if (!res.ok) {
    throw new Error(`OpenAPI 请求失败 HTTP ${res.status} ${res.statusText}：${url}`);
  }
  const body = await res.json();
  if (body && typeof body.code === 'number' && body.code !== 0) {
    throw new Error(`OpenAPI 返回错误 code=${body.code} msg=${body.msg || '(无错误信息)'}：${url}`);
  }
  return body.data as T;
}

/** 列出多维表格下的全部数据表，供「选择数据来源的数据表」下拉使用 */
export async function listTables(ctx: BitContext, appToken: string): Promise<TableItem[]> {
  const items: TableItem[] = [];
  let pageToken = '';
  // 表格数量有限，但仍按接口分页协议完整翻完，避免漏表
  for (let guard = 0; guard < 50; guard += 1) {
    const query = new URLSearchParams({ page_size: '100' });
    if (pageToken) query.set('page_token', pageToken);
    const data = await callApi<{ items?: TableItem[]; has_more?: boolean; page_token?: string }>(
      ctx,
      `/apps/${appToken}/tables?${query.toString()}`,
    );
    for (const it of data.items || []) items.push(it);
    if (!data.has_more || !data.page_token) break;
    pageToken = data.page_token;
  }
  return items;
}

/** 列出数据表字段，用于按字段类型格式化取值 */
export async function listFields(ctx: BitContext, appToken: string, tableId: string): Promise<FieldMeta[]> {
  const items: FieldMeta[] = [];
  let pageToken = '';
  for (let guard = 0; guard < 50; guard += 1) {
    const query = new URLSearchParams({ page_size: '100' });
    if (pageToken) query.set('page_token', pageToken);
    const data = await callApi<{ items?: FieldMeta[]; has_more?: boolean; page_token?: string }>(
      ctx,
      `/apps/${appToken}/tables/${tableId}/fields?${query.toString()}`,
    );
    for (const it of data.items || []) items.push(it);
    if (!data.has_more || !data.page_token) break;
    pageToken = data.page_token;
  }
  return items;
}

export interface ListRecordsOptions {
  pageSize?: number;
  maxRecords?: number;
  /** 传给接口的 filter 表达式（JSON 字符串），留空表示不过滤 */
  filter?: string;
  /** 传给接口的 sort 表达式（JSON 字符串），留空表示不排序 */
  sort?: string;
}

export interface ListRecordsResult {
  records: BitableRecord[];
  hasMore: boolean;
  returned: number;
}

/** 分页拉取记录。达到 maxRecords 即停止，并在结果里如实标记是否被截断 */
export async function listRecords(
  ctx: BitContext,
  appToken: string,
  tableId: string,
  options: ListRecordsOptions = {},
): Promise<ListRecordsResult> {
  const pageSize = Math.min(options.pageSize || 500, 500);
  const maxRecords = options.maxRecords ?? Infinity;
  const records: BitableRecord[] = [];
  let pageToken = '';
  let hasMore = false;

  for (let guard = 0; guard < 100; guard += 1) {
    const query = new URLSearchParams({ page_size: String(pageSize) });
    if (pageToken) query.set('page_token', pageToken);
    if (options.filter) query.set('filter', options.filter);
    if (options.sort) query.set('sort', options.sort);

    const data = await callApi<{
      items?: BitableRecord[];
      has_more?: boolean;
      page_token?: string;
      total?: number;
    }>(ctx, `/apps/${appToken}/tables/${tableId}/records?${query.toString()}`);

    const items = data.items || [];
    for (const item of items) {
      if (records.length >= maxRecords) break;
      records.push(item);
    }

    hasMore = Boolean(data.has_more) && records.length < maxRecords;
    if (!data.has_more || !data.page_token || records.length >= maxRecords) break;
    pageToken = data.page_token as string;
  }

  return { records, hasMore, returned: records.length };
}

/** 按 record_id 取单条记录 */
export async function getRecord(
  ctx: BitContext,
  appToken: string,
  tableId: string,
  recordId: string,
): Promise<BitableRecord> {
  const data = await callApi<{ record: BitableRecord }>(
    ctx,
    `/apps/${appToken}/tables/${tableId}/records/${recordId}`,
  );
  return data.record;
}