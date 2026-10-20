/**
 * test/index.ts —— 在本机直接跑 register.ts 的 execute /
 * formItemAutocomplete，便于在上传到 飞书后端之前先在本地验证一遍。
 *
 * 跑法：
 *   npm run build
 *   node dist/test/index.js
 */

import * as fs from 'fs';
import * as path from 'path';
import { createActionContext } from '@lark-opdev/block-basekit-server-api';
import { AttachmentRef } from '../src/lib/attachment';

const ART = path.resolve(__dirname, '..', '..');
const TPL_PATH = path.join(ART, 'BOOKING.xlsx');

async function runAutocomplete() {
  const ctx = await createActionContext({
    app: { token: 'mock-app', url: 'https://bondexgroup.feishu.cn/wiki/WHATEVER?bitable=RPg6bx3ALapBeQs7iwncyRjGn40' } as any,
    tenantAccessToken: 'mock-token',
    fetch: (() => {
      const stub = (url: string) => {
        if (url.includes('/tables')) {
          return Promise.resolve({
            ok: true,
            status: 200,
            statusText: 'OK',
            json: async () => ({
              code: 0,
              msg: 'ok',
              data: {
                items: [
                  { table_id: 'tbleyaBUxU9WwQr3', name: '数据表' },
                  { table_id: 'tblmezEONRPYxRKn', name: '空运表' },
                ],
                has_more: false,
              },
            }),
            buffer: async () => Buffer.alloc(0),
          } as any);
        }
        return Promise.reject(new Error(`未实现的 mock URL：${url}`));
      };
      return stub;
    })() as any,
  });
  // 直接 require 触发 register.ts 顶层的 addAction 调用
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const reg = require('../src/register');
  void reg;
  const action = (reg as any).default || (require('@lark-opdev/block-basekit-server-api').basekit as any);
  void action;

  // 触发 sourceTable 的 autocomplete
  const { testActionAutocomplete } = require('@lark-opdev/block-basekit-server-api');
  const captured: unknown[] = [];
  const orig = console.log;
  // eslint-disable-next-line no-console
  console.log = (...args: unknown[]) => {
    captured.push(args);
  };
  try {
    await testActionAutocomplete({} as any, ctx, { itemId: 'sourceTable' } as any);
  } finally {
    console.log = orig;
  }
  console.log('autocomplete sourceTable 返回：', JSON.stringify(captured, null, 2));
}

async function runExecuteStub() {
  // 仅做烟雾测试：加载 register.ts，调用 execute(args, ctx) 但拦截 fetch。
  // 真实数据由 verify/render-real.ts 串完，本脚本只证明基座能跑通。
  const ctx = await createActionContext({
    app: { token: 'mock-app', trigger: { url: 'https://bondexgroup.feishu.cn/wiki/MOCK?bitable=RPg6bx3ALapBeQs7iwncyRjGn40' } } as any,
    tenantAccessToken: 'mock-token',
    fetch: (() => {
      return async (url: string) => {
        return {
          ok: true,
          status: 200,
          statusText: 'OK',
          json: async () => ({ code: 0, msg: 'ok', data: {} }),
          buffer: async () => fs.readFileSync(TPL_PATH),
        } as any;
      };
    })() as any,
  });

  // eslint-disable-next-line @typescript-eslint/no-var-requires
  const reg = require('../src/register');
  const basekit = (reg as any).default || require('@lark-opdev/block-basekit-server-api').basekit;
  const executeFn = (basekit as any).action?.execute;
  if (typeof executeFn !== 'function') {
    console.error('未找到已注册的 execute 函数；请检查 src/register.ts 是否调用 basekit.addAction');
    process.exit(1);
  }

  const args = {
    templateFiles: [
      {
        id: 'mock',
        attachmentToken: 'mock-token',
        name: 'BOOKING.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        size: 44152,
        timeStamp: 0,
        extra: '',
      } as AttachmentRef,
    ],
    mappingJson: JSON.stringify({ fields: { D2: 'HAWB' }, externalRefs: true }),
    sourceTable: '数据表',
    exportMode: 'single',
  };

  try {
    const out = await executeFn(args, ctx);
    console.log('execute 返回（summary）：', (out as any).summary);
  } catch (e) {
    console.error('execute 抛错：', (e as Error).message);
    process.exit(2);
  }
}

(async () => {
  console.log('—— 1. autocomplete smoke test ——');
  await runAutocomplete();
  console.log('\n—— 2. execute smoke test (stub fetch) ——');
  await runExecuteStub();
  console.log('\n本机烟雾测试完成。建议在真实部署前再跑一次 npm run verify 渲染真实 BOOKING.xlsx。');
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
