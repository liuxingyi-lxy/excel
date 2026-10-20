/**
 * 附件下载与回传。
 *
 * - 下载：用户上传的模板走 drive/v1/medias/{attachmentToken}/download，需文档权限（permission.type = 2）
 * - 回传：SDK 提供的 uploadAttachments 走 drive/v1/medias/upload_all，结果可直接作为 resultType 的附件数组
 */

import { uploadAttachments } from '@lark-opdev/block-basekit-server-api';
import type { FetchLike } from './bitable';

export interface AttachmentRef {
  attachmentToken: string;
  extra: string;
  id: string;
  mimeType: string;
  name: string;
  size: number | string;
  timeStamp: number | string;
}

export interface DownloadCtx {
  fetch: FetchLike;
  tenantAccessToken?: string;
}

/** 下载用户上传的模板文件，返回二进制内容 */
export async function downloadAttachment(attachment: AttachmentRef, context: DownloadCtx): Promise<Buffer> {
  if (!context.tenantAccessToken) {
    throw new Error('缺少 tenantAccessToken：需在 addAction 中设置 useTenantAccessToken: true');
  }
  const query = new URLSearchParams({ extra: attachment.extra });
  const url = `https://open.feishu.cn/open-apis/drive/v1/medias/${attachment.attachmentToken}/download?${query.toString()}`;
  const res = await context.fetch(url, {
    method: 'GET',
    headers: { Authorization: `Bearer ${context.tenantAccessToken}` },
  });
  if (!res.ok) {
    throw new Error(`模板下载失败 HTTP ${res.status} ${res.statusText}`);
  }
  return Buffer.from(await res.buffer());
}

export interface OutgoingFile {
  name: string;
  file: Buffer;
}

/**
 * 把生成的文件上传为素材，返回可直接写入 resultType 的附件数组。
 * env 固定为 'feishu'，因为插件只在飞书多维表格沙箱内运行。
 */
export async function uploadFiles(files: OutgoingFile[], context: any): Promise<AttachmentRef[]> {
  const attachments = await uploadAttachments(files, { context, env: 'feishu' });
  return (attachments || []) as AttachmentRef[];
}