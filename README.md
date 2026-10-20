# 模板导出插件（excel-template-exporter）

把多维表格（Base）里一条或多条记录，按用户上传的 Excel 模板，把 `{{占位符}}` 或 `='[5]LINK DATA'!D2` 形式的内容替换成对应字段值，导出新的 Excel 附件。

> 本仓库是**插件源代码工程**——需配合 飞书开放平台 的"多维表格自动化操作"能力打包上传，不是直接打开运行的桌面应用。

## 形式与栈

- 多维表格自动化插件框架：`@lark-opdev/block-basekit-server-api`
- TypeScript + Node.js（已在 18.x/20.x 上验证）
- Excel 解析与重写：`exceljs`（主），`xlsx`（兼容保留，可在 `render.ts` 内切换）

## 主要能力

- 三件输入：模板 `.xlsx`（必填）/ 字段映射 JSON（必填）/ 主数据表（必填，autocomplete 动态拉）
- 单条导出（按当前打开记录的 `record_id`）+ 批量导出（整张表，最多 200 条）
- 两种占位符语法同时支持：
  - `{{name}}` 或 `{{name.suffix}}` 的文本占位符
  - `{{arr[0].flight}}` 的数组占位符，用于明细行区块
  - `='[5]LINK DATA'!D2` 这类遗留的外部工作簿引用，自动归一为 `D2` 作为查找键
- 输出文件名可由 `fileNamePattern` 模板控制（如 `{{hawb}}-BOOKING.xlsx`）
- 不计算公式——模板里的 `=TODAY()` 等公式原样保留，交付给 Excel 自算
- 不调用飞书云文档导出接口——不走任何"导出 Bitable 自身"的链路

## 工程结构

```
excel-template-exporter/
├── app.json                  ← 飞书应用配置（需替换为真实 appId）
├── block.json                ← 插件描述（需替换为真实 blockTypeID）
├── package.json              ← 依赖与脚本
├── tsconfig.json             ← 严格模式 TS，commonjs
├── src/
│   ├── register.ts           ← basekit.addAction 入口
│   └── lib/
│       ├── bitable.ts        ← 多维表格 OpenAPI 客户端
│       ├── attachment.ts     ← 附件下载 / uploadAttachments 回传
│       ├── mapping.ts        ← 字段映射 JSON 解析与校验
│       ├── format.ts         ← 按字段类型格式化多维表格值
│       └── render.ts         ← 模板渲染引擎
├── test/
│   └── index.ts              ← 本地烟雾测试（register.ts 加 autocomplete）
├── verify/
│   ├── render-real.ts        ← 直接跑渲染链路（无需飞书 tenant token）
│   ├── output_check.py       ← openpyxl 回读核对输出
│   └── output/               ← 生成产物
└── examples/
    ├── mapping.example.json            ← 配合 BOOKING.xlsx 的映射
    ├── build_placeholder_template.py   ← 把原始 BOOKING 改写为 {{占位符}} 版的脚本
    └── BOOKING-template-placeholder.xlsx  ← 可直接拿来做模板的 {{占位符}} 版 BOOKING
```

## 字段映射 JSON 完整说明

```jsonc
{
  // 必填或留空：占位符名 → 多维表格字段名（按字段名匹配）
  "fields": {
    "hawb": "HAWB",
    "{{hawb}}-BOOKING": "",  // 占位符名也支持 {{...}} 写法
    "D2": "MAWB"             // 外部引用归一后（D2）也按此 lookup
  },

  // 可选：是否启用外部工作簿引用作为占位符
  "externalRefs": true,      // 处理 ='[5]LINK DATA'!D2 这类遗留公式

  // 可选：只导出指定工作表
  "sheet": "BKCF (TRANSIT)", // 缺省导出所有工作表

  // 可选：日期字段的输出格式（exceljs 字段类型 5/1001/1002）
  "dateFormat": "yyyy/MM/dd HH:mm",

  // 可选：复选框 / 多选字段的取值
  "booleanTrue": "是",
  "booleanFalse": "否",
  "arrayJoin": "、",
  "emptyValue": "",

  // 可选：输出文件名模板，{{占位符}} 同样支持
  "fileNamePattern": "{{hawb}}-BOOKING",

  // 可选：明细行区块
  "arrays": [
    {
      "name": "schedule",                   // 占位符前缀
      "sheet": "BKCF (TRANSIT)",
      "startRow": 21,                       // 槽位区间的起始行
      "endRow": 22,                         // 槽位区间的结束行（含）
      "source": {
        "table": "空运表",                  // 从哪个数据表拉明细
        "matchField": "HAWB",               // 与当前主表记录关联的字段
        "matchValue": "{{hawb}}",           // 关联值（可引用当前记录的占位符）
        "maxItems": 10                      // 最多取几行
      },
      "items": {                            // 每行的列名 → 字段名
        "flight": "航班号",
        "date":   "ETD"
      }
    }
  ]
}
```

模板里这样写占位符：

```
A21: {{schedule[0].flight}}     ← 行 21（槽位 0）取明细第 1 条的 flight
B21: {{schedule[0].date}}
A22: {{schedule[1].flight}}     ← 行 22（槽位 1）取明细第 2 条
…
```

当明细记录多于槽位数时，本插件会自动 `duplicateRow` 复制区块，扩展到足够容纳所有明细。

## 本地安装与验证

```bash
cd excel-template-exporter
npm install
npm run build
npm run verify          # 直接跑渲染链路，写入 verify/output/ 并回读 spot-check
npm test                # 烟雾测试（autocomplete + execute 的桩实现）
```

`npm run verify` 会拿仓库内的真实 BOOKING.xlsx + 多维表格 3 条样本记录跑通整套渲染，输出 3 份 `.xlsx` 到 `verify/output/`。

## 上传与发布到飞书

1. 在 飞书开放平台 创建「企业自建应用」并开通「**多维表格 · 自动化操作**」能力。
2. 在权限管理里申请：**查看、评论、编辑和管理多维表格**。
3. 把应用安装到目标文档（REPORT AE → `RPg6bx3ALapBeQs7iwncyRjGn40`）。
4. 拿到 `appId` / `blockTypeID`，填进 `app.json` / `block.json`。
5. 安装 `@lark-opdev/cli`：

```bash
npm i -g @lark-opdev/cli
opdev login
```

6. 在项目根目录里：

```bash
opdev create --appcli-id <你的 appId> \
             --code <blockTypeID> \
             -a bitable-extensions \
             -s automated-action
# 之后每次改动：
npm run upload          # build + opdev upload
npm run preview         # build + opdev preview（生成预览链接，不发布）
```

7. 在多维表格里打开 REPORT AE → 「自动化」标签 → 添加「自动化流」→ 选这个插件 → 给定：
   - 模板文件：可用 `examples/BOOKING-template-placeholder.xlsx`（含 `{{占位符}}`）或 现有 `BOOKING.xlsx`（含 `='[5]LINK DATA'!D2`）
   - 字段映射 JSON：`examples/mapping.example.json`
   - 数据表：`数据表`
   - 导出范围：单条导出（当前记录）/ 批量导出（整张表）

## 已知限制

- **明细行复制不复制合并单元格**：`exceljs` 的 `duplicateRow` 只复制单元格值与样式，不会把同一行的「合并单元格」声明一并复制。若明细区是合并单元格，需手动改成无合并或刷新后重新合并。
- **不计算公式**：模板里的 `=TODAY()` 等公式保持原样输出到新文件；新文件在 Excel 里打开时会自动按本地环境重算。
- **批量上限 200 条**：超出会截断并提示；同时建议先缩小过滤条件再批量。
- **附件下载走 `useTenantAccessToken: true`**：插件部署时必须在「权限」里声明「查看与下载云空间中的文件」，否则模板下载会失败。
- **文件分享回链回到当前任务**：`uploadAttachments` 把生成文件回传到当前文档的 `Drive`，插件下载链接用 7 天有效期；如需长期保存，请在用户侧再人工存档。

## 数据流（一次执行）

```
用户表单 → 
  templateFiles (download via /drive/v1/medias/{token}/download) →
  mappingJson (parseMapping) →
  sourceTable (autocomplete = listTables) →
  records (listRecords; 单条 = 按 record_id 过滤) →
  arrays (listRecords + matchField 过滤) →
  renderTemplate (ExcelJS load → arrays 复制 → scalar 替换 → externalRefs 替换) →
  uploadAttachments (drive/v1/medias/upload_all) →
  { attachments: [Attachment...], summary: String }
```

## 许可

MIT（默认）。
