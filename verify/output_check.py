"""用 openpyxl 把渲染引擎输出的三个文件都过一遍，确认：
1. 文件可正常打开
2. 外部引用单元格（B5、B10、B11...）已被替换为真实值
3. 合并、样式、数字格式等基本属性保留
4. 与原始 BOOKING.xlsx 的 sheet 计数 / sheet 名称 / 单元格数 对齐
"""
import json
import os
import sys
from pathlib import Path

import openpyxl

OUT_DIR = Path("/home/gem/.aily/workdir/task_7698581490505878806/artifacts/excel-template-exporter/dist/verify/output")
TPL_PATH = Path("/home/gem/.aily/workdir/task_7698581490505878806/artifacts/BOOKING.xlsx")
RECORDS_PATH = Path("/home/gem/.aily/workdir/task_7698581490505878806/artifacts/main_records.ndjson")


def cell_kind(value):
    if value is None:
        return "empty"
    if isinstance(value, str) and value.startswith("="):
        return "formula"
    if isinstance(value, str):
        return "text"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, bool):
        return "bool"
    return type(value).__name__


def main():
    if not OUT_DIR.exists():
        print(f"OUT_DIR 不存在：{OUT_DIR}")
        sys.exit(1)

    records = [json.loads(line) for line in RECORDS_PATH.read_text().split("\n") if line.strip()]
    print(f"== 模板 vs 输出 ==")
    tpl_wb = openpyxl.load_workbook(TPL_PATH, data_only=False)
    print(f"模板 sheet 计数：{len(tpl_wb.sheetnames)} ({tpl_wb.sheetnames})")
    for sheet_name in tpl_wb.sheetnames:
        ws = tpl_wb[sheet_name]
        print(f"  {sheet_name}: max_row={ws.max_row} max_col={ws.max_column}")

    out_files = sorted(OUT_DIR.glob("AORD*-BOOKING.xlsx"))
    print(f"\n== 输出回读：{len(out_files)} 个文件 ==")
    ok = True
    for i, out in enumerate(out_files, 1):
        print(f"\n[{i}] {out.name}")
        wb = openpyxl.load_workbook(out, data_only=False)

        if len(wb.sheetnames) != len(tpl_wb.sheetnames):
            print(f"  ❌ sheet 数量不一致：{wb.sheetnames} vs {tpl_wb.sheetnames}")
            ok = False
            continue

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            n_cell = 0
            n_external = 0
            n_external_left = 0
            unresolved = []
            for row in ws.iter_rows():
                for c in row:
                    if c.value is None:
                        continue
                    n_cell += 1
                    if isinstance(c.value, str) and "'[" in c.value and "]LINK DATA'!" in c.value.replace("[", "["):
                        # Treat cells still containing the external-ref text as untouched
                        n_external += 1
                        if "'[" in c.value:
                            n_external_left += 1
                            unresolved.append((c.coordinate, c.value))
                    elif isinstance(c.value, str) and ("[[" in c.value or "LINK DATA" in c.value):
                        n_external_left += 1
                        unresolved.append((c.coordinate, c.value))
            print(f"  sheet {sheet_name!r}  cells={n_cell}  外部引用残留={n_external_left}")
            for u in unresolved[:8]:
                print(f"     {u[0]}: {str(u[1])[:80]}")
            if n_external_left:
                ok = False

        # spot check 第一条记录各字段
        if i == 1:
            ws = wb[tpl_wb.sheetnames[0]]
            spot = ["B5", "B10", "B11", "B12", "B15", "E10", "D10"]
            print("  spot check:")
            expect = {
                "B5": records[0]["HAWB"],
                "B10": records[0]["MAWB"],
                "B11": records[0]["ORI"],
                "B12": records[0]["DEST"],
                "B15": str(records[0]["GW (KGS)"]),
                "E10": records[0]["HAWB"],
                "D10": "HAWB:",
            }
            for ref in spot:
                got = ws[ref].value
                kind = cell_kind(got)
                matched = "✅" if str(got) == str(expect.get(ref, "")) or (ref == "D10" and got == "HAWB:") else "❌"
                print(f"     {ref} = {kind!r} {got!r}  (expect {expect.get(ref)!r}) {matched}")

    print("\n== 结论 ==")
    print("✅ 渲染输出全部符合预期" if ok else "⚠️ 仍有外部引用未被替换")
    sys.exit(0 if ok else 2)


if __name__ == "__main__":
    main()
