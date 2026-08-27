#!/usr/bin/env python3
"""填写实习生周报 xlsx 模板。

用法:
  python fill_weekly_report.py \
    --template assets/实习生周报标准模板.xlsx \
    --output 输出路径.xlsx \
    --name "姓名" --department "部门" --position "岗位" \
    --mentor "指导老师" --entry-date "2026-08-04" --fill-date "2026-08-14" \
    --sheet "8月第3周" \
    --tasks '[["内容1","完成情况1"],["内容2","完成情况2"]]' \
    --learnings '["收获1","收获2"]' \
    --problems '["问题1"]' \
    --plans '["计划1","计划2","计划3"]'

条目行数完全由内容驱动：
- 内容多于模板预留行 → 在板块末尾插入新行（复制上一行的样式与合并单元格）
- 内容少于模板预留行 → 删除板块末尾多余的空白行
"""

import argparse
import json
import shutil
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

# 尝试导入 openpyxl，若不可用则给出安装提示
try:
    import openpyxl
except ImportError:
    print("错误: 需要 openpyxl 库。请运行: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


# ── 板块表头识别文本（按模板中出现的顺序）──────────────
SECTION_HEADERS = [
    ("tasks", "一、本周工作内容"),
    ("learn", "二、学习与收获"),
    ("problem", "三、问题与困难"),
    ("plan", "四、下周工作计划"),
]


def main():
    args = _parse_args()

    # 解析 JSON 数据
    tasks     = json.loads(args.tasks)
    learnings = json.loads(args.learnings)
    problems  = json.loads(args.problems)
    plans     = json.loads(args.plans)

    # 复制模板 → 输出
    shutil.copy2(args.template, args.output)
    wb = openpyxl.load_workbook(args.output)

    # 定位或创建分表
    ws = _get_or_create_sheet(wb, args.sheet)

    # 填写基本信息
    _fill_basic_info(ws, args)

    # 动态定位四个板块的数据区
    sections = _locate_sections(ws)

    # 从后往前调整行数并填写（每个板块的调整只影响其自身数据区之后的行，
    # 因此后处理的板块行号不受影响）
    _resize_and_fill(ws, sections["plan"], plans)
    _resize_and_fill(ws, sections["problem"], problems)
    _resize_and_fill(ws, sections["learn"], learnings)
    _resize_and_fill(ws, sections["tasks"], tasks, kind="tasks")

    wb.save(args.output)
    print(f"✅ 周报已生成: {args.output}")


# ── 参数解析 ──────────────────────────────────────────

def _parse_args():
    p = argparse.ArgumentParser(description="填写实习生周报 xlsx 模板")
    p.add_argument("--template", required=True, help="模板 xlsx 路径")
    p.add_argument("--output", required=True, help="输出 xlsx 路径")
    p.add_argument("--name", required=True, help="姓名 → B2")
    p.add_argument("--department", default="", help="部门（末级组织）→ D2")
    p.add_argument("--position", default="", help="岗位 → F2")
    p.add_argument("--mentor", default="", help="指导老师 → B3")
    p.add_argument("--entry-date", default="", help="入职时间 → D3 (YYYY-MM-DD)")
    p.add_argument("--fill-date", default="", help="填写日期 → F3 (YYYY-MM-DD)")
    p.add_argument("--sheet", required=True, help='分表名称，如 "8月第3周"')
    p.add_argument("--tasks", default="[]",
                   help='JSON 数组，每项为 [工作内容, 完成情况]')
    p.add_argument("--learnings", default="[]",
                   help='JSON 字符串数组，不含序号前缀')
    p.add_argument("--problems", default="[]",
                   help='JSON 字符串数组，不含序号前缀')
    p.add_argument("--plans", default="[]",
                   help='JSON 字符串数组，不含序号前缀')
    return p.parse_args()


# ── 分表管理 ──────────────────────────────────────────

def _get_or_create_sheet(wb, sheet_name):
    """返回目标分表；若不存在则从干净模板复制新建。"""
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]

    # 选择最干净的 sheet 作为复制源（优先选无预填数据的）
    source_sheet = _pick_cleanest_sheet(wb)
    new_ws = wb.copy_worksheet(source_sheet)
    new_ws.title = sheet_name

    # 清理复制源可能带过来的示例占位数据
    _clean_example_data(new_ws)

    return new_ws


def _pick_cleanest_sheet(wb):
    """在模板的所有分表中选数据最'干净'的作为复制源。
    判断标准：B2（姓名）为空或为示例文本 且 B3（指导老师）为空或为示例。
    """
    best = wb[wb.sheetnames[0]]
    for name in wb.sheetnames:
        ws = wb[name]
        b2 = str(ws['B2'].value or '')
        b3 = str(ws['B3'].value or '')
        # 完全空白的最优先
        if b2 == '' and b3 == '':
            return ws
        # 含"示例"的次之
        if '示例' in b2 or '示例' in b3:
            best = ws
    return best


def _clean_example_data(ws):
    """清除复制源可能残留的示例占位文本（基本信息区）。

    列表板块中的 "1、" "2、" 占位符无需在此清理：填写时会覆盖，
    多余行会被删除（见 _resize_and_fill）。
    """
    clean_cells = ['B2', 'D2', 'F2', 'B3', 'D3', 'F3']
    for ref in clean_cells:
        val = str(ws[ref].value or '')
        if '示例' in val:
            ws[ref].value = None


# ── 板块定位 ──────────────────────────────────────────

def _locate_sections(ws):
    """扫描 A 列，按表头文本动态定位四个板块的数据区。

    返回 {key: {"start": row, "end": row}}，数据区不含表头行本身：
    - 前三个板块以相邻板块的表头行为边界
    - 任务板块额外跳过列头行（"序号/工作内容/完成情况"）
    - 最后一个板块以"说明"文字行（或 max_row）为边界
    """
    header_rows = {}
    for row in range(1, ws.max_row + 1):
        val = str(ws.cell(row=row, column=1).value or '')
        for key, text in SECTION_HEADERS:
            if text in val and key not in header_rows:
                header_rows[key] = row

    missing = [text for key, text in SECTION_HEADERS if key not in header_rows]
    if missing:
        sys.exit(f"错误: 模板中未找到板块表头 {missing}，请确认模板结构")

    # 最后一个板块之后的边界：说明文字起始行
    last_header = header_rows[SECTION_HEADERS[-1][0]]
    tail = ws.max_row + 1
    for row in range(last_header + 1, ws.max_row + 1):
        val = str(ws.cell(row=row, column=1).value or '')
        if val and '说明' in val:
            tail = row
            break

    sections = {}
    for i, (key, _) in enumerate(SECTION_HEADERS):
        start = header_rows[key] + 1
        if i + 1 < len(SECTION_HEADERS):
            end = header_rows[SECTION_HEADERS[i + 1][0]] - 1
        else:
            end = tail - 1
        if key == "tasks":
            # 跳过"序号/工作内容/完成情况"列头行
            col_header = str(ws.cell(row=start, column=1).value or '')
            if '序号' in col_header:
                start += 1
        sections[key] = {"start": start, "end": end}
    return sections


# ── 行数调整 ──────────────────────────────────────────

def _resize_and_fill(ws, section, items, kind="list"):
    """把板块数据区调整为恰好 len(items) 行并填写内容。

    - 内容多于预留行 → 在数据区末尾插入新行（复制样式与合并单元格）
    - 内容少于预留行 → 删除数据区末尾多余的空白行
    """
    start, end = section["start"], section["end"]
    needed = len(items)
    current = end - start + 1

    if current < 1:
        sys.exit(f"错误: 板块数据区为空（第 {start} 行起），请确认模板结构")

    if needed > current:
        _insert_rows(ws, end + 1, needed - current)
    elif needed < current:
        _delete_rows(ws, start + needed, current - needed)

    # 填写内容（数据区起始行不变，始终写入前 needed 行）
    for i in range(needed):
        row = start + i
        if kind == "tasks":
            content, completion = items[i]
            ws.cell(row=row, column=1, value=i + 1)      # A: 序号
            ws.cell(row=row, column=2, value=content)    # B: 工作内容
            ws.cell(row=row, column=4, value=completion) # D: 完成情况
        else:
            ws.cell(row=row, column=1, value=f"{i + 1}、{items[i]}")


def _insert_rows(ws, row_idx, count):
    """在 row_idx 处插入 count 行。

    1. 将 row_idx 及之后的合并区域整体下移（跨插入点的区域扩大）
    2. 新行复制 row_idx-1 行的样式、行高与单行合并区域
    """
    template_row = row_idx - 1

    # 收集全部合并区域（clear 后需要完整重建）
    unchanged = []  # 插入点之前，保持不动 (min_col, min_row, max_col, max_row)
    shifted = []    # 插入点及之后，整体下移
    copied = []     # 上一行的单行合并列范围 (min_col, max_col)
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        if min_row >= row_idx:
            shifted.append((min_col, min_row + count, max_col, max_row + count))
        elif max_row >= row_idx:
            # 合并区域跨插入点 → 向下扩大
            shifted.append((min_col, min_row, max_col, max_row + count))
        else:
            unchanged.append((min_col, min_row, max_col, max_row))
        if min_row == template_row and max_row == template_row:
            copied.append((min_col, max_col))

    ws.merged_cells.ranges.clear()
    ws.insert_rows(row_idx, count)
    for min_col, min_row, max_col, max_row in unchanged + shifted:
        ws.merge_cells(start_row=min_row, start_column=min_col,
                       end_row=max_row, end_column=max_col)

    # 新行复制样式、行高与合并
    for i in range(count):
        r = row_idx + i
        for col in range(1, ws.max_column + 1):
            dst = ws.cell(row=r, column=col)
            dst._style = copy(ws.cell(row=template_row, column=col)._style)
        if ws.row_dimensions[template_row].height:
            ws.row_dimensions[r].height = ws.row_dimensions[template_row].height
        for min_col, max_col in copied:
            ws.merge_cells(start_row=r, start_column=min_col,
                           end_row=r, end_column=max_col)


def _delete_rows(ws, first_row, count):
    """删除 first_row 起的 count 行，并同步维护合并区域。

    模板中所有合并区域均为单行合并，被删的空白条目行对应整行区域，
    因此只处理"完全在删除区内（丢弃）"与"完全在删除区后（上移）"两种情况。
    """
    last_row = first_row + count - 1
    kept = []
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        if min_row >= first_row and max_row <= last_row:
            continue  # 完全位于删除区内 → 丢弃
        if min_row > last_row:
            min_row -= count
            max_row -= count
        elif max_row >= first_row:
            # 起点在删除区上方、终点落入删除区 → 截断到删除区上方
            max_row = first_row - 1
            if max_row < min_row:
                continue
        kept.append((min_col, min_row, max_col, max_row))

    ws.merged_cells.ranges.clear()
    ws.delete_rows(first_row, count)
    for min_col, min_row, max_col, max_row in kept:
        ws.merge_cells(start_row=min_row, start_column=min_col,
                       end_row=max_row, end_column=max_col)


# ── 基本信息 ──────────────────────────────────────────

def _fill_basic_info(ws, args):
    """填写姓名 / 部门 / 岗位 / 指导老师 / 入职时间 / 填写日期。"""
    ws['B2'] = args.name
    ws['D2'] = args.department
    ws['F2'] = args.position
    ws['B3'] = args.mentor

    # 日期字段：尝试解析为 datetime，保留 yyyy/m/d 格式
    _set_date_cell(ws, 'D3', args.entry_date)
    _set_date_cell(ws, 'F3', args.fill_date)


def _set_date_cell(ws, cell_ref, date_str):
    """将单元格设为日期值（datetime）并应用 yyyy/m/d 数字格式。"""
    if not date_str:
        return
    try:
        dt = datetime.strptime(date_str, '%Y-%m-%d')
        ws[cell_ref] = dt
        ws[cell_ref].number_format = 'yyyy/m/d'
    except ValueError:
        ws[cell_ref] = date_str  # 无法解析则原样写入


if __name__ == "__main__":
    main()
