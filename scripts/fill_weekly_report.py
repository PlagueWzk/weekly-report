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
"""

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

# 尝试导入 openpyxl，若不可用则给出安装提示
try:
    import openpyxl
except ImportError:
    print("错误: 需要 openpyxl 库。请运行: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


# ── 模板中的行号常量 ──────────────────────────────────
HEADER_ROW     = 1   # 标题 "实习生周报"
INFO_ROW       = 2   # 姓名 / 部门 / 岗位
MENTOR_ROW     = 3   # 指导老师 / 入职时间 / 填写日期
TASK_HEADER    = 4   # 一、本周工作内容（表头）
TASK_START     = 6   # 任务数据起始行（第 1 条）
TASK_END       = 10  # 任务数据结束行（第 5 条）
LEARN_HEADER   = 11  # 二、学习与收获（表头）
LEARN_START    = 12  # 学习与收获起始行
LEARN_END      = 15  # 学习与收获结束行
PROBLEM_HEADER = 16  # 三、问题与困难（表头）
PROBLEM_START  = 17  # 问题与困难起始行
PROBLEM_END    = 20  # 问题与困难结束行
PLAN_HEADER    = 21  # 四、下周工作计划（表头）
PLAN_START     = 22  # 工作计划起始行
PLAN_END       = 25  # 工作计划结束行


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

    # 填写四个板块
    _fill_tasks(ws, tasks)
    _fill_list_section(ws, LEARN_START, LEARN_END, learnings)
    _fill_list_section(ws, PROBLEM_START, PROBLEM_END, problems)
    _fill_list_section(ws, PLAN_START, PLAN_END, plans)

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
    """清除可能残留的示例占位文本及列表占位符。"""
    # 基本信息中的"示例"占位
    clean_cells = ['B2', 'D2', 'F2', 'B3', 'D3', 'F3']
    for ref in clean_cells:
        val = str(ws[ref].value or '')
        if '示例' in val:
            ws[ref].value = None

    # 列表板块中的 "1、" "2、" 等占位符（防止新建分表时残留）
    for row_num in range(TASK_START, PLAN_END + 1):
        cell = ws.cell(row=row_num, column=1)
        val = str(cell.value or '')
        # 只清除纯粹的占位符（如 "1、"），不误删已填写的实际内容
        if val.strip() in ['1、', '2、', '3、', '4、', '5、']:
            cell.value = None


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


# ── 任务表格 ──────────────────────────────────────────

def _fill_tasks(ws, tasks):
    """填写"一、本周工作内容"表格（第 6-10 行）。

    每行结构：A 列 = 序号, B:C 合并 = 工作内容, D:F 合并 = 完成情况。
    写入合并区域的左上角即可。
    """
    for i, row_num in enumerate(range(TASK_START, TASK_END + 1)):
        if i < len(tasks):
            content, completion = tasks[i]
            ws.cell(row=row_num, column=1, value=i + 1)       # A: 序号
            ws.cell(row=row_num, column=2, value=content)      # B: 工作内容
            ws.cell(row=row_num, column=4, value=completion)   # D: 完成情况
        else:
            # 清除多余行
            for col in (1, 2, 4):
                ws.cell(row=row_num, column=col, value=None)


# ── 编号列表板块（学习与收获 / 问题与困难 / 下周计划）─

def _fill_list_section(ws, start_row, end_row, items):
    """填写"二/三/四"板块（每项占一整行，A:F 合并）。

    自动为每条内容添加 "1、" "2、" 等序号前缀。
    不足模板行数时清除多余行。
    """
    for i, row_num in enumerate(range(start_row, end_row + 1)):
        if i < len(items):
            text = f"{i + 1}、" + items[i]
            ws.cell(row=row_num, column=1, value=text)  # A:F 合并, 写左上角
        else:
            ws.cell(row=row_num, column=1, value='')


if __name__ == "__main__":
    main()
