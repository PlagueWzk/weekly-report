# 实习生周报生成器 (Weekly Report Generator)

一个 Claude Code 技能，用于按标准模板生成实习生周报，输出 **Markdown**（适合 Obsidian）和 **Excel**（用于正式提交）双格式文件。

## 功能概述

- 🗣️ **交互式引导**：分五步引导用户完成周报填写，包括人员信息、日期范围、内容收集、质量自检
- 📝 **Markdown 输出**：带 frontmatter 的 `.md` 文件，可直接导入 Obsidian 知识库
- 📊 **Excel 输出**：基于标准模板自动填写 `.xlsx`，支持多分表（每周一个 sheet），条目行数随内容自动增减（多了插入行、少了删除空白行）
- ✅ **内容质量把关**：内置禁止项检查（禁止空洞表述如"无""正常""顺利"），确保周报言之有物

## 项目结构

```
weekly-report/
├── SKILL.md                          # Claude Code 技能定义（核心指令）
├── assets/
│   └── 实习生周报标准模板.xlsx         # Excel 周报模板
├── scripts/
│   └── fill_weekly_report.py         # Python 脚本：自动填写 xlsx 模板
└── README.md
```

## 前置依赖

- **Claude Code**：本技能运行在 Claude Code 环境中
- **Python 3.8+**（仅 xlsx 生成需要）
- **openpyxl**：Python Excel 库

```bash
pip install openpyxl
# 或
uv pip install openpyxl
```

## 使用方式

在 Claude Code 中，直接说以下任意关键词即可触发：

- "生成周报"
- "写周报"
- "本周周报"
- "实习生周报"
- "weekly report"

然后按 Claude 的引导逐步完成即可。

## 工作流程

```
确认基本信息 → 确认日期 → 收集内容 → 生成 .md → 生成 .xlsx
```

1. **基本信息**：姓名、部门、岗位、指导老师、入职时间、输出目录
2. **日期确认**：自动计算本周周一～周日范围，确认填写日期（周五）
3. **内容收集**：支持口头描述、指定文件、交互式问答三种方式
4. **生成 .md**：按模板生成 Markdown 文件，含 frontmatter 元数据
5. **生成 .xlsx**：通过 Python 脚本自动填写 Excel 模板

## fill_weekly_report.py 用法

如果需要在技能流程外独立使用：

```bash
python scripts/fill_weekly_report.py \
  --template assets/实习生周报标准模板.xlsx \
  --output 输出路径.xlsx \
  --name "姓名" \
  --department "部门" \
  --position "岗位" \
  --mentor "指导老师" \
  --entry-date "2026-08-04" \
  --fill-date "2026-08-14" \
  --sheet "8月第3周" \
  --tasks '[["内容1","完成情况1"],["内容2","完成情况2"]]' \
  --learnings '["收获1","收获2"]' \
  --problems '["问题1"]' \
  --plans '["计划1","计划2","计划3"]'
```

## 周报四大板块

| 板块 | 格式 | 说明 |
|------|------|------|
| 一、本周工作内容 | 表格（序号+内容+完成情况） | 按任务逐条列出 |
| 二、学习与收获 | 编号列表 | 新知识、新技能、新方法 |
| 三、问题与困难 | 编号列表 | 困难、未解决问题、需协调事项 |
| 四、下周工作计划 | 编号列表 | 下周拟开展的主要工作 |

## License

MIT
