# 实习生周报生成器 (Weekly Report Generator)

一个 Claude Code 技能，用于按标准模板生成实习生周报，输出 **Markdown**（适合 Obsidian）和 **Excel**（用于正式提交）双格式文件。

## 功能概述

- 👤 **个人信息档案**：姓名/部门/岗位等保存在技能目录的 `profile.json`（已 gitignore），首次填写后每周自动沿用
- 📅 **日期智能默认**：周二~周五默认填本周，周一/周末默认补写上周（周报时限为周五 20:00），展示确认后可随时调整
- 🔗 **上周计划联动**：自动定位上一周周报的"下周工作计划"，与本周素材对照收集；相违和时询问以哪个为准
- 📝 **Markdown 输出**：带 frontmatter 的 `.md` 文件，可直接导入 Obsidian 知识库
- 📊 **Excel 输出**：`.md` 交给脚本一键生成 `.xlsx`（`--from-md`），支持多分表（每周一个 sheet）、条目行数随内容自动增减、行高随内容自适应
- ✅ **内容质量把关**：内置禁止项检查（禁止空洞表述如"无""正常""顺利"），确保周报言之有物

## 项目结构

```
weekly-report/
├── SKILL.md                          # Claude Code 技能定义（核心指令）
├── profile.json                      # 个人信息档案（首次运行生成，已 gitignore）
├── assets/
│   └── 实习生周报标准模板.xlsx         # Excel 周报模板（单一"模板"分表，干净结构）
├── scripts/
│   └── fill_weekly_report.py         # Python 脚本：解析周报 md 并填写 xlsx 模板
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
- "本周周报" / "上周周报"
- "实习生周报"
- "weekly report"

然后按 Claude 的引导逐步完成即可。

## 工作流程

```
读取/收集基本信息 → 确认日期与分表 → 对照上周计划收集内容 → 生成 .md → 生成 .xlsx
```

1. **基本信息**：首次询问姓名、部门、岗位、指导老师、入职时间、输出目录并保存到 `profile.json`；之后自动沿用
2. **日期确认**：按星期自动给出默认周报范围（周二~周五 → 本周；周一/周末 → 上周），并建议分表名
3. **内容收集**：先对照上周计划，再按口头描述、指定文件、交互式问答三种方式收集
4. **生成 .md**：按模板生成 Markdown 文件，含 frontmatter 元数据
5. **生成 .xlsx**：`.md` 直接交给脚本解析填写 Excel 模板

## fill_weekly_report.py 用法

推荐：直接解析技能生成的周报 md（Windows cmd 下无 JSON 引号转义问题）：

```bash
python scripts/fill_weekly_report.py \
  --template assets/实习生周报标准模板.xlsx \
  --from-md "输出目录/姓名工作周报(20260831-20260906)-指导老师.md" \
  --sheet "9月第1周" \
  --output "输出目录/姓名工作周报(20260831-20260906)-指导老师.xlsx"
```

脚本自动解析 md 的信息区与四个板块，并完成：分表定位/新建、基本信息填写、条目行数增减（插入行复制样式与合并单元格）、行高自适应、日期写入为 `yyyy/m/d`。

备用：命令行直接传 JSON（bash 引号写法，个别字段覆盖 md 时也可追加）：

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

完整参数说明：`python scripts/fill_weekly_report.py --help`

## 周报四大板块

| 板块 | 格式 | 说明 |
|------|------|------|
| 一、本周工作内容 | 表格（序号+内容+完成情况） | 按任务逐条列出 |
| 二、学习与收获 | 编号列表 | 新知识、新技能、新方法 |
| 三、问题与困难 | 编号列表 | 困难、未解决问题、需协调事项 |
| 四、下周工作计划 | 编号列表 | 下周拟开展的主要工作 |

## 更新与分发

本仓库是技能的唯一源码处，更新流程：修改并提交 → 推送 GitHub → cc-switch 统一拉取分发到各工具的技能目录。安装与链接由 cc-switch 管理，本仓库不包含任何安装逻辑。

## License

MIT
