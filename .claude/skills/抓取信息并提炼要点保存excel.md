---
name: 抓取信息并提炼要点保存excel
description: 抓取互联网最新信息，交叉核实来源真实性（官方 vs 媒体报道），提炼结构化要点（时间/主题/简介/内容/来源URL等字段，可扩展视频脚本），并导出为 Excel 保存到本地。当用户说"抓取/搜集/收集最新XX信息""从网上找XX资料""提炼要点""整理成Excel""保存到结果文件夹/桌面"时使用。
---

# 抓取信息并提炼要点保存 Excel

把"从互联网抓信息 → 核实真实性 → 提炼结构化要点 → 导出 Excel 到本地"做成一条流水线。
目标是：**给用户可直接拿去用的信息表，且每条信息都能溯源、分得清"官方确认"和"媒体报道"。**

## ⚖️ 核心原则（先读这个）

1. **真实性优先，诚实汇报**：能确认就说"已确认"，不能确认就说"待核实"，绝不把媒体推测当官方事实写进表格。
2. **官方来源第一**：优先找公司官方渠道（官方博客/官方文档/官方 GitHub/官方 Hugging Face/官方产品站），媒体只做交叉佐证。
3. **✓/△ 分开标注**：每个事实标注属于「官方已确认（✓）」还是「媒体报道·待核实（△）」，两列分开放，不混写。
4. **给用户可自行访问的 URL**：Excel 里放完整可点击的链接，让用户能亲自验证。
5. **先讲清自己的核实边界**：如果当前环境打不开某些域名（如 openai.com 被网络策略拦截），如实说明"URL 来自搜索引擎索引 + 多源一致转述"，不要假装访问过。

---

## 📥 流程 Step 1：抓取信息

1. **并行搜索**：同时发多个 WebSearch，覆盖不同角度，避免只信一条：
   - 主题关键词 + "最新/发布/上线"
   - 主题 + 具体主体名（如公司/产品名）
   - 中文源 + 英文源各一组
2. **优先取官方域名**：搜索结果里筛选官方来源（`api-docs.xxx.com`、`openai.com`、`github.com`、`huggingface.co`、公司官网等）单独摘出来。
3. **交叉验证数字**：同一组关键数字（金额、参数、时间、分数）要在 ≥2 个独立来源里一致才写实。
4. **记录日期**：以用户当前日期为基准，每条新闻标注具体时间（如 `2026-07-31`），别只写"近期"。

---

## 🔍 流程 Step 2：提炼要点

按用途决定字段集。**核心字段（必填）**：

| 字段 | 说明 |
|---|---|
| 标题 | 每条信息的一句话标题 |
| 时间 | 发布日期，精确到日 |
| 主题 | 所属类别（如"大模型发布/融资/开源/工具"） |
| 简介 | 2-3 句概述，说清"是什么事" |
| 核心内容要点 | 编号列出的关键事实（参数/金额/时间线/数据），每条独立成行 |
| 官方已确认事实（✓） | 有官方来源背书的事实 |
| 媒体报道·待核实（△） | 媒体转述/推测/未经官方回应的事项 |
| 官方来源URL | 可直接访问的官方链接 |
| 权威媒体来源URL | 上证报/每经/36氪/IT之家等佐证链接 |
| 信息真实性评级 | 高 / 中高 / 中 / 待核实，一句话说明理由 |

**增强字段（按需补充，可扩展）**：

| 字段 | 说明 |
|---|---|
| 视频脚本 | 完整口播文案（若用户是视频创作者，按"钩子-展开-收尾"结构给） |
| 分镜脚本 | `时间段 \| 画面 \| 台词` 表格 |
| 钩子方向 | 推荐的开场切入角度 |
| 目标受众视角 | "这跟谁有什么关系"（如打工人视角） |
| 带得走的行动 | 用户看完能立刻做的事 |
| 制作/使用备注 | 红线提醒（如"融资数字需说'据报道'"） |

> 若用户给了背景资料（如 Obsidian 里的账号定位），提炼字段要贴合他的视角，别写通用套话。

---

## ✅ 流程 Step 3：核实真实性

对每条关键事实做分类：

- **✓ 官方已确认**：官方博客原文、官方文档、官方技术报告、官方开源仓库、官方产品站。来源写完整 URL。
- **△ 媒体报道**：权威媒体（财经类/科技类大媒体）报道但公司官方未回应，或 The Information 这类外媒独家转述。**必须标注"据媒体报道/未经官方确认"**，不能写成既定事实。
- **⚠ 待核实**：单个小号/自媒体信源、数字前后矛盾、无法溯源。宁可删掉或标注，不写进"已确认"列。

给每条信息一个**真实性评级**：`高 / 中高 / 中 / 低 / 待核实`，并在"核实状态说明"列写一句话理由（核了几家、官方是否回应、哪些不能打开）。

---

## 📊 流程 Step 4：导出 Excel

### 4.1 环境检查

先确认环境（一次性）：
```bash
python --version && python -c "import openpyxl; print('openpyxl', openpyxl.__version__)"
```
没有 openpyxl 就 `pip install openpyxl`。桌面路径用 `os.path.expanduser("~")`。

### 4.2 生成脚本

把下面模板写入临时 `.py` 文件（**UTF-8**，中文内容不要用 GBK），改 `HEADERS / DATA / FILENAME` 后运行：

```python
# -*- coding: utf-8 -*-
"""抓取信息提炼要点 → 导出 Excel（通用模板）"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ===== 1. 每次按实际内容改这里 =====
SHEET_TITLE = "信息表"
HEADERS = ["标题", "时间", "主题", "简介", "核心内容要点",
           "官方已确认事实（✓）", "媒体报道·待核实（△）",
           "官方来源URL", "权威媒体来源URL", "信息真实性评级",
           "核实状态说明", "备注"]
# 每行元素个数必须与 HEADERS 一致；每条信息一行
DATA = [
    ["示例标题", "2026-08-01", "大模型发布",
     "一句话概述……",
     "1) 事实A\n2) 事实B",
     "✓ 官方文档：……",
     "△ 媒体报道：……",
     "https://官方.com/xxx",
     "https://媒体.com/xxx",
     "高（官方）/ 中（第三方）",
     "官方文档 URL 来自搜索引擎索引，多源一致转述",
     "制作时注意……"],
]
OUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "结果")  # 默认桌面/结果
os.makedirs(OUT_DIR, exist_ok=True)
OUT_PATH = os.path.join(OUT_DIR, "信息表.xlsx")
# ========================================

wb = Workbook()
ws = wb.active
ws.title = SHEET_TITLE

header_fill = PatternFill("solid", fgColor="1F4E79")
header_font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
body_font = Font(name="微软雅黑", size=9)
thin = Side(style="thin", color="B0B0B0")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap_top = Alignment(wrap_text=True, vertical="top", horizontal="left")
center = Alignment(wrap_text=True, vertical="center", horizontal="center")

ws.append(HEADERS)
for c in range(1, len(HEADERS) + 1):
    cell = ws.cell(row=1, column=c)
    cell.fill = header_fill; cell.font = header_font
    cell.alignment = center; cell.border = border

for r in DATA:
    assert len(r) == len(HEADERS), f"行元素数 {len(r)} != 表头 {len(HEADERS)}"
    ws.append(list(r))
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=ws.max_row, column=c)
        cell.font = body_font; cell.alignment = wrap_top; cell.border = border
    ws.row_dimensions[ws.max_row].height = 300

# 列宽：按列内容估算，长文本列（简介/要点/URL）给宽，短列给窄
widths = {1: 28, 2: 20, 3: 18, 4: 45, 5: 60, 6: 45, 7: 45,
          8: 45, 9: 45, 10: 20, 11: 50, 12: 40}
for idx, w in widths.items():
    ws.column_dimensions[get_column_letter(idx)].width = w
ws.row_dimensions[1].height = 30
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{1 + len(DATA)}"

wb.save(OUT_PATH)
print("OK saved:", OUT_PATH)
```

运行注意：
- Windows 下运行用 `PYTHONIOENCODING=utf-8 python xxx.py`，避免中文输出乱码。
- 若内容超长（如视频脚本很长），行高给足，别截断。

### 4.3 验证

生成后读回校验一次：行数、列数、无空列、每条"官方URL"和"媒体URL"各归其位、关键数字没丢。有问题就改脚本重跑。

---

## 📤 流程 Step 5：汇报给用户

1. **如实交代核实结论**：哪些是官方确认、哪些是媒体报道待核实、哪些我这边打不开（环境限制）只做了多源交叉验证——直接说，不遮掩。
2. **给出文件位置**：完整路径（如 `C:\Users\admin\Desktop\结果\xxx.xlsx`）。
3. **列出可直接访问的验证 URL**：把每个话题的官方 + 媒体链接在对话里列出来，方便用户自己打开核对。
4. **提醒使用红线**：哪些数字视频/文章里要加"据媒体报道/约/以官网为准"。

---

## 📌 使用场景示例

- "抓取最新的 AI 新闻" → 抓取 → 按必填字段提炼 → 存 Excel
- "抓取 XX 事件的最新进展" → 抓取 → 提炼时间线 → 存 Excel
- "收集几家公司的信息对比" → 抓取 → 提炼对比字段 → 存 Excel
- 若用户是 MelonAI 视频创作者：增强字段补"视频脚本/分镜/钩子方向/打工人视角"，并与 `.claude/skills/melona.md` 的账号宪法对齐。

## 红线

- 不把媒体推测写成官方事实；不编造来源 URL；不夸大或脑补细节。
- 数字无法多源一致的，标注"待核实"，宁缺毋滥。
- Excel 文件名和保存路径先按用户要求，没要求就 `桌面/结果/`。
