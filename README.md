# humanities-paper-skills

人文社科纯文科论文全流程 skill 系列。适用于无实证、无数据、纯文字性的人文社科类论文（课程论文、期末论文、学位论文）。

## 设计理念

把论文写作拆成若干独立环节，每个环节做成一个 skill。各 skill 互相独立又彼此衔接，按需调用。

```
选题 → 材料整理 → 写作 → 格式规范 → ……
```

## 包含的 skill

| Skill | 状态 | 职责 |
|-------|------|------|
| **academic-paper-format** | ✅ 已完成 | 脚注圈码 ①②③ + 每页重新编号 + GB/T 7714 参考文献 + Word 排版 |
| academic-paper-topic | 🚧 规划中 | 选题、拟定题目与结构 |
| academic-paper-research | 🚧 规划中 | 材料搜集、分类整理、提炼 |
| academic-paper-write | 🚧 规划中 | 正文写作 |

## academic-paper-format 详解

把中文学术论文草稿（markdown 或 docx）加工成格式规范、可直接提交的 Word 文档。

### 核心能力

- **脚注**：页下注，圈码 ①②③，每页重新编号，标注具体页码
- **参考文献**：GB/T 7714-2015 规范，只列被引用过的，不标页码，不重复
- **Word 排版**：标题黑体居中、正文宋体小四、1.5 倍行距、首行缩进、参考文献单起一页、全篇黑色
- **页码核实**：有源 PDF 时定位真实页码，无 PDF 时标注待确认，绝不编造

### 技术方案

| 难题 | 方案 | 脚本 |
|------|------|------|
| 圈码 ①②③ 不渲染 | `customMarkFollows` 属性 + 写入 Unicode 圈码字符 | `footnote_circle_marks.py` |
| 每页重新编号 | Word COM 接口查真实页码，按页分组重新赋圈码 | `renumber_footnotes_per_page.py` |
| 字体/行距/排版 | pandoc 转换 + python-docx 后处理 | `render_docx.py` |

### 依赖

- [pandoc](https://pandoc.org/) — markdown ↔ docx 转换
- [python-docx](https://python-docx.org/) — docx XML 操作
- [pywin32](https://pypi.org/project/pywin32/) — Word/WPS COM 接口（查真实页码）
- Microsoft Word 或 WPS Office（任一即可）

### 目录结构

```
academic-paper-format/
├── SKILL.md                          主流程说明
├── references/
│   └── gbt7714-format.md             GB/T 7714 各类文献著录格式
└── scripts/
    ├── render_docx.py                pandoc 后处理（字体/行距/排版）
    ├── footnote_circle_marks.py      圈码 ①②③（customMarkFollows 方案）
    └── renumber_footnotes_per_page.py  每页重新编号（Word COM 查页码）
```

## 适用范围

人文社科类纯文字性论文（哲学、历史、文学、马克思主义理论、政治学等）。不适用于含实验数据、图表、公式的实证研究论文。
