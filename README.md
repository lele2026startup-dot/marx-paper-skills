# marx-paper-skills

马克思主义理论一级学科论文全流程 skill 系列。

## 适用学科

马克思主义理论一级学科（0305），含六个二级方向：

| 编号 | 二级方向 |
|------|----------|
| 1 | 马克思主义基本原理 |
| 2 | 马克思主义中国化研究 |
| 3 | 思想政治教育 |
| 4 | 国外马克思主义研究 |
| 5 | 马克思主义发展史 |
| 6 | 中国近现代史基本问题研究 |

适用于这六个方向下纯文科、无实证、纯文字性的课程论文、期末论文、学位论文。

## 设计理念

把马理论科论文写作拆成若干独立环节，每个环节做成一个 skill。各 skill 互相独立又彼此衔接，按需调用。

```
选题（topic）→ 找文献 → 整理素材（research）→ 写作 → 格式规范（format）
```

## 包含的 skill

| Skill | 状态 | 职责 |
|-------|------|------|
| **marx-paper-format** | ✅ 已完成 | 脚注圈码 ①②③ + 每页重新编号 + GB/T 7714 参考文献 + Word 排版 |
| **marx-paper-topic** | ✅ 三层完成 | 特征识别 + 情况确认与分流 + 拟题配方；末尾衔接"找文献" |
| **marx-paper-research** | 🚧 雏形完成 | 读文献、按来源分类整理原话引用、提取出处信息成素材库 |
| marx-paper-write | 🚧 规划中 | 正文写作 |

## marx-paper-format 详解

把马理论科论文草稿（markdown 或 docx）加工成格式规范、可直接提交的 Word 文档。

### 核心能力

- **脚注**：页下注，圈码 ①②③，每页重新编号，标注具体页码
- **参考文献**：GB/T 7714-2015 规范，只列被引用过的，不标页码，不重复
- **Word 排版**：标题黑体居中、正文宋体小四、1.5 倍行距、首行缩进、参考文献单起一页、全篇黑色
- **引号配对**：修正 pandoc 中文引号配对错误，直引号正确配对为弯引号
- **页码核实**：三种来源——有源 PDF 定位真实页码、无 PDF 标注待确认、来自 research 整理素材的复用原话直接用别人标好的出处（不必再核实，但完成后于对话里告知用户哪些是复用的），绝不编造

### 技术方案

| 难题 | 方案 | 脚本 |
|------|------|------|
| 圈码 ①②③ 不渲染 | customMarkFollows 属性 + 写入 Unicode 圈码字符 | footnote_circle_marks.py |
| 每页重新编号 | Word COM 接口查真实页码，按页分组重新赋圈码 | renumber_footnotes_per_page.py |
| 中文引号配对 | 关 pandoc smart，自配对（奇数左、偶数右） | render_docx.py |
| 字体/行距/排版 | pandoc 转换 + python-docx 后处理 | render_docx.py |

### 依赖

- [pandoc](https://pandoc.org/) — markdown ↔ docx 转换
- [python-docx](https://python-docx.org/) — docx XML 操作
- [pywin32](https://pypi.org/project/pywin32/) — Word/WPS COM 接口（查真实页码）
- Microsoft Word 或 WPS Office（任一即可）

## marx-paper-topic 详解

帮马理论科论文拟定、打磨、评估题目。三层已完成：特征识别 + 情况确认与分流 + 拟题配方。

### 已完成

- **第一步 特征识别**：5 种标题骨架、6 类研究路数、用词偏好三层、7 条深层气质
- **第二步 情况确认与分流**：三问（场景/有无资料/方向明确度）→ 场景细化 → 有资料/没资料路径分叉 → 没资料路径五个提问锚点
- **第三步 拟题配方**：方向 × 路数 × 骨架 × 切口 = 候选题目，筛选后留 10 个左右让用户选
- **衔接段**：主题确定后提醒用户去知网找 25-30 篇文献（关键词组合搜），征得同意建参考论文文件夹，衔接 research skill
- **44 题样例库**：按路数分类归档，每题标注骨架与切口（`marx-paper-topic/references/title-samples.md`）

## marx-paper-research 详解

把用户找来的参考文献里的原话引用整理成分类素材库。当前是雏形版本。

### 核心能力

- **读文献**：PDF（用 pdf skill）或 MD（直接读），建议转 MD 但不强求
- **分类整理原话**：按来源分文件夹——领导人、马克思恩格斯、列宁、毛泽东、国外学者等（具体分类跟用户商量确认）
- **提取出处信息**：每条原话连同别人论文里的标注信息（页码、版本等）一起提取，供 format skill 复用
- **建文件夹前征得同意**：不擅自创建目录

### 设计要点

- 复用别人用过的原话是合理做法：别人已标好出处，直接拿来用，不必每条都翻原书核实
- 分类不写死：默认几类，但跟用户商量后按实际情况定
- 衔接 format skill：复用素材里的原话，format 不必再核实页码，但会在对话里告知用户哪些是复用的

## 目录结构

```
marx-paper-skills/
├── README.md                           系列总览
├── marx-paper-format/                  格式规范 skill
│   ├── SKILL.md
│   ├── references/
│   │   └── gbt7714-format.md
│   └── scripts/
│       ├── render_docx.py
│       ├── footnote_circle_marks.py
│       └── renumber_footnotes_per_page.py
├── marx-paper-topic/                   选题拟题 skill
│   ├── SKILL.md
│   └── references/
│       └── title-samples.md
└── marx-paper-research/                素材整理 skill（雏形）
    └── SKILL.md
```
