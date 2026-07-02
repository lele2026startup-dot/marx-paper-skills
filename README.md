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
选题 → 材料整理 → 写作 → 格式规范 → ……
```

## 包含的 skill

| Skill | 状态 | 职责 |
|-------|------|------|
| **marx-paper-format** | ✅ 已完成 | 脚注圈码 ①②③ + 每页重新编号 + GB/T 7714 参考文献 + Word 排版 |
| **marx-paper-topic** | 🚧 第一层完成 | 论文题目特征识别（骨架/路数/用词/气质）；后续补拟题与判题方法 |
| marx-paper-research | 🚧 规划中 | 材料搜集、分类整理、提炼 |
| marx-paper-write | 🚧 规划中 | 正文写作 |

## marx-paper-format 详解

把马理论科论文草稿（markdown 或 docx）加工成格式规范、可直接提交的 Word 文档。

### 核心能力

- **脚注**：页下注，圈码 ①②③，每页重新编号，标注具体页码
- **参考文献**：GB/T 7714-2015 规范，只列被引用过的，不标页码，不重复
- **Word 排版**：标题黑体居中、正文宋体小四、1.5 倍行距、首行缩进、参考文献单起一页、全篇黑色
- **引号配对**：修正 pandoc 中文引号配对错误，直引号正确配对为弯引号
- **页码核实**：有源 PDF 时定位真实页码，无 PDF 时标注待确认，绝不编造

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

帮马理论科论文拟定、打磨、评估题目。当前版本完成第一层：优质题目的特征识别。

### 已完成

- **5 种标题骨架**：主+破折号副、冒号式、设问式、从X到Y、陈述式
- **6 类研究路数**：概念史/话语史、文本/思想个案、历史事件/运动/制度、理论批判/辨析、现实应用/路径创新、比较/演进
- **用词偏好三层**：动作词、范式词组、限定语
- **7 条深层气质**：大背景+小切口、问题意识、材料限定、理论框架、学术张力、主副分工、时间括注
- **44 题样例库**：按路数分类归档，每题标注骨架与切口（`marx-paper-topic/references/title-samples.md`）

### 待做

- 从研究领域+材料+问题意识生成候选题目
- 题目检验清单（好不好写、站不站得住）
- 题目迭代打磨套路
- 不同场景的选题策略（课程论文 vs 学位论文 vs 期刊投稿）

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
└── marx-paper-topic/                   选题拟题 skill
    ├── SKILL.md
    └── references/
        └── title-samples.md
```
