# marx-paper-skills

马克思主义理论论文工作流，包含选题、阅读、结构、素材、写作、独立审核与 Word 排版共 11 个 skill。适合马克思主义理论相关方向的理论研究论文；实证研究的数据与方法不由本系列包办。

## 先弄清三处文件

| 位置 | 作用 | 如何更新 |
|---|---|---|
| 本地 Git 源仓库，本机历史安装通常在 `~/.pi/agent/skills/marx-paper-skills` | **唯一编辑源**，保存全部模块、脚本、参考资料、安装元数据和测试 | 只在这里维护，通过测试后提交 |
| Codex 使用目录 `~/.codex/skills/marx-paper*` | 运行副本，Codex 读取这里的文件 | 用同步脚本从源仓库复制，禁止独立改出另一个版本 |
| [GitHub 仓库](https://github.com/lele2026startup-dot/marx-paper-skills) | 远程版本备份与共享 | 本地验证并提交后推送，核对远程提交 |

这三处代表同一套内容。正常收尾时：源仓库与 Codex 中受管理的所有模块文件逐字节一致；GitHub main 与本地提交一致。仓库根 README、测试和同步工具属于维护资料，无需伪装成一个 Codex skill；Codex 根目录的 `.marx-paper-sync.json` 只是同步哈希记录。

实际机器绝对路径、工作记录与备份目录写在当地维护工作区，不提交个人绝对路径、论文内容或联系方式。临时验证副本与历史备份不作为维护源。

## 模块

| Skill | 职责 |
|---|---|
| marx-paper | 总入口与单步路由 |
| marx-paper-topic | 选题、研究切口、候选拟题与检索准备 |
| marx-paper-material | 原著/讲话/政策/学术材料全文检索、TXT+Word 入库 |
| marx-paper-read | 阅读报告与用户思考记录，可选 |
| marx-paper-structure | 结构方案、标题与字数，版本化保存 |
| marx-paper-research | 引文、观点与数据素材，保存真实来源与核验状态 |
| marx-paper-quickwrite | 依据结构与素材快速写作 |
| marx-paper-cowrite | 逐段提问、收集想法、共写与过目 |
| marx-paper-review | 独立检查来源、论证、任务要求与语体 |
| marx-paper-thesis-format | 课程/学位论文脚注与 Word 排版 |
| marx-paper-journal-format | 期刊稿真尾注、重复引用域及默认双栏模板 |

常见流程：topic → 找文献 → read（可选）→ structure → research → quickwrite/cowrite → review → thesis-format/journal-format。指定全文检索与入库直接走 material。

## 保留的默认偏好

期刊式初始框架约 11500–12000 字、3 大×3 小，具体按题目和本次要求调整。课程/学位稿默认圈码页下注和每页重编号；期刊稿默认真尾注与双栏模板。共写保留逐段提问、收想法和过目机制。默认偏好不能覆盖用户已经确认的任务要求。

引用与跨环节格式统一遵循 [共同约定](marx-paper/references/workflow-contract.md)。每处引用均保留，文献表按来源去重：首次引用生成真尾注，重复位置生成指向该尾注的 Word NOTEREF 域，不再自动删后面的标记。域更新与跳转必须在目标 Word/WPS 中验证，XML 自检不等于版面验收。

标准版本按学校/期刊要求。2026-10-25 核对国家标准平台，GB/T 7714—2025 已实施；历史 2015 版示例不应冒充现行完整规范。圈码、双栏、尾注均属模板/本工作流选项。

## 安装与一致性检查

先下载/克隆整个源仓库，不要从网页手工截取 SKILL.md。使用可用的 Python 运行：

```powershell
python scripts/sync_skills.py --check
python scripts/sync_skills.py --apply --backup-dir "本地备份目录" --overwrite-installed
python scripts/sync_skills.py --check --verify-remote
```

默认安装到当前用户 `~/.codex/skills`，可用 `--target` 指定其他安装目录。第一次安装或运行副本曾被手工改动时，需显式 `--overwrite-installed`，旧文件会先备份。脚本检查所需文件与关键章节，再按 SHA-256 比对所有受管理文件；不删除未知文件。差异或远端不一致时返回失败状态，不能忽略。

从 GitHub 更新时先查看本地改动，再 `git pull --ff-only`；有冲突先处理，不能硬覆盖用户改动。维护时在源仓库改→运行测试→审阅差异→提交→同步 Codex→在用户授权范围内推送 GitHub→检查远端与安装副本。完整规则见 [维护与同步](docs/maintenance.md)。

## 验证与依赖

```powershell
python -m unittest discover -s tests -v
```

核心依赖是 python-docx、lxml；抓取使用 Python 标准库。课程/学位排版另需 pandoc；真实每页圈码需 pywin32 和 Word/WPS COM。环境缺项时报告具体缺项，使用已有替代能力，不盲目安装系统依赖。

脚本默认拒绝覆盖已有输出。授权覆盖时显式传 `--overwrite`，备份旧文件后发布；文档生成会先验证临时文件，失败不替换原成果。测试覆盖编码、引用转换与保留、错误输入、防覆盖和安装同步；真实学术内容与最终 Word/WPS 版面仍需任务内核验。
