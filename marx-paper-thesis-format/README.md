# 课程/学位论文排版

另有独立预置 **大学本科毕业论文格式一**：先读 [规则与输入说明](references/undergraduate-01.md)，通过 `build_docx.py --preset undergraduate-01 --meta 本地元信息.json` 选择；本页其余说明为原通用格式。


本目录是 marx-paper-skills 的排版模块，入口为 SKILL.md；完整系列及安装同步方式见仓库根 README.md。

默认以 Pandoc 脚注生成 Word，采用圈码、每页重编号、宋体小四和1.5倍行距。具体学校模板与用户本次要求优先。圈码是本工作流偏好，不是所有学校或国标的统一要求。

脚本入口：`python scripts/build_docx.py 草稿.md 成品.docx`；简单编号加 `--no-circle`。默认拒绝覆盖已有输出；授权覆盖时加 `--overwrite` 并保留备份。已有 docx 不直接送入 MD 生成脚本，按文档工具流程保留原格式和脚注后编辑。

依赖：pandoc、python-docx；每页圈码需 pywin32 和 Microsoft Word/WPS COM。开始前检查实际环境；没有排版引擎时不能声称完成真实分页检查。

著录和引文核验遵循 `../marx-paper/references/workflow-contract.md`。文献版本与真实读取文本一致；二手页码保留核验状态。
