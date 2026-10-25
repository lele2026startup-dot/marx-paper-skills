# marx-paper-skills 维护规则

本仓库是唯一编辑源；具体三处位置与命令见 README.md、docs/maintenance.md。Codex 的 marx-paper* 是生成的运行副本，不独立修改。未知安装文件先调查，不删除。

1. 开始修改前读 git status，保留已有或并发改动；读取本次用户授权与当地维护工作区约定。
2. 涉及选题/快写时确认完整工作流仍在；同步脚本会验证 skills-manifest.json 的所需文件和关键章节。
3. 引文规则与环节交接由 marx-paper/references/workflow-contract.md 统一维护；模块不可另写相反要求。保留用户写作偏好和本次特殊要求的区别。
4. 修改脚本后跑针对性的回归测试并检查最终 diff。新增/移动资源同时更新 skills-manifest.json；不靠字数阈值证明行为正确。
5. 只暂存本次文件，提交说明用清楚中文；不改写历史。完成授权的同步任务时推送 GitHub，最后运行安装哈希与远端提交检查。无推送授权时明确报告未推送。
6. 个人论文、meta.json、联系方式、机器绝对路径、备份和测试生成文档不入库；测试样例须虚构。保留安装元数据 agents/openai.yaml。
7. 新增防错机制配真实行为测试。XML 自检与 Word/WPS 视觉验收分开报告，未完成的项目不得声称通过。
