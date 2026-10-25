# 期刊尾注技术参考（顺序编码制 · OOXML 实现）

本文档沉淀期刊投稿式排版的全部技术细节与踩坑记录。来源：一次真实期刊投稿（双栏学报模板）从 V6 到最终提交版的完整迭代经验。SKILL.md 讲流程，本文档讲机制。

## 1. 机制总览

本模块默认采用顺序编码真尾注；学校/期刊本次模板优先。顺序编码、脚注/尾注和栏数是不同层次的选择，不能当作普遍强制要求。

| | 学位论文/课程论文式（thesis-format） | 期刊投稿式（本 skill） |
|---|---|---|
| 文中标注 | 页下注圈码 ①②③，每页重编号，标具体页码 | 上标 [N]，按首次出现编号，具体引文定位按模板保留 |
| 文末 | 参考文献表（保留必要出版范围、不重复） | 尾注区：[N] + 文献条目（含文章起止页码） |
| Word 机制 | footnotes.xml（脚注） | endnotes.xml（尾注）+ 上标 run |
| 编号 | 每页重新编号 | 首次出现连续编号；重复位置用NOTEREF域 |

**真尾注，不做假尾注**：文中 `[N]` 必须是 `endnoteReference` 字段（链接 endnotes.xml，Word/WPS 可点击跳转、自动编号），不能是纯文本假上标。用户对"真尾注"有明确要求——纯文本方案（哪怕视觉一样）会被退回来。

## 2. 核心 XML 模式

### 2.1 文中引用标记：三个上标 run

每个引用位置由三个 run 组成：`[` 上标 + `endnoteReference` 上标 + `]` 上标。

```xml
<w:r>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
    <w:vertAlign w:val="superscript"/>
    <w:sz w:val="18"/><w:szCs w:val="18"/>
  </w:rPr>
  <w:t>[</w:t>
</w:r>
<w:r>
  <w:rPr>
    <w:rStyle w:val="EndnoteReference"/>
    <w:vertAlign w:val="superscript"/>
    <w:sz w:val="18"/><w:szCs w:val="18"/>
  </w:rPr>
  <w:endnoteReference w:id="3"/>
</w:r>
<w:r>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
    <w:vertAlign w:val="superscript"/>
    <w:sz w:val="18"/><w:szCs w:val="18"/>
  </w:rPr>
  <w:t>]</w:t>
</w:r>
```

**关键点**：
- `vertAlign=superscript` 必须显式写在每一个 run 的 rPr 里。缺了它，引用号就落在基线上（用户看到的"右下角"而不是"右上角"）。EndnoteReference 样式本身理论上带上标，但不可依赖——显式写才稳。
- sz=18（9pt 小五）让上标比正文（五号 10.5pt）小一号，观感标准。
- 方括号是 run 里的普通文本，内置好，用户不用手动加。

### 2.2 尾注条目（endnotes.xml）

```xml
<w:endnote w:id="3">
  <w:p>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>[</w:t></w:r>
    <w:r><w:rPr><w:rStyle w:val="EndnoteReference"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:endnoteRef/></w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
        <w:sz w:val="18"/><w:szCs w:val="18"/>
      </w:rPr>
      <w:t xml:space="preserve">] 作者.题名[J].刊名,年(期):起止页码.</w:t>
    </w:r>
  </w:p>
</w:endnote>
```

- `w:endnoteRef` 是自动编号字段，显示的数字由 Word 按引用顺序生成。
- 编号格式靠 `endnotePr` 的 `numFmt=decimal` 指定（Word 尾注默认是小写罗马数字 i, ii, iii，不改就是错的）。要同时写在：① settings.xml 的全局 endnotePr；② 每个节的 sectPr endnotePr。
- endnotes.xml 开头必须有 separator（id=-1）和 continuationSeparator（id=0）两条特殊尾注。

### 2.3 参考文献标题：第一条有效尾注内的独立段落

标题作为第一条被引用尾注（id=2）的第一段，文献条目放其下一段；标题段设 keepNext。不要把标题做成没有正文引用的孤立尾注：本次 Word 16.0 实测会忽略那条孤立记录，导致文件XML有标题而PDF没有。

真正文献条目的 id 从2开始，不再额外占用标题尾注。校验既检查标题文字，也要在目标程序导出后目视确认。

### 2.4 尾注渲染位置：sectEnd

参考文献要出现在"正文之后、英文部分之前"（期刊模板的标准顺序）。做法：每个节的 sectPr 设 `endnotePr` 的 `pos=sectEnd`，文中所有引用都在正文节，尾注就渲染在正文节末尾：

```xml
<w:endnotePr>
  <w:pos w:val="sectEnd"/>
  <w:numFmt w:val="decimal"/>
</w:endnotePr>
```

### 2.5 作者简介：第一页真脚注 + 文末完整段

- **第一页脚注**：footnotes.xml 里一条内容脚注（作者简介，可含基金项目、收稿日期，每项一段），作者名后放 `footnoteReference` 引用它，自动编号显示为上标 1。六号（7.5pt）宋体。
- **文末作者简介段**：征文通知常要求"文末附作者简介"（含联系电话、电子邮箱），在英文部分之后加一段小五号宋体，内容 = 作者简介 + 联系方式。两处并存不冲突，各满足各的要求。

## 3. 重复引用：真尾注与 NOTEREF 域

每个来源首次出现时建立一个 endnoteReference，并用 `cite_N` 书签包住该引用字段。重复位置插入带 `\h` 链接开关的 Word NOTEREF 域，缓存当前编号，外侧方括号上标。这样参考文献只列一次，每处正文标注仍保留。

脚本按正文首次出现顺序重排；校验原引用号是否存在于原文献表，再生成文件。重复的完整著录仅做空白归一后精确匹配，不猜测不同版本与页码是否可合并。

Word/WPS 可能需更新交叉引用域。须在实际目标程序检查显示、更新和跳转；XML校验只保证结构与缓存，不证明所有版本兼容。异常时保留所有引用，不用删标记或隐藏条目的方式绕过。

历史的一文一引删标记方案已停用；共享同一endnote id给多个原生引用、合并编号加vanish等旧方案不再作为推荐路径。

## 4. 踩坑清单（全部真实发生过）

1. **上标丢失**：endnoteReference 的 rPr 没写 vertAlign=superscript → 引用号落在基线。修复：显式写，见 2.1。
2. **空尾注**：用户在 WPS 里手动给尾注加方括号，WPS 把 13 条尾注扩成 18 条，其中 5 条内容被清空。**教训：绝不引导用户手动编辑尾注 XML 层面的东西；方括号必须脚本内置。**
3. **编码乱码**：WPS 重新保存后把 UTF-8 中文存成 GBK 乱码。教训：脚本生成后尽量别让 WPS 反复重存；如需改动优先重新跑脚本。
4. **同一 endnote id 被多处引用的不稳定**：V7 曾让同一文献的多次引用共享同一个 endnote id，WPS 渲染出空白行。现用首次真尾注、重复位置NOTEREF交叉引用，交付前检查目标程序兼容性。
5. **WPS 重存丢 endnotePr pos**：最终提交版被 WPS 重存后，`pos=sectEnd` 被丢掉，尾注退回全文末尾（参考文献跑到了英文部分后面）。教训：交付时提醒用户——如果 WPS 打开再保存过，检查一下参考文献位置；位置不对就重新跑脚本，别手动挪。
6. **"参考文献"标题分离/不显示**：标题是正文段落可能脱节，孤立尾注又会被Word忽略。修复：标题放第一条有效尾注的首段，见2.3。
7. **编号格式**：不写 numFmt=decimal 时尾注显示 i, ii, iii。
8. **纯文本假尾注被拒**：视觉等同也不算数，用户明确要真尾注。别再提纯文本方案。

## 5. 内置双栏期刊模板规格（默认版式）

页面：A4；页边距上下 3cm、左右 2cm；文档网格 lines/linePitch=312。单倍行距为主。

| 部位 | 字号 | 字体 | 其他 |
|---|---|---|---|
| 中文题目 | 小三 15pt | 宋体加粗 | 居中 |
| 中文作者 | 小四 12pt | 楷体 | 居中，后跟作者简介脚注引用 |
| 中文单位 | 五号 10.5pt | 宋体 | 居中 |
| 摘要/关键词 | 小五 9pt | 标签黑体加粗 + 内容楷体 | |
| 中图分类号/文献标识码 | 小五 9pt | 标签黑体 | |
| 一级标题（一、二、三） | 五号 10.5pt | 黑体 | 双栏内 |
| 二级标题（（一）（二）） | 五号 10.5pt | 楷体 | |
| 三级标题（1. 2.） | 五号 10.5pt | 宋体 | |
| 正文 | 五号 10.5pt | 宋体 | 首行缩进 2 字符，双栏（栏距 425 twips） |
| 文中引用 [N] | 小五 9pt | 上标 | 三个 run，见 2.1 |
| 参考文献标题 | 五号 10.5pt | 宋体加粗 | 第一条有效尾注首段（见 2.3） |
| 参考文献条目 | 小五 9pt | 宋体/Times New Roman | 单倍行距，段后 0 |
| 英文题目 | 小三 15pt | Times New Roman 加粗 | 居中，单栏 |
| 英文作者 | 小四 12pt | Times New Roman | 居中 |
| 英文单位 | 五号 10.5pt | Times New Roman | 居中 |
| 英文摘要/关键词 | 五号 10.5pt | Times New Roman | 标签加粗 |
| 作者简介脚注 | 六号 7.5pt | 宋体 | 第一页页下 |
| 文末作者简介 | 小五 9pt | 宋体 | 含联系方式 |

结构顺序（三节）：
1. **第一节（单栏）**：题目、作者、单位、摘要、关键词、中图分类号
2. **第二节（双栏）**：正文（含引言段落、各级标题、结语）→ 尾注区渲染在此节末尾（参考文献标题 + 条目）
3. **第三节（单栏）**：英文题目、英文作者、英文单位、英文摘要、英文关键词、文末作者简介段

注："引言"按此模板不排标题，其段落直接作为正文开头；"结语"排一级标题样式。模板可增补。字号等参数在TPL；栏数与结构还需调整build中的分节逻辑，不能只改TPL就声称支持任意模板。

## 6. 交付后提醒用户（话术要点）

- 打开检查三件事：① 文中 [N] 都在右上角；② "参考文献"标题紧贴列表；③ 尾注可点击跳转。
- 不要在 WPS 里手动改尾注/方括号——改了容易出空条目和乱码，有问题重新跑脚本。
- 如果文件被 WPS 重新保存过，检查参考文献位置是否还在英文部分之前；跑了位置就重跑脚本。
- 重复引用已改为NOTEREF域；目标程序更新域后检查编号、跳转与导出PDF，不能把隐藏重复条目当成默认方案。

## 7. Word实测补充（2026-10-25）

Word 16.0 的样例验证了首次真尾注、重复NOTEREF域更新后的1/2编号、实际显示的参考文献标题，以及正文后英文前的尾注位置。为避免作者简介脚注遇连续分节被强制换页，在settings兼容项加入footnoteLayoutLikeWW8；全局endnotePr同时设sectEnd。此结论只覆盖本次Word环境，WPS仍需按任务验证。

对应规范说明：[Microsoft Open XML footnoteLayoutLikeWW8](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.footnotelayoutlikeword8)。未提供中图分类号时不自动猜填G641。
