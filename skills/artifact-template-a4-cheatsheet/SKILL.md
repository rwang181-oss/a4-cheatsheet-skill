---
name: artifact-template-a4-cheatsheet
description: "Use when the user requests an exam cheatsheet, cheat sheet, 考试速查表, 小抄, 双面 A4 Word, or revisions to this Chinese question-and-answer layout. Also use when A4 Cheatsheet or $artifact-template-a4-cheatsheet is named."
---

# A4 Cheatsheet

复用中英问答考试速查表样式。下列偏好为默认值，本轮明确要求优先。参考文档使用通用示例，只提供样式，不把示例题目带进用户的课程。

## 默认规则

| 项目 | 默认值 |
|---|---|
| 取舍顺序 | 原有问答正文完整保留 → 在页数限制内选择较大的可读字体 |
| 交付 | 可编辑 Word；一张 A4 双面，即两页，横向 297 × 210 mm |
| 语言 | 保留中文问题、中文释义及英文考试关键词 |
| 问题标题 | 非粗体，深红 B00020；中文、英文和题号统一处理 |
| 章节标题 | 非粗体，深蓝 1F4E79 |
| 答案 | 黑色常规字重；数字、公式及关键词局部浅色高亮 |
| 高亮 | 黄 FFF0A8 概念；蓝 DAECFF 数值；粉 FFDCDD 限定/逻辑；绿 E0F2DD 案例 |
| 空间 | 四边 4 mm、栏距 1 mm、细竖分栏线，无页眉、页脚、页码或打印说明 |
| 密集参考参数 | 8 栏、正文 3.5 pt、固定行距 3.8 pt；并非所有新文档的固定字号 |

## 执行

1. 读取 `artifact-template.json`，查看 `assets/preview.png`，保留 `assets/reference.docx` 原样。配合当前 Documents skill 与 bundled workspace runtime 创建及渲染 Word。
2. 清点输入和题目。已有问答逐字保留正文，含中文、数字、单位、例子、限定条件和公式。来源行、目录、重复页眉页脚、编辑优先级标签移到内部记录；不能误删“资金来源”等正文。保存原标签分类，但不能把笔记标签说成老师确认会考。
3. 补充以用户授权及实际课程资料为据，标记新增题；原答案采用追加。内部保留来源映射，打印稿不附逐题来源。仅有课件/录音时，先构建中英问答和覆盖清单；不把未读来源说成已覆盖。
4. 按[输入结构](references/input-format.md)整理内容，用 `scripts/build_cheatsheet.py` 创建候选。支持任意题数及课程；排版时不改写正文。
5. 按[排版与核验](references/production.md)调参。先试 4 pt；内容较少时放大字号、减少栏数。需要时使用3.5 pt 密集参考参数。候选间只变排版，不为放大字号合并、缩写或删题。仍放不下时保存完整稿，准确说明冲突，按本轮授权决定精简；不能悄悄交付第三页或删减版。
6. 使用 Documents 的 `render_docx.py` 渲染最终 Word，逐页检查中文、公式、上下标、栏末、页末和高亮。运行 `scripts/check_cheatsheet.py` 核验所有题目、顺序、全文、颜色、非粗体和实际页数。再回到原输入核对；JSON 自洽不等于没有漏抄源文档。
7. 修订成品按指定路径编辑；“原文件修改”指当前成品时直接改当前成品，不据此改最初的课程源材料。纯标题样式修改可用 `scripts/restyle_titles.py`，文字不变并重新渲染。不通过重建覆盖用户后来做的编辑。
8. 只交付要求的 Word，说明实际页数、字号和内容保留情况。内部 PDF、PNG 和核验清单不额外交付。打印提示放在聊天：100% 实际大小、横向双面、短边翻转。

样式可稳定复用，但分页受内容、字体与渲染环境影响，每次都需验证。4 mm 不能保证所有打印机无裁切；不承诺微小字体人人可读。复杂公式通过 Documents 数学工作流保持结构，不能随意拼普通文字。用户本轮的新页数、字号或取舍要求优先。
