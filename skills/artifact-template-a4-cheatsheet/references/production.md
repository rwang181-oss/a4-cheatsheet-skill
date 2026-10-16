# 排版与核验

## 调参

`assets/reference.docx` 是公开通用示例，`assets/preview.png` 为第一页。示例使用 2 栏、12 pt 字体，输入为 `examples/sample.json`。生成器复制样式及页面结构，清除示例内容后填入本次数据。示例题数和页数不约束用户的输出；密集参数需用用户的实际内容重新验证。

1. 全文固定，先试 4 pt、6–8 栏，行距约字号 + 0.3–0.6 pt。
2. 溢出时先去多余空行、段距及非必要换栏；使用 1 mm 栏距、4 mm 边距。再以 0.5 pt 步长缩至3.5 pt 密集参考字号。
3. 有明显空白时放大字号，必要时减少栏数；比较实际渲染，不凭字数宣布“两页”。少量内容不复制旧题或加无关材料凑满。
4. 每题问答通常保持同栏；单题比整栏还高时合理分段，不让 keep-together 挤出空白页或裁掉内容。
5. 最终选择保全正文的较大可用字号。准确报告 pt 数，不把 3.5 pt 写成“四号”或 4 pt。

英文 Arial Narrow，中文 Arial Unicode MS，显式设置 ASCII、HAnsi、EastAsia 和 CS 字体。字体替换后重新核验分页。标题不能重新加粗，色彩规则见主说明。

## 渲染排障

调用 `load_workspace_dependencies` 并定位当前 Documents 的 `render_docx.py`；不把历史版本号或用户名作为跨机器硬依赖。

macOS bundled LibreOffice 曾忽略系统字体，中文空白但文字提取仍有内容。已验证修复：渲染进程设置 `FONTCONFIG_FILE` 指向 bundled LibreOffice 的 `Contents/Resources/fontconfig/fonts.conf`。从本轮 runtime 根目录定位并确认存在，必要时用 bundled `fc-match` 检查解析。

```text
env FONTCONFIG_FILE=<verified fonts.conf> <bundled-python> <Documents/render_docx.py> result.docx --output_dir render-final --emit_pdf --dpi 300
```

每次迭代使用独立渲染目录，避免旧 page-3.png 误导。渲染返回 session ID 就等待实际结束；初始空输出不是完成证据。

小字号原生 OMML 曾被渲染为约 12 pt 而覆盖邻行。先尝试修复数学字号；简单式子可改为可编辑 Word 文本和真实上下标，注明表示方式变化。复杂分式、积分、矩阵等必须保留结构，通过 Documents 数学工作流解决，不能随意拼成扁平文本。

## 交付证据

- 来源清点和原文核对：全部来源已读或明确标出缺失；原题、新增题、追加说明和移除的非正文分别记录；原始材料哈希不变，除非用户明确要求改它。
- Word：ID、次序、问答、中文及公式与确认后的输入一致；红色问题/蓝色章节均非粗体；无页眉、页脚、页码、脚注或逐题来源。
- PDF：实际两页 A4 横向；运行 `check_cheatsheet.py`。页数元数据不能替代渲染后的页数；遇到失败先解决，不跳过。
- 图片：看完整布局，再用 300 dpi 四个重叠象限或局部栏检查。中文完整、无粘连重叠；公式上下标正确；标题不压前段；栏末页末不裁字。文本提取通过不能代替目视检查。
- 最新成品路径与核验记录哈希对应。原位修改后也要重渲染；不能沿用旧预览宣布完成。

成品底角不印字号参数、页码、版本、来源或打印提示；在聊天中交付 Word 和必要的真实参数说明。4 mm 打印机边界不是通用保证。
