# 输入结构

输入为 UTF-8 JSON。每题有稳定且唯一的 `id`、`section`、`question`、`answer`。题数不固定；正文按 `items` 的次序输出。已有题号保持原样，新增题可用后缀。

```json
{
  "title": "课程考试问答速查",
  "sections": {"L1": "L1 核心概念"},
  "logic_terms": ["AND", "OR", "not automatically"],
  "core_terms": ["CAGR"],
  "case_terms": [],
  "items": [
    {"id": "L1-01", "section": "L1", "question": "Market size 与 growth rate 有什么区别？", "answer": "Market size measures value or volume; growth rate measures change over time.", "priority": "core"},
    {"id": "L1-02", "section": "L1", "question": "弹性模量的简化关系是什么？", "answer": "Under the stated simplified model, [[F:elasticity]], where ρ is density and v is shear-wave speed.", "priority": "normal"}
  ],
  "formulas": {
    "elasticity": {"runs": [{"text": "E = 3ρv"}, {"text": "2", "superscript": true}], "plain_text": "E = 3ρv2"}
  }
}
```

- `priority` 为 `core`、`normal`、`case`，只控制题号底色，不断言考试概率。
- 可加 `origin`、`source` 等内部审阅字段；生成器忽略它们，不打印来源。
- 普通答案为字符串，可含中文、希腊字母和符号。`[[F:key]]` 插入公式；简单公式的 `runs` 支持 `superscript` 和 `subscript`。
- 复杂公式可以用 `{"omml": "<m:oMath xmlns:m=...>...</m:oMath>", "plain_text": "..."}` 保留原生 Word 数学结构。XML 来自核对过的源式；`plain_text` 用于核验，不能替代公式原图检查。
- 不支持的对象、表格或图示不能静默跳过：先判断是否携带必要知识，用 Documents 工作流补入或经授权转成等价问答。
- DOCX 提取需检查表格、文本框、跨段答案及原生公式；只读 `Answer:` 第一段不是完整答案。
- 内部记录原题数、添加题数、删除的非正文段落、追加内容和原文件哈希；布局成功不能抵消内容遗漏。

使用 workspace dependencies 返回的 Python，明确传入输入输出路径：

```text
python scripts/build_cheatsheet.py sheet.json --output result.docx --font 3.5 --line 3.8 --columns 8
python scripts/check_cheatsheet.py sheet.json result.docx --pdf rendered/result.pdf --expected-pages 2
python scripts/restyle_titles.py result.docx
```

上面的 `python` 都替换为 bundled Python 绝对路径。建稿默认拒绝覆盖现有文件；确认输入已反映用户最新编辑、且本轮要原位修订时才用 `--overwrite`。`restyle_titles.py` 专门原位改标题样式，不改变文字。
