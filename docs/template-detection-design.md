# DOCX 模板探测设计

日期：2026-06-01

## 目标

模板探测器的最终目标是：给少量输入/输出样例，反推出一个可以继续复用的 Word 模板。

模板分两层：

1. 格式层：字体、段落、表格、页眉页脚、编号、边框、页面设置等。
2. 内容层：哪些文本是静态内容，哪些应成为结构化字段、循环、条件块。

当前实现先落第一层的可验证 MVP：从单个 `.docx` 生成格式保持型模板。

## 当前 MVP

命令：

```bash
office-tools template inspect-format input.docx
office-tools template infer-format input.docx --out-template template.docx --out-data data.json --out-profile profile.json
office-tools template render template.docx --data data.json --out rendered.docx
office-tools template compare-format input.docx rendered.docx
```

实现策略：

- 直接读取 DOCX zip 包。
- 只改 `word/document.xml`、页眉页脚、脚注、尾注、批注等文本 part 中的 `<w:t>`。
- 模板保留原始 OOXML 结构，只把文本节点替换为 `{{text.0001}}` 这类占位符。
- `data.json` 保存每个占位符对应的原始文本。
- `profile.json` 保存占位符位置、段落/表格计数和格式指纹。
- 渲染时只把占位符文本节点替换回数据，不重建段落、run、表格。

这种路线的好处是：第一版就能高度保真，因为所有格式 XML 都沿用原文档。

## 验证标准

`compare-format` 目前检查：

- 提取文本是否一致。
- 段落数量是否一致。
- run 数量是否一致。
- 表格数量是否一致。
- styles.xml / numbering.xml 字节长度是否一致。
- 去除文本后的 OOXML 结构指纹是否一致。

已验证：

- 合成 DOCX smoke：`npm run smoke:template`
- 本机真实文档：`C:\Code\doctemplate_codex\历史沿革报告.docx`
  - 探测出 543 个文本字段。
  - 渲染回填后 `textEqual=true`。
  - 渲染回填后 `profileEqual=true`。
  - 段落数 326、表格数 10、run 数 541 均一致。

## 当前边界

已覆盖：

- 主文档文本。
- 页眉页脚文本。
- 脚注、尾注、批注文本。
- 格式保持型模板生成。
- 重新渲染验证。

暂未解决：

- 多样例对齐。
- 字段语义命名，例如 `company.name`。
- 表格循环行。
- 条件块。
- 图片、文本框、图形中的动态内容。
- Word 域代码、目录、交叉引用的智能更新。

## 下一步

1. 多样例格式对齐：比较多个输出文档，提取稳定段落、稳定表格和样式簇。
2. 内容动态区识别：对多个样例的文本节点做 diff，区分静态文本和变量文本。
3. 字段聚合：把相邻 run 中属于同一业务字段的文本合并成一个 slot。
4. 表格循环识别：识别重复行模式，生成 `items[]` 结构。
5. 语义命名：用 LLM 只做字段命名和类型建议，不直接改 DOCX。
6. 验证闭环：每次推断都必须 render back，再做文本和格式指纹比较。

## 技术原则

- 格式层尽量确定性，优先 OOXML 结构分析，不靠截图猜。
- LLM 只用于语义层辅助，不作为格式保真的唯一依据。
- 模板生成必须能回归验证；不能验证的推断只能进入 report，不能默认写入模板。
