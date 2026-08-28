# FORGE v1.2 Release Notes / 发布说明

**Preserve the content. Reforge the format.**

FORGE v1.2 turns a template generator into a schema-driven DOCX formatting and
transformation engine for AI agents.

FORGE v1.2 将原有模板生成器升级为面向 AI Agent 的 Schema 驱动 DOCX
格式化与文档转换引擎。它支持生成、检查、内容保真重排、批量汇编和安全定点编辑，
输出仍是可继续编辑的 Word 文档；除非执行明确的编辑指令，否则不会改写用户原文。

## Major Capabilities / 主要能力

- **Content-preserving DOCX reformatting** — source-preserving renderer plus hard preservation validation (text, structure, media, sequence, tables, media relationships).
  **内容保真的 DOCX 重排**——使用源文件保真渲染器，并对文本、结构、媒体、内容顺序、表格结构和媒体关系进行严格校验。
- **Batch assembly** — normalize + package-aware merge for many DOCX files, with per-item preservation and assembly payload fingerprints.
  **批量文档汇编**——对多份 DOCX 进行格式归一化和包感知合并，同时验证每份源文档的内容保真，并生成汇编载荷指纹。
- **Reference / Custom / Guided formatting profiles** — learn formats from a sample, define overrides, or walk through a guided session; all persisted in `FORGE_HOME`.
  **参考 / 自定义 / 引导式格式方案**——可从样本文档提取格式、定义自定义覆盖项，或通过引导流程创建方案，并统一保存在 `FORGE_HOME`。
- **Dynamic templates** — register user DOCX templates without touching core code.
  **动态模板**——无需修改核心代码即可注册用户自己的 DOCX 模板。
- **Source-preserving targeted editing** — explicit edits protected by source-version checks and an Edit Contract.
  **源文件保真的定点编辑**——通过源文件版本检查和编辑契约保护每一项明确编辑，阻止未授权内容变化。
- **Checkpoint/resume for large assemblies** — interrupted batches resume from completed items only after config/source verification.
  **大型汇编断点续作**——批量任务中断后，在配置和源文件校验通过的前提下复用已完成项目并继续执行。

## Honest Known Limitations / 已知限制

- TOC, nested table editing, tracked-changes editing, and complex field editing are not implemented.
  尚未实现目录生成、嵌套表格编辑、修订内容编辑和复杂域编辑。
- Replacements cannot cross hyperlink / field / image boundaries.
  文本替换不能跨越超链接、域或图片边界。
- Page numbers do not support PAGE/NUMPAGES, roman numerals, or complex odd/even layouts.
  页码暂不支持 `PAGE/NUMPAGES` 域、罗马数字以及复杂的奇偶页布局。
- Large DOCX processing still loads the package XML in memory.
  处理大型 DOCX 时仍需将包内 XML 加载到内存，尚未采用流式解析。
- Benchmarks are development-machine reference values, not guarantees.
  性能数据仅为开发机上的参考结果，不构成通用性能保证。

## Compatibility / 兼容性

- Legacy `reformat_docx` and built-in generation behavior are unchanged.
  旧版 `reformat_docx` 与内置文档生成行为保持不变。
- Public v1.1 code and tag remain untouched.
  公开版 v1.1 的代码和标签均保持不变。
