# 独立可行性验证摘要

验证日期：2026-08-30

## 验证范围

本次只在 `D:\pdf_validation` 的 Python 3.11 虚拟环境中运行，未修改 Prism 主工程。验证覆盖：

1. Windows CPU 环境的 PaddlePaddle、PaddleOCR 和 PaddleX 依赖安装。
2. PDF 多页输入、文本识别、版面解析、表格识别和 DOCX 导出。
3. 模型缓存后的重复运行，以及跳过模型源连通性检查的运行方式。
4. DOCX ZIP/OOXML 完整性和可编辑表格结构。
5. 第一版 HTTP 服务的任务生命周期、认证、输入校验和 DOCX 下载。

## 结果

| 路线 | 样本 | 结果 | 关键指标 |
| --- | --- | --- | --- |
| `structure-lite` | 两页文本、标题、表格 | 成功 | 初始化 13.4 秒，推理 21.0 秒；表格未启用结构识别 |
| `structure-table-lite` | 单页 4 列 × 4 行表格 | 成功 | 初始化 18.0 秒，推理 23.5 秒；16 个单元格内容正确，DOCX 含真实表格 |
| `vl` | 单页表格 | 成功 | 初始化 74.1 秒，推理 98.0 秒；表格内容正确，DOCX 含真实表格 |
| `vl` | 两页文本和表格 | 成功 | 初始化 72.6 秒，推理 323.6 秒；2 页重组为 1 个 DOCX |
| `structure` 完整配置 | 两页样本 | 不建议 | oneDNN 路径报错；关闭 oneDNN 后资源占用约 7 GiB，未完成可接受的推理 |

## 第一版服务器服务

服务文件：`D:\pdf_validation\src\pdf_to_word_service.py`

端到端验证使用 `structure-table-lite` 处理 `C:\Users\Lenovo\Desktop\日报-2026-08-17.pdf`，结果如下：

- 健康检查返回 `200`，未携带凭证访问受保护接口返回 `401`。
- 任务创建返回 `200`，识别为 2 页；最终状态为 `succeeded`，本机首次服务任务从创建到完成约 145 秒。
- DOCX 下载返回 `200`，文件约 38 KiB；ZIP/OOXML 检查通过，包含 25 个非空段落，其中 18 个使用 Word 原生编号或项目符号样式，日报中的关键标题、`plugin_*` 条目和“验证结果”均存在。
- 使用 4 列 × 4 行合成表格再次调用服务，生成 DOCX 中包含 1 个真实 Word 表格、4 行、16 个单元格，单元格内容检查通过。
- 上传非 PDF 文件返回 `415`，未进入模型推理。

当前实现适合第一版独立联调：任务状态保存在单进程内存，任务文件按 TTL 清理，默认只监听本机回环地址，支持 Bearer Token 或 `X-API-Key`。测试同时观察到 CPU 推理阶段 HTTP 状态查询会出现延迟，客户端需要容错重试；正式部署前应把推理执行移到独立工作进程，并在反向代理层配置 TLS、访问控制和请求限流。

## 真实日报样本

输入文件：`C:\Users\Lenovo\Desktop\日报-2026-08-17.pdf`

- 文件为 2 页、Letter 页面、未加密的文字型 PDF，没有表单、JavaScript 或嵌入图片，适合验证中文文字、段落、编号列表、项目符号和分页重组。
- `structure-lite` 和 `structure-table-lite` 均成功完成，耗时分别约 63 秒和 67 秒；正文及编号列表基本可读，但项目符号存在符号粘连或丢失，且各生成两个分页 DOCX。
- `vl` 成功完成，初始化 73.7 秒，CPU 推理 938.6 秒，2 页重组为 1 个 DOCX。正文和编号列表内容基本完整，5 个 `plugin_*` 条目内容均保留，但项目符号被转换为普通段落。
- VL 产物通过 DOCX ZIP/OOXML 检查，包含 30 个文本节点，原文中的 `plugin_scan`、`plugin_prepare_install` 和未来计划内容均存在；当前环境缺少 LibreOffice，未完成 DOCX PNG 渲染级检查。

因此，这份日报证明 VL 对中文文字型文档的内容恢复质量明显优于轻量路线，但其 CPU 耗时约 15.6 分钟，不能直接作为 Prism 的默认交互式转换路径。轻量路线适合默认处理，VL 更适合作为用户主动选择或低置信度时的增强路径；如果要求保留 Word 原生编号/项目符号，还需要在导出层根据解析结果重新构建 Word 列表。

## 依赖和运行资源

- 依赖安装过程中，除 `paddleocr` 外还需要显式安装 `paddlex[ocr]`，否则 `PPStructureV3` 的 OCR 相关依赖不完整。
- DOCX 导出需要显式安装 `python-docx`。
- 所有已加载模型的缓存总量约 2.84 GiB，其中 VL 模型权重约 1.79 GiB。
- 轻量结构化路线适合 CPU 本地处理和后续集成验证；VL 路线在本机可运行，但 CPU 延迟很高，不宜直接作为交互式默认路径。

## 产物索引

- 轻量文本报告：`artifacts\outputs\structure-lite_20260830_152338\report.json`
- 表格报告：`artifacts\outputs\structure-table-lite_20260830_152228\report.json`
- VL 两页合并报告：`artifacts\outputs\vl_20260830_153400\report.json`
- VL 两页 DOCX：`artifacts\outputs\vl_20260830_153400\word\synthetic_text_table.docx`
- 真实日报 VL 报告：`artifacts\outputs\vl_20260830_170021\report.json`
- 真实日报 VL DOCX：`artifacts\outputs\vl_20260830_170021\word\日报-2026-08-17.docx`
- 第一版服务日报 DOCX：`artifacts\outputs\service-e2e\daily.docx`
- 第一版服务表格 DOCX：`artifacts\outputs\service-e2e\synthetic_table.docx`

## 结论与下一步

当前可以确认“独立运行、表格识别、DOCX 导出、多页重组”在技术上可行，但尚不足以确认真实业务 PDF 的转换质量。建议下一阶段先加入至少三类真实样本：

- 中文扫描件：验证识别准确率、倾斜和低清晰度。
- 多栏或带页眉页脚报告：验证阅读顺序和分页。
- 跨页复杂表格：验证列对齐、合并单元格和续表；当前已通过带边框续表探针，仍需真实复杂表格样本。

在真实样本通过后，再决定 Prism 内部采用 `structure-table-lite` 作为常规路径，还是把 VL 作为按需增强路径；目前不建议直接集成完整 `PPStructureV3` 默认配置。
