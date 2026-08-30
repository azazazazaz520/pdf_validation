# PDF 转 Word 独立可行性验证

本目录用于在 Prism 主工程之外验证 PaddleOCR PDF 解析、表格识别、多页重组和 DOCX 导出能力。

## 环境

- Python：3.11.15
- PaddlePaddle：3.3.1，CPU 版
- PaddleOCR：3.7.0
- PaddleX：3.7.2，安装 `ocr` 额外依赖
- Word 导出：`python-docx` 1.2.0
- PDF 页面渲染：`pypdfium2` 5.13.0
- PDF 结构检查：`pypdf` 6.16.2
- 模型缓存：`D:\pdf_validation\model_cache`

Python 3.11 用于兼容当前 Windows CPU 版 PaddlePaddle。执行脚本时使用虚拟环境解释器的完整路径，避免调用系统 Python：

```powershell
uv venv --python 3.11 .venv
& .\.venv\Scripts\python.exe -m pip install --upgrade pip
& .\.venv\Scripts\python.exe -m pip install --index-url https://www.paddlepaddle.org.cn/packages/stable/cpu/ paddlepaddle==3.3.1
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 运行验证

```powershell
& .\.venv\Scripts\python.exe generate_fixture.py
& .\.venv\Scripts\python.exe generate_table_fixture.py

# 轻量文本/版面解析，不启用表格识别
& .\.venv\Scripts\python.exe run_validation.py --engine structure-lite

# 轻量版面解析，启用表格识别
& .\.venv\Scripts\python.exe run_validation.py --engine structure-table-lite --input .\fixtures\synthetic_table_one_page.pdf

# PaddleOCR-VL，支持多页重组
& .\.venv\Scripts\python.exe run_validation.py --engine vl --input .\fixtures\synthetic_text_table.pdf
```

使用已缓存模型进行离线复测时，在当前 PowerShell 会话中设置：

```powershell
$env:PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK = 'True'
& .\.venv\Scripts\python.exe run_validation.py --engine structure-lite
```

每次运行会在 `outputs\<engine>_<timestamp>` 下保存 JSON、Markdown、DOCX 和 `report.json`。

## 第一版服务器服务

服务文件为 `pdf_to_word_service.py`，默认先检测 PDF 文本层：可直接提取的文字型 PDF 走快速文本路线，其余文件使用常驻的 `structure-table-lite` OCR 路线。任务状态保存在内存，任务文件保存在 `service_data\jobs`，过期任务按 TTL 清理。

安装服务依赖后启动：

```powershell
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
$env:PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK = 'True'
$env:PDF_SERVICE_TOKEN = '请替换为随机长令牌'
& .\.venv\Scripts\python.exe .\pdf_to_word_service.py
```

默认仅监听 `127.0.0.1:8765`。部署到服务器时，应通过防火墙或反向代理限制访问范围；确需局域网访问时再设置 `PDF_SERVICE_HOST=0.0.0.0`，并保留 `PDF_SERVICE_TOKEN`。当前服务仍为单进程原型，正式部署前应将 OCR 推理移到独立 worker。

接口：

```text
GET    /health
POST   /api/pdf-to-word/jobs              multipart 字段名：file
GET    /api/pdf-to-word/jobs/{job_id}
GET    /api/pdf-to-word/jobs/{job_id}/result
DELETE /api/pdf-to-word/jobs/{job_id}
```

环境变量：

- `PDF_SERVICE_ENGINE`：默认 `structure-table-lite`，也可使用 `structure-lite`。
- `PDF_SERVICE_ROUTE_MODE`：默认 `auto`。`auto` 自动检测文本层；`text` 强制文本层快速路线；`ocr` 跳过检测并强制使用 OCR。
- `PDF_SERVICE_TEXT_MIN_PAGE_CHARS`：文本层页面的最小有效字符数，默认 20。
- `PDF_SERVICE_TEXT_MIN_PAGE_RATIO`：达到最小字符数的页面占比阈值，默认 0.6。
- `PDF_SERVICE_EXPORT_MODE`：默认 `hybrid`。`hybrid` 保留原始页面图像，并在文档后半部分附加 OCR 可编辑文本；`text` 仅导出 OCR 文本。
- `PDF_SERVICE_PAGE_IMAGE_MAX_PIXELS`：混合模式单页图像像素上限，默认 4194304。
- `PDF_SERVICE_PAGE_IMAGE_JPEG_QUALITY`：混合模式页面图像 JPEG 质量，默认 88。
- `PDF_SERVICE_TOKEN`：配置后要求 `Authorization: Bearer <token>` 或 `X-API-Key`。
- `PDF_SERVICE_DATA_ROOT`：任务临时目录根路径。
- `PDF_SERVICE_MAX_UPLOAD_BYTES`：默认 50 MiB。
- `PDF_SERVICE_MAX_PAGES`：默认 100 页。
- `PDF_SERVICE_JOB_TTL_SECONDS`：默认 3600 秒。
- `PDF_SERVICE_ALLOWED_ORIGINS`：可选，逗号分隔的 CORS 来源列表。

## 当前验证结论

- `structure-lite`：两页样本成功，缓存后约 34 秒；关闭表格识别时，表格会作为版面内容或图片处理，不应作为最终的可编辑表格方案。
- `structure-table-lite`：单页表格成功识别，16 个单元格内容与输入一致，生成的 DOCX 中包含真实 `w:tbl` 表格；缓存后约 42 秒。
- `vl`：两页样本成功重组为一个 DOCX，表格内容可编辑；缓存后初始化约 73 秒，两页 CPU 推理约 324 秒，模型权重约 1.79 GiB，运行工作集约 4.3-6 GiB。
- 真实日报样本：两页 VL 推理成功，正文和编号列表内容基本完整；项目符号条目内容保留但转换为普通段落，未保留为 Word 列表，初始化约 74 秒，CPU 推理约 939 秒。
- 第一版服务端到端：`structure-table-lite` 处理真实日报成功，任务创建、状态查询、DOCX 下载和输入校验均通过；本机首次服务任务约 145 秒，生成 38 KiB DOCX。服务原型在 CPU 推理期间可能出现状态查询延迟，后续接入服务器时应使用独立工作进程或进程级任务队列。
- `PPStructureV3` 默认完整配置：启用 oneDNN 时触发当前 CPU 路径的属性转换错误；关闭 oneDNN 后资源占用持续增长到约 7 GiB 且未在可接受时间内完成，不建议作为当前机器的默认配置。
- 文本层快速路线：已实现自动检测；满足阈值的 PDF 不加载 OCR 模型，直接按页提取文本并生成 DOCX。当前已通过自动分流、无模型加载和文本内容回归测试。
- 真实日报文本层快速路线：分析约 0.131 秒，DOCX 导出约 0.034 秒，总耗时约 0.165 秒；该结果不代表扫描件或复杂版面文件的耗时。

验证目录中的 `fixtures` PDF 是合成样本；真实扫描 PDF、复杂表格、多栏文档、旋转页面、图片和公式仍需补充样本后单独验收。图片型 PDF 建议使用默认的 `hybrid` 模式：前半部分保留页面视觉内容，后半部分提供 OCR 可编辑文本。混合模式会增加 DOCX 页数和文件体积，页面图像像素上限可通过环境变量调整。
