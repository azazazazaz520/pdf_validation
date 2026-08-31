# 项目目录说明

`D:\pdf_validation` 是独立的 PDF 转 Word 验证项目。目录按代码、测试、样本、文档和运行产物分离，便于后续抽取到 Prism 的服务端或 Rust/Tauri 调用层。

```text
D:\pdf_validation
├─ src/                  核心实现和服务入口
│  ├─ pdf_to_word_service.py
│  ├─ pdf_worker.py
│  ├─ pdf_to_word_exporter.py
│  ├─ pdf_layout.py
│  ├─ pdf_routing.py
│  ├─ ocr_quality.py
│  ├─ job_store.py
│  └─ run_validation.py
├─ tests/                unittest 自动化测试
├─ scripts/              合成 PDF 样本生成脚本
├─ fixtures/             可复现的合成 PDF 输入
├─ docs/                 评估报告、验证摘要和实施结论
├─ artifacts/            本地验证产物，已加入忽略规则
│  ├─ outputs/            PaddleOCR 运行结果
│  └─ qa/                 手工检查、渲染和端到端验证结果
├─ model_cache/          PaddleOCR 模型缓存，不纳入版本控制
├─ service_data/         服务运行状态和任务文件，不纳入版本控制
├─ requirements.txt      Python 依赖
└─ README.md             使用入口
```

## 目录维护规则

- 新增核心模块放入 `src/`，模块之间使用 `src` 包内相对导入。
- 新增自动化测试放入 `tests/`，从项目根目录执行 `python -m unittest discover -s tests -v`。
- 新增合成输入放入 `fixtures/`；生成脚本放入 `scripts/`。
- OCR、服务端到端和 Word 渲染结果放入 `artifacts/qa/`；可重复运行的模型输出放入 `artifacts/outputs/`。
- `model_cache/`、`service_data/`、`artifacts/` 和 Python 缓存目录只作为本地运行目录，不提交到版本库。
- WeKnora 等真实样本如果需要长期保留，应放在项目外部或受控的样本目录中，避免将原始文档和业务内容混入代码仓库。
