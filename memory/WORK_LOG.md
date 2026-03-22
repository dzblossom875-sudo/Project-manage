# 持有型ABS项目爬虫 — 跨会话上下文

## 项目概述
爬取上交所和深交所的持有型ABS申报项目，与基准文件比对后标注变更，导出 Excel。

**筛选条件**：债券名称含"持有型"或"机构间"；与基准重复 ≥ 3 条即停止翻页。

---

## 当前文件清单

| 文件 | 功能 |
|------|------|
| `main.py` | 主程序入口，解析参数，串联各模块 |
| `src/baseline_reader.py` | 读取 `data/input/` 最新xlsx，构建比对索引 |
| `src/data_exporter.py` | 合并基准+爬取数据，标注变更，导出Excel |
| `src/sse_crawler.py` | 上交所 Playwright 爬虫（ABS品种翻页+本地过滤）|
| `src/szse_crawler.py` | 深交所爬虫（Excel下载 → API → Playwright 三级降级）|
| `src/models.py` | ABSProject dataclass 数据模型 |

---

## 当前状态

- **最后操作**：项目整理（新增 docs/、修复 Python 路径、清理 __pycache__）
- **最后运行**：2026-03-22，输出 `data/output/持有型ABS项目汇总_20260322_190322.xlsx`
  - 基准 72 条，状态变更 4 条，新增 0 条
- **Python 路径**：`D:\install\python.exe`（已固化到 `.claude/settings.local.json`）
- **待续事项**：无，项目功能正常，可直接运行

---

> 版本历史见 `docs/CHANGELOG.md`，技术设计见 `docs/summary.md`，数据流见 `docs/workflow.md`
