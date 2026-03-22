# CHANGELOG

所有版本变更按时间倒序记录。

---

## [docs] - 2026-03-22

### 项目规范整理
- 新增 `docs/` 目录（CHANGELOG.md / summary.md / workflow.md）
- 新增 `.gitignore`，屏蔽 `__pycache__`、`data/output/`、Excel 临时锁定文件
- 固化 Python 路径（`D:\install\python.exe`）到 `.claude/settings.local.json`
- 精简 `memory/WORK_LOG.md`：移除历史版本内容（归 CHANGELOG）和技术难点（归 summary），保留跨会话上下文+当前状态
- 精简 `使用说明.md`：移除"爬取策略"内部细节，聚焦用户操作
- 合并重复内容：`workflow.md` 补充架构图（原 summary.md），`summary.md` 移除重复的架构概览

---

## [v2.0] - 2026-03-19

### feat
- **baseline_reader.py**（新增）：读取 `data/input/` 最新 xlsx，按 `交易所||债券名称` 建索引，提供 `check_duplicate` 快速查重
- **data_exporter.py**（重写）：以基准文件为底合并爬取数据，输出 3 个 Sheet（全部/新增/状态变更），新增行绿色高亮，变更行黄色高亮
- **sse_crawler.py**（重写）：放弃依赖搜索框，改为选 ABS 品种后翻页+本地关键词过滤，避免 JS 动态渲染导致的搜索失效
- **szse_crawler.py**（重写）：三级降级链路 Excel下载 → API分页 → Playwright 页面，提升成功率
- **增量停止机制**：爬取时与基准比对，重复项目 ≥ 3 个即停止翻页，避免无效全量爬取

### fix
- 移除 v1 所有调试文件（`debug_*.py/html`）、旧爬虫（`browser_crawler.py`、`playwright_crawler*.py`、`sse_final_crawler.py`、`manual_crawler.py`）、旧启动脚本（`run_*.bat/py`）

### 关联文件
`main.py`, `src/baseline_reader.py`, `src/data_exporter.py`, `src/sse_crawler.py`, `src/szse_crawler.py`

---

## [v1.0] - 2026-03-18

### feat
- 项目初始化：搭建 `src/`, `data/`, `memory/` 目录结构
- 基础模块：`sse_crawler`、`szse_crawler`、`data_exporter`、`models`
- 手动辅助模式：用户手动搜索后脚本提取页面数据

### 探索过的方案（已废弃）
- 直接 API 调用 → 交易所反爬，返回空数据
- Selenium 自动化 → 搜索框无法触发筛选
- Playwright 搜索 → 同上，JS 触发机制不兼容
