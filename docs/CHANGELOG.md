# CHANGELOG

所有版本变更按时间倒序记录。

---

## [v2.1] - 2026-03-22

### refactor: 架构重构 — 输出层解耦

**新增模块：**
- `merger.py`：合并逻辑独立，返回标准化 DataFrame，不负责文件写入
  - 字段映射：原"项目状态"→"交易所审批状态"、原"更新日期"→"交易所系统更新日期"、原"投放和跟进情况"→"跟进和投放状态"
  - 新增字段："项目发行状态"（默认"未发行"）、"预计发行时间"（默认空）
  - 比对字段更新为：["交易所审批状态", "拟发行金额(亿元)", "计划管理人"]

- `output/` 目录：解耦的输出层
  - `excel_exporter.py`：提交版 Excel，4 个 Sheet（全部项目/新增项目/状态变更/跟进状态汇总）
    - 微软雅黑字体、深蓝表头、自动列宽（中文按2计算）
    - 新增行绿色底色 RGB(198,239,206)，状态变更行黄色底色 RGB(255,235,156)
    - 全部项目 Sheet 含合计行（拟发行金额列用 SUM 公式）
  - `html_exporter.py`：单文件自包含 HTML Dashboard
    - 数据序列化为 JSON 内嵌，支持筛选/排序/分组/行展开
    - 跟进状态可 inline 修改，变更日志实时记录
    - 支持下载 CSV/XLS

- `scripts/recalc.py`：Excel 公式验证脚本，返回 `{"status": "success"}` 或错误信息

**修改：**
- `main.py`：使用新架构（merger + output），添加 `--export-only` 参数用于测试
- `data_exporter.py`：添加 deprecation 注释，保留以兼容旧版

**验证：**
- `python -c "from merger import merge; print('merger OK')"` ✓
- `python main.py --export-only` 成功导出 Excel 和 HTML
- `python scripts/recalc.py <excel>` 返回 success

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
