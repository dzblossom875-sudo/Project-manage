# 数据流与工作流

## 模块架构

```
main.py
├── baseline_reader.py   → 读取 data/input/ 最新xlsx，构建 {交易所||债券名称: row} 索引
├── sse_crawler.py       → Playwright 爬上交所，ABS品种翻页+本地关键词过滤
├── szse_crawler.py      → 三级降级爬深交所（Excel→API→Playwright）
└── data_exporter.py     → 合并基准与爬取数据，标注变更，导出 Excel
```

---

## 主流程

```
[用户执行 python main.py]
        │
        ▼
[1. 读取基准文件]
   data/input/*.xlsx
   → baseline_reader.find_latest_baseline()
   → baseline_reader.load_baseline()         ← 字段标准化
   → baseline_reader.build_baseline_index()  ← {交易所||债券名称: row_dict}
        │
        ├─────────────────────────────────────┐
        ▼                                     ▼
[2. 爬取上交所]                        [3. 爬取深交所]
   sse_crawler.crawl_sse()               szse_crawler.crawl_szse()
   │                                     │
   ├─ Playwright 打开上交所债券页          ├─ 尝试 Excel 下载（最快）
   ├─ 选择 ABS 品种                       ├─ 失败 → API 分页
   ├─ 翻页，本地过滤"持有型"/"机构间"      └─ 失败 → Playwright 页面
   └─ 重复≥3条 → 停止翻页                       └─ 重复≥3条 → 停止
        │                                     │
        └──────────────┬──────────────────────┘
                       ▼
               [sse_projects + szse_projects]
               (List[ABSProject] dataclass)
                       │
                       ▼
[4. 合并与标注]
   data_exporter.export_projects()
   │
   ├─ crawled_to_df()        → 爬取数据转标准 DataFrame
   ├─ merge_with_baseline()  → 以基准为底，逐行比对
   │   ├─ 基准有 + 爬取有 → detect_field_changes() → 标"状态变更"或"无变化"
   │   ├─ 基准有 + 爬取无 → 标"无变化"（保留原样）
   │   └─ 基准无 + 爬取有 → 标"新增"
   └─ export_to_excel()      → 写多Sheet Excel + 颜色高亮
                       │
                       ▼
              [data/output/持有型ABS项目汇总_YYYYMMDD_HHMMSS.xlsx]
```

## 增量停止逻辑

```
for each page:
    for each row:
        if row in baseline_index:
            dup_count += 1
        else:
            collect(row)
    if dup_count >= 3:
        break  ← 停止翻页
```

## 比对字段

`COMPARE_FIELDS = ["项目状态", "拟发行金额(亿元)", "计划管理人"]`

仅当以上3个字段中至少一个发生变化时，记为"状态变更"。
