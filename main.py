#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
持有型ABS项目动态跟踪爬虫 v2

工作流:
1. 读取 data/input/ 中最新的xlsx基准文件
2. 分别爬取上交所和深交所的持有型ABS项目
3. 与基准文件比对，标注新增和状态变更
4. 导出含变更标注的Excel和HTML Dashboard到 data/output/

停止条件: 每个交易所爬取时，与基准重复项目>=3个即停止翻页
搜索策略: 选ABS品种后翻页+本地过滤持有型关键词

使用方法:
    python main.py                    # 默认(有头浏览器)
    python main.py --headless         # 无头模式
    python main.py --exchange sse     # 只爬上交所
    python main.py --exchange szse    # 只爬深交所
    python main.py --export-only      # 跳过爬取，仅导出（用于测试）
"""

import io
import sys
import os
import argparse
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))


def parse_args():
    parser = argparse.ArgumentParser(description="持有型ABS项目爬虫 v2")
    parser.add_argument("--headless", action="store_true", help="浏览器无头模式")
    parser.add_argument(
        "--exchange",
        choices=["sse", "szse", "both"],
        default="both",
        help="选择交易所: sse-上交所, szse-深交所, both-两者",
    )
    parser.add_argument(
        "--export-only",
        action="store_true",
        help="跳过爬取，仅导出（用于测试输出层）",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    print("=" * 60)
    print("持有型ABS项目动态跟踪爬虫 v2")
    print("=" * 60)

    # Step 1: 读取基准文件
    from baseline_reader import find_latest_baseline, load_baseline, build_baseline_index

    baseline_df = None
    baseline_index = {}
    baseline_path = find_latest_baseline("data/input")
    if baseline_path:
        baseline_df = load_baseline(baseline_path)
        baseline_index = build_baseline_index(baseline_df)
        print(f"基准索引: {len(baseline_index)} 个项目\n")
    else:
        print("未找到基准文件, 将不做增量比对\n")

    sse_projects = []
    szse_projects = []
    headless = args.headless

    # Step 2: 爬取上交所（如果不是仅导出模式）
    if not args.export_only and args.exchange in ["sse", "both"]:
        print("-" * 60)
        print("开始爬取上交所...")
        print("-" * 60)
        try:
            from sse_crawler import crawl_sse
            sse_projects = crawl_sse(
                baseline_index=baseline_index,
                headless=headless,
            )
        except Exception as e:
            print(f"上交所爬取失败: {e}")
        print()

    # Step 3: 爬取深交所（如果不是仅导出模式）
    if not args.export_only and args.exchange in ["szse", "both"]:
        print("-" * 60)
        print("开始爬取深交所...")
        print("-" * 60)
        try:
            from szse_crawler import crawl_szse
            szse_projects = crawl_szse(
                baseline_index=baseline_index,
                headless=headless,
            )
        except Exception as e:
            print(f"深交所爬取失败: {e}")
        print()

    # Step 4: 导出（使用新的输出层架构）
    from merger import merge
    from output.excel_exporter import export_excel
    from output.html_exporter import export_html

    # 爬取汇总
    print(f"\n=== 爬取汇总 ===")
    print(f"上交所: {len(sse_projects)} 条")
    print(f"深交所: {len(szse_projects)} 条")
    print(f"================")

    if sse_projects or szse_projects or baseline_df is not None:
        print("-" * 60)
        print("正在合并基准数据与爬取数据...")
        print("-" * 60)

        # 合并数据
        merged_df = merge(sse_projects, szse_projects, baseline_df)

        if not merged_df.empty:
            # 生成时间戳
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            os.makedirs('data/output', exist_ok=True)

            # 导出 Excel
            excel_path = f'data/output/持有型ABS项目汇总_{timestamp}.xlsx'
            print(f"\n导出 Excel: {excel_path}")
            export_excel(merged_df, excel_path)

            # 导出 HTML Dashboard
            html_path = f'data/output/持有型ABS项目追踪_{timestamp}.html'
            print(f"导出 HTML: {html_path}")
            export_html(merged_df, html_path)

            print(f"\n>> 导出成功")
            print(f"   Excel: {excel_path}")
            print(f"   HTML:  {html_path}")
        else:
            print("\n>> 没有数据可导出")
    else:
        print("\n没有获取到任何项目数据, 也没有基准文件")

    print()
    print("=" * 60)
    print("执行完成")
    print("=" * 60)


if __name__ == "__main__":
    main()
