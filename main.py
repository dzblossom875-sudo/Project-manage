#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
持有型ABS项目动态跟踪爬虫 v2

工作流:
1. 读取 data/input/ 中最新的xlsx基准文件
2. 分别爬取上交所和深交所的持有型ABS项目
3. 与基准文件比对，标注新增和状态变更
4. 导出含变更标注的Excel到 data/output/

停止条件: 每个交易所爬取时，与基准重复项目>=3个即停止翻页
搜索策略: 选ABS品种后翻页+本地过滤持有型关键词

使用方法:
    python main.py                    # 默认(有头浏览器)
    python main.py --headless         # 无头模式
    python main.py --exchange sse     # 只爬上交所
    python main.py --exchange szse    # 只爬深交所
"""

import io
import sys
import os
import argparse

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

    # Step 2: 爬取上交所
    if args.exchange in ["sse", "both"]:
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

    # Step 3: 爬取深交所
    if args.exchange in ["szse", "both"]:
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

    # Step 4: 导出 (基准 + 爬取合并)
    from data_exporter import export_projects

    if sse_projects or szse_projects or baseline_df is not None:
        print("-" * 60)
        print("正在合并基准数据与爬取数据...")
        print("-" * 60)
        output_path = export_projects(
            sse_projects,
            szse_projects,
            output_dir="data/output",
            baseline_df=baseline_df,
            baseline_index=baseline_index if baseline_index else None,
        )
        if output_path:
            print(f"\n>> 导出成功: {output_path}")
        else:
            print("\n>> 导出失败")
    else:
        print("\n没有获取到任何项目数据, 也没有基准文件")

    print()
    print("=" * 60)
    print("执行完成")
    print("=" * 60)


if __name__ == "__main__":
    main()
