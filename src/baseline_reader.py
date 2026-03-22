"""
基准文件读取模块
从 data/input/ 读取最新的 xlsx 基准文件，标准化字段，提供比对索引
"""

import os
import pandas as pd
from typing import Dict, Optional


STANDARD_COLUMNS = [
    "债券名称", "交易所", "计划管理人", "品种",
    "拟发行金额(亿元)", "项目状态", "更新日期", "受理日期",
    "投放和跟进情况", "备注"
]

COMPARE_FIELDS = ["项目状态", "拟发行金额(亿元)", "计划管理人"]


def find_latest_baseline(input_dir: str = "data/input") -> Optional[str]:
    """
    在 input_dir 中找到最新的 .xlsx 文件（按修改时间排序）

    Returns:
        最新 xlsx 文件的完整路径，找不到返回 None
    """
    if not os.path.isdir(input_dir):
        print(f"基准目录不存在: {input_dir}")
        return None

    xlsx_files = [
        f for f in os.listdir(input_dir)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if not xlsx_files:
        print(f"基准目录中没有 xlsx 文件: {input_dir}")
        return None

    xlsx_files.sort(
        key=lambda f: os.path.getmtime(os.path.join(input_dir, f)),
        reverse=True
    )

    latest = os.path.join(input_dir, xlsx_files[0])
    print(f"找到基准文件: {xlsx_files[0]}")
    return latest


def load_baseline(filepath: str) -> pd.DataFrame:
    """
    读取基准 Excel 并标准化字段

    Returns:
        标准化后的 DataFrame
    """
    df = pd.read_excel(filepath)

    exchange_col_map = {
        "上交所": "上交所",
        "深交所": "深交所",
        "深": "深交所",
        "沪": "上交所",
    }
    if "交易所" in df.columns:
        df["交易所"] = df["交易所"].map(
            lambda x: exchange_col_map.get(str(x).strip(), str(x).strip())
        )

    if "受理日期" in df.columns:
        df["受理日期"] = pd.to_datetime(df["受理日期"], errors="coerce")
        df["受理日期"] = df["受理日期"].dt.strftime("%Y-%m-%d")

    if "更新日期" in df.columns:
        df["更新日期"] = df["更新日期"].astype(str).str.strip()

    print(f"基准文件加载完成: {len(df)} 条记录")
    return df


def build_baseline_index(df: pd.DataFrame) -> Dict[str, dict]:
    """
    用 (交易所, 债券名称) 构建基准索引，用于快速查重和状态比对

    Returns:
        {key: row_dict} 字典，key = "交易所||债券名称"
    """
    index = {}
    for _, row in df.iterrows():
        exchange = str(row.get("交易所", "")).strip()
        bond_name = str(row.get("债券名称", "")).strip()
        if bond_name:
            key = f"{exchange}||{bond_name}"
            index[key] = row.to_dict()
    return index


def make_key(exchange: str, bond_name: str) -> str:
    return f"{exchange}||{bond_name}"


def check_duplicate(baseline_index: Dict[str, dict], exchange: str, bond_name: str) -> bool:
    """检查项目是否已存在于基准文件中"""
    return make_key(exchange, bond_name) in baseline_index


if __name__ == "__main__":
    fpath = find_latest_baseline()
    if fpath:
        df = load_baseline(fpath)
        idx = build_baseline_index(df)
        print(f"基准索引: {len(idx)} 个项目")
        for k in list(idx.keys())[:5]:
            print(f"  {k}")
