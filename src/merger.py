"""
数据合并模块

职责：合并爬取数据与基准数据，返回标准化 DataFrame，不负责文件写入。
"""

import pandas as pd
from typing import List, Dict, Optional
from dataclasses import asdict


# 输出字段顺序（严格按此顺序）
OUTPUT_COLUMNS = [
    "债券名称",
    "交易所",
    "计划管理人",
    "品种",
    "拟发行金额(亿元)",
    "交易所审批状态",
    "项目发行状态",
    "跟进和投放状态",
    "交易所系统更新日期",
    "预计发行时间",
    "受理日期",
    "备注",
    "变更类型",
    "变更详情",
]

# 比对字段（用于检测变更）
COMPARE_FIELDS = ["交易所审批状态", "拟发行金额(亿元)", "计划管理人"]


def _make_key(exchange: str, bond_name: str) -> str:
    """生成唯一键用于比对"""
    return f"{exchange}||{bond_name}"


def _detect_field_changes(old: dict, new: dict) -> List[str]:
    """
    比较两行数据的关键字段，返回变更描述列表

    Args:
        old: 基准行数据
        new: 新爬取行数据

    Returns:
        变更描述列表，如 ["交易所审批状态: 已受理 -> 已反馈"]
    """
    changes = []
    for field in COMPARE_FIELDS:
        old_val = str(old.get(field, "")).strip()
        new_val = str(new.get(field, "")).strip()
        if old_val == "nan" or old_val == "NaT" or old_val == "None":
            old_val = ""
        if new_val == "nan" or new_val == "NaT" or new_val == "None":
            new_val = ""
        if old_val != new_val and new_val:
            changes.append(f"{field}: {old_val} -> {new_val}")
    return changes


def _crawled_to_df(sse_projects: List, szse_projects: List) -> pd.DataFrame:
    """
    将爬取的项目对象列表转为标准化 DataFrame

    Args:
        sse_projects: 上交所项目列表（ABSProject dataclass）
        szse_projects: 深交所项目列表（ABSProject dataclass）

    Returns:
        标准化的 DataFrame，包含映射后的字段
    """
    all_projects = sse_projects + szse_projects
    if not all_projects:
        return pd.DataFrame()

    data = [asdict(p) for p in all_projects]
    df = pd.DataFrame(data)

    # 字段映射（从爬虫模型到标准字段）
    column_mapping = {
        "exchange": "交易所",
        "bond_name": "债券名称",
        "manager": "计划管理人",
        "bond_type": "品种",
        "amount": "拟发行金额(亿元)",
        "status": "交易所审批状态",  # 原"项目状态"改为"交易所审批状态"
        "accept_date": "受理日期",
        "update_date": "交易所系统更新日期",  # 原"更新日期"改为"交易所系统更新日期"
    }
    df = df.rename(columns=column_mapping)

    # 新增字段默认值
    df["项目发行状态"] = "未发行"  # 新增字段，默认未发行
    df["跟进和投放状态"] = ""  # 新增字段，默认空
    df["预计发行时间"] = ""  # 新增字段，默认空
    df["备注"] = ""  # 新增字段，默认空
    df["变更类型"] = "新增"  # 爬取到的默认为新增
    df["变更详情"] = ""  # 默认空

    return df


def _build_index(df: pd.DataFrame) -> Dict[str, dict]:
    """
    用 (交易所, 债券名称) 构建索引，用于快速查重和状态比对

    Args:
        df: DataFrame 数据

    Returns:
        {key: row_dict} 字典，key = "交易所||债券名称"
    """
    index = {}
    for _, row in df.iterrows():
        exchange = str(row.get("交易所", "")).strip()
        bond_name = str(row.get("债券名称", "")).strip()
        if bond_name:
            key = _make_key(exchange, bond_name)
            index[key] = row.to_dict()
    return index


def merge(
    sse_projects: List,
    szse_projects: List,
    baseline_df: Optional[pd.DataFrame] = None
) -> pd.DataFrame:
    """
    合并爬取结果与基准数据，返回标准化 merged_df

    合并逻辑（以基准为底）：
    1. 基准有 + 爬取有 → 用爬取数据更新比对字段，标注"状态变更"或"无变化"
    2. 基准有 + 爬取无 → 保留基准原样，标注"无变化"
    3. 基准无 + 爬取有 → 追加，标注"新增"

    字段映射规则：
    - 原"项目状态" → "交易所审批状态"
    - 原"更新日期" → "交易所系统更新日期"
    - 原"投放和跟进情况" → "跟进和投放状态"
    - 新增"项目发行状态"：基准有则保留，否则默认"未发行"
    - 新增"预计发行时间"：默认空字符串，人工填写

    Args:
        sse_projects: 上交所爬取项目列表
        szse_projects: 深交所爬取项目列表
        baseline_df: 基准数据 DataFrame（可选）

    Returns:
        合并后的 DataFrame，包含 OUTPUT_COLUMNS 定义的所有字段
    """
    # 转换爬取数据
    crawled_df = _crawled_to_df(sse_projects, szse_projects)

    # 没有基准数据时，直接返回爬取数据（全标为新增）
    if baseline_df is None or baseline_df.empty:
        if crawled_df.empty:
            return pd.DataFrame(columns=OUTPUT_COLUMNS)
        # 确保所有输出字段存在
        for col in OUTPUT_COLUMNS:
            if col not in crawled_df.columns:
                crawled_df[col] = ""
        return crawled_df[OUTPUT_COLUMNS]

    # 标准化基准数据字段名（兼容旧字段名）
    baseline_df = baseline_df.copy()

    # 字段名迁移（旧→新）
    field_mapping = {
        "项目状态": "交易所审批状态",
        "更新日期": "交易所系统更新日期",
        "投放和跟进情况": "跟进和投放状态",
    }
    for old_field, new_field in field_mapping.items():
        if old_field in baseline_df.columns and new_field not in baseline_df.columns:
            baseline_df[new_field] = baseline_df[old_field]

    # 确保新增字段存在
    if "项目发行状态" not in baseline_df.columns:
        baseline_df["项目发行状态"] = "未发行"
    if "预计发行时间" not in baseline_df.columns:
        baseline_df["预计发行时间"] = ""

    # 构建索引
    crawled_index = _build_index(crawled_df)
    baseline_index = _build_index(baseline_df)

    rows_out = []

    # 第一轮：遍历基准数据（以基准为底）
    for _, base_row in baseline_df.iterrows():
        exchange = str(base_row.get("交易所", "")).strip()
        bond_name = str(base_row.get("债券名称", "")).strip()
        key = _make_key(exchange, bond_name)

        row_dict = base_row.to_dict()

        if key in crawled_index:
            # 基准有 + 爬取有 → 检测变更
            crawled_row = crawled_index.pop(key)
            changes = _detect_field_changes(row_dict, crawled_row)

            if changes:
                # 有变更：更新比对字段
                for field in COMPARE_FIELDS:
                    new_val = str(crawled_row.get(field, "")).strip()
                    if new_val and new_val != "nan":
                        row_dict[field] = new_val
                # 更新交易所系统更新日期
                if "交易所系统更新日期" in crawled_row:
                    row_dict["交易所系统更新日期"] = crawled_row["交易所系统更新日期"]
                row_dict["变更类型"] = "状态变更"
                row_dict["变更详情"] = "; ".join(changes)
            else:
                # 无变更
                row_dict["变更类型"] = "无变化"
                row_dict["变更详情"] = ""
        else:
            # 基准有 + 爬取无 → 保留原样
            row_dict["变更类型"] = "无变化"
            row_dict["变更详情"] = ""

        rows_out.append(row_dict)

    # 第二轮：处理爬取中有但基准中没有的（新增）
    for key, crawled_row in crawled_index.items():
        row_dict = crawled_row.copy()
        row_dict["变更类型"] = "新增"
        row_dict["变更详情"] = "基准文件中不存在"
        rows_out.append(row_dict)

    # 构建结果 DataFrame
    result = pd.DataFrame(rows_out)

    # 确保所有输出字段存在，按指定顺序排列
    for col in OUTPUT_COLUMNS:
        if col not in result.columns:
            result[col] = ""

    result = result[OUTPUT_COLUMNS]

    # 打印统计信息
    new_count = sum(1 for t in result["变更类型"] if t == "新增")
    changed_count = sum(1 for t in result["变更类型"] if t == "状态变更")
    unchanged_count = sum(1 for t in result["变更类型"] if t == "无变化")

    print(f"\n=== 合并统计 ===")
    print(f"基准项目: {len(baseline_df)} 条")
    print(f"本次爬取: {len(crawled_df)} 条")
    print(f"合并后:   {len(result)} 条")
    print(f"  - 新增:     {new_count}")
    print(f"  - 状态变更: {changed_count}")
    print(f"  - 无变化:   {unchanged_count}")
    print(f"================\n")

    return result
