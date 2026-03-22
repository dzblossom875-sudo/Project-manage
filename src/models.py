"""数据模型定义"""

from dataclasses import dataclass


@dataclass
class ABSProject:
    """ABS项目标准数据结构"""
    bond_name: str
    manager: str
    bond_type: str
    amount: str
    status: str
    update_date: str
    accept_date: str
    exchange: str = ""
