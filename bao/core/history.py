"""历史案例写入 — 已适配 FBA 装箱单

保留接口兼容性，当前仅打印提示信息。
"""

import sys
from datetime import date
from pathlib import Path
from typing import Optional


# zhishiku 路径
_ZHISHIKU_ROOT = Path(__file__).parent.parent.parent.parent / "zhishiku"


def append_historical_case(
    declaration=None,
    report=None,
    attachments: Optional[dict] = None,
) -> Optional[Path]:
    """预留接口：写入历史案例（当前未实现）"""
    return None
