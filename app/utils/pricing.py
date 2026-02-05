from __future__ import annotations

from datetime import datetime
from typing import Optional

from app.utils.time_utils import get_now

# 定价规则:
# - 30 天 = 15 元
# - 其余按天线性折算
DEFAULT_BASE_DAYS = 30
DEFAULT_BASE_PRICE_CENTS = 1500  # 15.00 元


def calculate_remaining_days(expires_at: Optional[datetime], now: Optional[datetime] = None) -> Optional[int]:
    """
    计算从“今天”到到期日期的剩余天数(按日期差计算)。

    - 返回 None: 无到期时间
    - 返回 0: 已到期或到期日为今天
    """
    if not expires_at:
        return None

    if now is None:
        now = get_now()

    remaining_days = (expires_at.date() - now.date()).days
    return max(int(remaining_days), 0)


def calculate_price_cents(
    remaining_days: Optional[int],
    base_days: int = DEFAULT_BASE_DAYS,
    base_price_cents: int = DEFAULT_BASE_PRICE_CENTS,
) -> Optional[int]:
    """
    根据剩余天数计算价格(分)。

    规则: price = remaining_days / base_days * base_price
    """
    if remaining_days is None:
        return None
    if remaining_days <= 0:
        return 0

    numerator = int(remaining_days) * int(base_price_cents)
    # 四舍五入到“分”(整数)
    return int((2 * numerator + int(base_days)) // (2 * int(base_days)))


def format_price_yuan(price_cents: Optional[int]) -> Optional[str]:
    """将分格式化为元字符串(去掉无意义的 0)。"""
    if price_cents is None:
        return None
    yuan = price_cents / 100.0
    return f"{yuan:.2f}".rstrip("0").rstrip(".")

