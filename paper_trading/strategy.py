from __future__ import annotations

from dataclasses import dataclass
from typing import Sequence


@dataclass(frozen=True)
class StrategyDecision:
    action: str
    reason: str
    short_ma: float | None = None
    long_ma: float | None = None


def moving_average(values: Sequence[float], window: int) -> float | None:
    if len(values) < window:
        return None
    return sum(values[-window:]) / window


def decide(
    closes: Sequence[float],
    position_qty: int,
    entry_price: float | None,
    current_price: float,
    short_window: int = 20,
    long_window: int = 60,
    stop_loss_pct: float = 0.03,
    take_profit_pct: float = 0.05,
) -> StrategyDecision:
    if position_qty > 0 and entry_price:
        return_pct = (current_price - entry_price) / entry_price
        if return_pct <= -stop_loss_pct:
            return StrategyDecision("SELL", f"손절 조건 도달 ({return_pct:.2%})")
        if return_pct >= take_profit_pct:
            return StrategyDecision("SELL", f"익절 조건 도달 ({return_pct:.2%})")

    if len(closes) < long_window + 1:
        return StrategyDecision("HOLD", f"5분봉 준비 중 ({len(closes)}/{long_window + 1})")

    previous = closes[:-1]
    prev_short = moving_average(previous, short_window)
    prev_long = moving_average(previous, long_window)
    now_short = moving_average(closes, short_window)
    now_long = moving_average(closes, long_window)

    if position_qty == 0 and prev_short is not None and prev_long is not None:
        if prev_short <= prev_long and now_short > now_long:
            return StrategyDecision("BUY", "20이평이 60이평을 상향 돌파", now_short, now_long)

    return StrategyDecision("HOLD", "조건 미충족", now_short, now_long)
