from __future__ import annotations

import asyncio
from dataclasses import asdict, dataclass
from datetime import datetime
import json
import time
from typing import Any
from zoneinfo import ZoneInfo

import websockets

from .config import PaperConfig
from .kis_client import KISPaperClient
from .state_store import StateStore
from .strategy import decide


KST = ZoneInfo("Asia/Seoul")


@dataclass
class Candle:
    start: datetime
    open: int
    high: int
    low: int
    close: int
    volume: int = 0

    def update(self, price: int, volume: int) -> None:
        self.high = max(self.high, price)
        self.low = min(self.low, price)
        self.close = price
        self.volume += max(volume, 0)

    def to_dict(self) -> dict[str, Any]:
        result = asdict(self)
        result["start"] = self.start.isoformat()
        return result


def candle_start(timestamp: datetime) -> datetime:
    return timestamp.replace(minute=(timestamp.minute // 5) * 5, second=0, microsecond=0)


class PaperTradingEngine:
    def __init__(self, config: PaperConfig, store: StateStore | None = None):
        self.config = config
        self.client = KISPaperClient(config)
        self.store = store or StateStore()
        self.candles: list[Candle] = []
        self.current: Candle | None = None
        self.position_qty = 0
        self.entry_price = 0.0
        self.pending_order = False
        self.pending_action = ""
        self.safety_lock_reason = ""
        self.last_signal = {"action": "HOLD", "reason": "시작 전"}
        self.events: list[dict[str, Any]] = []
        self.last_price = 0
        self.last_received_at: datetime | None = None
        self._restore_state()

    def _restore_state(self) -> None:
        saved = self.store.read()
        if not saved:
            return
        restored: list[Candle] = []
        for item in saved.get("candles", []):
            try:
                restored.append(
                    Candle(
                        start=datetime.fromisoformat(item["start"]),
                        open=int(item["open"]),
                        high=int(item["high"]),
                        low=int(item["low"]),
                        close=int(item["close"]),
                        volume=int(item.get("volume", 0)),
                    )
                )
            except (KeyError, TypeError, ValueError):
                continue
        self.candles = restored[-200:]
        self.events = list(saved.get("events", []))[-100:]
        self.pending_order = bool(saved.get("pending_order", False))
        self.pending_action = str(saved.get("pending_action", ""))
        if self.pending_order:
            self._event("SYSTEM", "재시작 전 미확인 주문이 있어 신규 주문을 잠급니다.")

    def _event(self, kind: str, message: str, **extra: Any) -> None:
        self.events.append({"time": datetime.now(KST).isoformat(), "kind": kind, "message": message, **extra})
        self.events = self.events[-100:]

    def _state(self, status: str = "RUNNING", error: str = "") -> dict[str, Any]:
        return {
            "status": status,
            "paper_only": True,
            "symbol": self.config.symbol,
            "symbol_name": self.config.symbol_name,
            "quantity_per_order": self.config.quantity,
            "updated_at": datetime.now(KST).isoformat(),
            "last_received_at": self.last_received_at.isoformat() if self.last_received_at else None,
            "last_price": self.last_price,
            "position_qty": self.position_qty,
            "entry_price": self.entry_price,
            "pending_order": self.pending_order,
            "pending_action": self.pending_action,
            "safety_lock_reason": self.safety_lock_reason,
            "strategy": "5분봉 MA20/MA60 골든크로스, -3% 손절, +5% 익절",
            "last_signal": self.last_signal,
            "candles": [c.to_dict() for c in self.candles[-80:]],
            "events": self.events[-50:],
            "error": error,
        }

    def _sync_position(self) -> None:
        position = self.client.position()
        self.position_qty = position["quantity"]
        self.entry_price = position["average_price"]
        if self.position_qty > self.config.quantity:
            self.safety_lock_reason = f"보유 수량이 안전 한도({self.config.quantity}주)를 초과했습니다."
        if self.pending_action == "BUY" and self.position_qty > 0:
            self.pending_order = False
            self.pending_action = ""
            self._event("FILL", "모의 매수 체결을 잔고에서 확인")
        elif self.pending_action == "SELL" and self.position_qty == 0:
            self.pending_order = False
            self.pending_action = ""
            self._event("FILL", "모의 매도 체결을 잔고에서 확인")

    def _handle_completed_candle(self) -> None:
        if not self.candles:
            return
        closes = [c.close for c in self.candles]
        decision = decide(
            closes,
            self.position_qty,
            self.entry_price,
            closes[-1],
            self.config.short_window,
            self.config.long_window,
            self.config.stop_loss_pct,
            self.config.take_profit_pct,
        )
        self.last_signal = {
            "action": decision.action,
            "reason": decision.reason,
            "time": datetime.now(KST).isoformat(),
            "short_ma": decision.short_ma,
            "long_ma": decision.long_ma,
        }
        self._event("SIGNAL", f"{decision.action}: {decision.reason}")
        if decision.action in ("BUY", "SELL") and not self.pending_order:
            self._place_order(decision.action)

    def _handle_live_exit(self, price: int) -> None:
        if self.position_qty <= 0 or not self.entry_price or self.pending_order or self.safety_lock_reason:
            return
        return_pct = (price - self.entry_price) / self.entry_price
        if return_pct <= -self.config.stop_loss_pct:
            self.last_signal = {
                "action": "SELL",
                "reason": f"실시간 손절 조건 도달 ({return_pct:.2%})",
                "time": datetime.now(KST).isoformat(),
            }
            self._event("SIGNAL", self.last_signal["reason"])
            self._place_order("SELL")
        elif return_pct >= self.config.take_profit_pct:
            self.last_signal = {
                "action": "SELL",
                "reason": f"실시간 익절 조건 도달 ({return_pct:.2%})",
                "time": datetime.now(KST).isoformat(),
            }
            self._event("SIGNAL", self.last_signal["reason"])
            self._place_order("SELL")

    def _place_order(self, action: str) -> None:
        if self.safety_lock_reason:
            self._event("ERROR", f"안전 잠금으로 주문 거부: {self.safety_lock_reason}")
            return
        if action == "BUY" and self.position_qty > 0:
            return
        if action == "SELL" and self.position_qty <= 0:
            return
        self.pending_order = True
        self.pending_action = action
        started = time.perf_counter()
        try:
            response = self.client.order_market(action.lower(), self.config.quantity)
            output = response.get("output", {})
            latency_ms = round((time.perf_counter() - started) * 1000, 1)
            self._event(
                "ORDER",
                f"모의 {action} 시장가 {self.config.quantity}주 주문 접수",
                order_no=output.get("ODNO"),
                latency_ms=latency_ms,
            )
            # 주문 응답은 체결이 아니므로 실제 잔고가 바뀔 때까지 다음 봉에서 재동기화한다.
        except Exception as exc:
            self._event("ERROR", f"주문 실패: {exc}")
            self.pending_order = False
            self.pending_action = ""
        finally:
            try:
                self._sync_position()
            except Exception as exc:
                self._event("ERROR", f"주문 후 잔고 확인 실패: {exc}")

    def _on_trade(self, trade_time: str, price: int, volume: int) -> None:
        now = datetime.now(KST)
        try:
            tick_time = now.replace(
                hour=int(trade_time[0:2]), minute=int(trade_time[2:4]), second=int(trade_time[4:6]), microsecond=0
            )
        except (ValueError, IndexError):
            tick_time = now
        start = candle_start(tick_time)
        self.last_price = price
        self.last_received_at = now
        self._handle_live_exit(price)

        if self.current is None:
            self.current = Candle(start, price, price, price, price, volume)
        elif start > self.current.start:
            self.candles.append(self.current)
            self.candles = self.candles[-200:]
            self._sync_position()
            self._handle_completed_candle()
            self.current = Candle(start, price, price, price, price, volume)
        elif start == self.current.start:
            self.current.update(price, volume)
        self.store.write(self._state())

    async def run(self) -> None:
        self._sync_position()
        approval_key = self.client.approval_key()
        subscribe = {
            "header": {"approval_key": approval_key, "custtype": "P", "tr_type": "1", "content-type": "utf-8"},
            "body": {"input": {"tr_id": "H0STCNT0", "tr_key": self.config.symbol}},
        }
        self._event("SYSTEM", "한국투자증권 모의투자 엔진 시작")
        self.store.write(self._state("CONNECTING"))
        retry = 0
        while True:
            try:
                async with websockets.connect(self.config.websocket_url, ping_interval=30, ping_timeout=10) as ws:
                    await ws.send(json.dumps(subscribe))
                    retry = 0
                    self._event("SYSTEM", "실시간 체결가 WebSocket 연결")
                    async for raw in ws:
                        if raw.startswith(("0|", "1|")):
                            parts = raw.split("|", 3)
                            if len(parts) == 4 and parts[1] == "H0STCNT0":
                                fields = parts[3].split("^")
                                if len(fields) > 12:
                                    self._on_trade(fields[1], int(fields[2]), int(fields[12] or 0))
                        else:
                            message = json.loads(raw)
                            if message.get("header", {}).get("tr_id") == "PINGPONG":
                                await ws.send(raw)
            except asyncio.CancelledError:
                self.store.write(self._state("STOPPED"))
                raise
            except Exception as exc:
                retry += 1
                self._event("ERROR", f"WebSocket 오류, 재연결 {retry}회: {exc}")
                self.store.write(self._state("RECONNECTING", str(exc)))
                await asyncio.sleep(min(2**retry, 30))
