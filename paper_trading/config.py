from __future__ import annotations

from dataclasses import dataclass
import os


@dataclass(frozen=True)
class PaperConfig:
    app_key: str
    app_secret: str
    account_no: str
    account_product_code: str = "01"
    symbol: str = "454910"
    symbol_name: str = "두산로보틱스"
    quantity: int = 1
    short_window: int = 20
    long_window: int = 60
    stop_loss_pct: float = 0.03
    take_profit_pct: float = 0.05
    rest_url: str = "https://openapivts.koreainvestment.com:29443"
    websocket_url: str = "ws://ops.koreainvestment.com:31000"

    @classmethod
    def from_env(cls) -> "PaperConfig":
        values = {
            "app_key": os.getenv("KIS_PAPER_APP_KEY", "").strip(),
            "app_secret": os.getenv("KIS_PAPER_APP_SECRET", "").strip(),
            "account_no": os.getenv("KIS_PAPER_ACCOUNT_NO", "").strip(),
            "account_product_code": os.getenv("KIS_PAPER_ACCOUNT_PRODUCT_CODE", "01").strip(),
        }
        missing = [key for key in ("app_key", "app_secret", "account_no") if not values[key]]
        if missing:
            names = ", ".join(f"KIS_PAPER_{name.upper()}" for name in missing)
            raise ValueError(f"필수 환경변수가 없습니다: {names}")
        if len(values["account_no"]) != 8 or not values["account_no"].isdigit():
            raise ValueError("KIS_PAPER_ACCOUNT_NO에는 계좌번호 앞 8자리만 입력하세요.")
        return cls(**values)
