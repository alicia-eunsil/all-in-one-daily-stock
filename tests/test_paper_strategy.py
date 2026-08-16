from datetime import datetime
from zoneinfo import ZoneInfo

from paper_trading.engine import candle_start
from paper_trading.config import PaperConfig
from paper_trading.kis_client import KISPaperClient
from paper_trading.strategy import decide


def test_golden_cross_generates_buy():
    closes = [100.0] * 60
    closes[-20:] = [99.0] * 20
    closes.append(200.0)
    result = decide(closes, 0, None, closes[-1])
    assert result.action == "BUY"


def test_stop_loss_and_take_profit():
    assert decide([], 1, 100.0, 96.9).action == "SELL"
    assert decide([], 1, 100.0, 105.1).action == "SELL"


def test_five_minute_bucket():
    dt = datetime(2026, 8, 16, 10, 7, 45, tzinfo=ZoneInfo("Asia/Seoul"))
    assert candle_start(dt).minute == 5
    assert candle_start(dt).second == 0


def test_real_domain_is_rejected():
    config = PaperConfig("key", "secret", "12345678", rest_url="https://openapi.koreainvestment.com:9443")
    try:
        KISPaperClient(config)
    except ValueError as exc:
        assert "모의투자" in str(exc)
    else:
        raise AssertionError("실전 도메인이 차단되지 않았습니다.")
