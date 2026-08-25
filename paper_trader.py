"""노트북에서 실행하는 한국투자증권 모의투자 엔진."""

import asyncio

from paper_trading.config import PaperConfig
from paper_trading.engine import PaperTradingEngine


def main() -> None:
    config = PaperConfig.from_env()
    print("두산로보틱스 모의투자 엔진을 시작합니다. 종료: Ctrl+C")
    asyncio.run(PaperTradingEngine(config).run())


if __name__ == "__main__":
    main()
