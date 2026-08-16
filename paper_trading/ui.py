from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from .state_store import StateStore


def render_paper_trading_tab() -> None:
    st.subheader("🧪 한국투자증권 모의투자")
    st.caption("두산로보틱스(454910) · 1주 · 완성된 5분봉 · 모의투자 전용")
    st.info("매매 엔진은 노트북의 `paper_trader.py`에서 실행됩니다. 이 탭은 상태와 기록을 확인하는 화면입니다.")

    state = StateStore().read()
    if not state:
        st.warning("아직 모의투자 상태가 없습니다. 노트북에서 `python3 paper_trader.py`를 실행하세요.")
        with st.expander("필요한 환경변수"):
            st.code(
                "KIS_PAPER_APP_KEY\nKIS_PAPER_APP_SECRET\nKIS_PAPER_ACCOUNT_NO (앞 8자리)\n"
                "KIS_PAPER_ACCOUNT_PRODUCT_CODE=01\nFIREBASE_DATABASE_URL (선택)\nFIREBASE_DATABASE_TOKEN (선택)"
            )
        return

    updated_at = state.get("updated_at")
    stale = False
    if updated_at:
        try:
            stale = (datetime.now(ZoneInfo("Asia/Seoul")) - datetime.fromisoformat(updated_at)).total_seconds() > 30
        except ValueError:
            pass

    cols = st.columns(5)
    cols[0].metric("상태", "지연/중단" if stale else state.get("status", "-"))
    cols[1].metric("현재가", f"{int(state.get('last_price') or 0):,}원")
    cols[2].metric("보유", f"{int(state.get('position_qty') or 0)}주")
    cols[3].metric("평균 매수가", f"{float(state.get('entry_price') or 0):,.0f}원")
    cols[4].metric("마지막 신호", state.get("last_signal", {}).get("action", "-"))

    if stale:
        st.error("마지막 상태 갱신 후 30초가 지났습니다. 노트북·인터넷·매매 엔진을 확인하세요.")
    if state.get("error"):
        st.error(state["error"])
    if state.get("safety_lock_reason"):
        st.error(f"안전 잠금: {state['safety_lock_reason']}")

    st.markdown(f"**전략:** {state.get('strategy', '-')}")
    if state.get("pending_order"):
        st.warning(f"{state.get('pending_action', '')} 주문이 접수되어 체결 확인을 기다리고 있습니다.")
    signal = state.get("last_signal", {})
    st.write(f"최근 판단: **{signal.get('action', '-')}** — {signal.get('reason', '-')}")

    candles = pd.DataFrame(state.get("candles", []))
    if not candles.empty:
        candles["start"] = pd.to_datetime(candles["start"])
        st.markdown("#### 최근 완성 5분봉")
        st.dataframe(candles.tail(20), use_container_width=True, hide_index=True)
    else:
        st.caption("완성된 5분봉을 기다리는 중입니다.")

    events = pd.DataFrame(state.get("events", []))
    if not events.empty:
        st.markdown("#### 신호·주문·오류 기록")
        st.dataframe(events.iloc[::-1], use_container_width=True, hide_index=True)

    st.warning("모의투자 주문도 자동으로 제출됩니다. 실행 전 계좌번호와 모의투자 앱 키를 다시 확인하세요.")
