from __future__ import annotations

import time
from typing import Any

import requests

from .config import PaperConfig


class KISAPIError(RuntimeError):
    pass


class KISPaperClient:
    """모의투자 도메인만 사용하는 최소 KIS REST 클라이언트."""

    def __init__(self, config: PaperConfig):
        if "openapivts" not in config.rest_url:
            raise ValueError("안전상 모의투자(openapivts) 도메인만 허용합니다.")
        self.config = config
        self.session = requests.Session()
        self._token = ""
        self._expires_at = 0.0

    def access_token(self) -> str:
        if self._token and time.time() < self._expires_at - 60:
            return self._token
        response = self.session.post(
            f"{self.config.rest_url}/oauth2/tokenP",
            json={
                "grant_type": "client_credentials",
                "appkey": self.config.app_key,
                "appsecret": self.config.app_secret,
            },
            timeout=10,
        )
        data = self._json(response, "접근토큰 발급")
        self._token = data["access_token"]
        self._expires_at = time.time() + int(data.get("expires_in", 3600))
        return self._token

    def approval_key(self) -> str:
        response = self.session.post(
            f"{self.config.rest_url}/oauth2/Approval",
            json={"grant_type": "client_credentials", "appkey": self.config.app_key, "secretkey": self.config.app_secret},
            timeout=10,
        )
        return self._json(response, "WebSocket 접속키 발급")["approval_key"]

    def _headers(self, tr_id: str) -> dict[str, str]:
        return {
            "content-type": "application/json; charset=utf-8",
            "authorization": f"Bearer {self.access_token()}",
            "appkey": self.config.app_key,
            "appsecret": self.config.app_secret,
            "tr_id": tr_id,
            "custtype": "P",
        }

    @staticmethod
    def _json(response: requests.Response, action: str) -> dict[str, Any]:
        try:
            data = response.json()
        except ValueError as exc:
            raise KISAPIError(f"{action} 실패: HTTP {response.status_code}") from exc
        if not response.ok or str(data.get("rt_cd", "0")) not in ("0", ""):
            raise KISAPIError(f"{action} 실패: {data.get('msg1') or data}")
        return data

    def order_market(self, side: str, quantity: int = 1) -> dict[str, Any]:
        if side not in ("buy", "sell"):
            raise ValueError("side는 buy 또는 sell이어야 합니다.")
        tr_id = "VTTC0012U" if side == "buy" else "VTTC0011U"
        payload = {
            "CANO": self.config.account_no,
            "ACNT_PRDT_CD": self.config.account_product_code,
            "PDNO": self.config.symbol,
            "ORD_DVSN": "01",
            "ORD_QTY": str(quantity),
            "ORD_UNPR": "0",
            "EXCG_ID_DVSN_CD": "KRX",
            "SLL_TYPE": "01" if side == "sell" else "",
            "CNDT_PRIC": "",
        }
        response = self.session.post(
            f"{self.config.rest_url}/uapi/domestic-stock/v1/trading/order-cash",
            headers=self._headers(tr_id),
            json=payload,
            timeout=10,
        )
        return self._json(response, "모의주문")

    def position(self) -> dict[str, Any]:
        params = {
            "CANO": self.config.account_no,
            "ACNT_PRDT_CD": self.config.account_product_code,
            "AFHR_FLPR_YN": "N",
            "OFL_YN": "",
            "INQR_DVSN": "02",
            "UNPR_DVSN": "01",
            "FUND_STTL_ICLD_YN": "N",
            "FNCG_AMT_AUTO_RDPT_YN": "N",
            "PRCS_DVSN": "00",
            "CTX_AREA_FK100": "",
            "CTX_AREA_NK100": "",
        }
        response = self.session.get(
            f"{self.config.rest_url}/uapi/domestic-stock/v1/trading/inquire-balance",
            headers=self._headers("VTTC8434R"),
            params=params,
            timeout=10,
        )
        data = self._json(response, "모의잔고 조회")
        for item in data.get("output1", []):
            if str(item.get("pdno", "")) == self.config.symbol:
                return {
                    "quantity": int(float(item.get("hldg_qty", 0) or 0)),
                    "average_price": float(item.get("pchs_avg_pric", 0) or 0),
                }
        return {"quantity": 0, "average_price": 0.0}
