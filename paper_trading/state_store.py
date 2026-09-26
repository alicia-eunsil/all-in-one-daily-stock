from __future__ import annotations

import json
import os
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

import requests


DEFAULT_STATE_PATH = Path(__file__).resolve().parent.parent / "data" / "paper_trading_state.json"


class StateStore:
    def __init__(self, path: Path | None = None):
        self.path = path or DEFAULT_STATE_PATH
        self.firebase_url = os.getenv("FIREBASE_DATABASE_URL", "").rstrip("/")
        self.firebase_token = os.getenv("FIREBASE_DATABASE_TOKEN", "")

    def _firebase_endpoint(self) -> str:
        endpoint = f"{self.firebase_url}/paper_trading/state.json"
        if self.firebase_token:
            endpoint += f"?auth={self.firebase_token}"
        return endpoint

    def write(self, state: dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile("w", encoding="utf-8", dir=self.path.parent, delete=False) as tmp:
            json.dump(state, tmp, ensure_ascii=False, indent=2)
            temp_path = Path(tmp.name)
        temp_path.replace(self.path)

        if self.firebase_url:
            try:
                requests.put(self._firebase_endpoint(), json=state, timeout=5).raise_for_status()
            except requests.RequestException:
                # 로컬 매매 엔진은 Firebase 장애 때문에 멈추지 않는다.
                pass

    def read(self) -> dict[str, Any]:
        if self.firebase_url:
            try:
                response = requests.get(self._firebase_endpoint(), timeout=5)
                response.raise_for_status()
                remote = response.json()
                if isinstance(remote, dict):
                    return remote
            except requests.RequestException:
                pass

        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return {}
