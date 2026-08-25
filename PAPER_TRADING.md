# 한국투자증권 모의투자

두산로보틱스(`454910`) 1주를 대상으로 하는 로컬 모의매매 엔진입니다. 실전투자 도메인은 코드에서 차단합니다.

## 전략

- 완성된 5분봉만 사용
- 20개 이동평균이 60개 이동평균을 상향 돌파하면 시장가 1주 매수
- 실시간 체결가가 평균 매수가 대비 3% 하락하면 시장가 1주 매도
- 실시간 체결가가 평균 매수가 대비 5% 상승하면 시장가 1주 매도
- 최대 보유 수량 1주, 미확인 주문이 있으면 추가 주문 금지

첫 실행에는 이동평균 계산을 위해 완성된 5분봉 61개가 필요합니다. 이후 봉은 `data/paper_trading_state.json`에 저장되어 재시작 시 복구됩니다.

## 설치

```bash
python3 -m pip install -r requirements.txt
```

## 환경변수

```bash
export KIS_PAPER_APP_KEY="모의투자 앱키"
export KIS_PAPER_APP_SECRET="모의투자 앱시크릿"
export KIS_PAPER_ACCOUNT_NO="계좌번호 앞 8자리"
export KIS_PAPER_ACCOUNT_PRODUCT_CODE="01"
```

Streamlit Cloud에서 노트북의 상태를 보려면 Firebase Realtime Database를 함께 설정합니다.

```bash
export FIREBASE_DATABASE_URL="https://프로젝트.firebaseio.com"
export FIREBASE_DATABASE_TOKEN="데이터베이스 접근 토큰"
```

동일한 Firebase 값은 Streamlit Cloud의 Secrets에도 등록해야 합니다. Firebase를 설정하지 않으면 상태는 노트북의 로컬 JSON에만 저장됩니다.

## 실행

```bash
python3 paper_trader.py
```

macOS에서는 바탕화면의 `두산로보틱스_모의투자.command`를 더블클릭해도 됩니다. 이 아이콘은 `.env.paper`의 설정을 읽고 로그를 `logs/paper_trader.log`에 저장합니다.

종료는 `Ctrl+C`입니다. 실행 전 한국투자증권 모의투자 계좌에 기존 보유 종목이나 미체결 주문이 없는지 확인하세요.

## 현재 안전 제한

- 모의투자 REST 주소(`openapivts`)만 허용
- 두산로보틱스 1주만 주문
- 주문 응답을 체결로 간주하지 않고 실제 잔고로 재확인
- 재시작 전 주문이 미확인 상태이면 신규 주문 잠금
- Firebase 장애가 나도 매매 엔진은 계속 실행하고 로컬 상태를 보존

실전투자에는 사용하지 마세요. 정정·취소 및 미체결 주문 상세조회는 다음 단계에서 별도로 검증해야 합니다.
