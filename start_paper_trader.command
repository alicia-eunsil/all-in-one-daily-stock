#!/bin/bash

set -u

TRADER_DIR="/Users/alicia/Desktop/#python/all-in-one-daily-stock"
ENV_FILE="$TRADER_DIR/.env.paper"
LOG_DIR="$TRADER_DIR/logs"
LOG_FILE="$LOG_DIR/paper_trader.log"

cd "$TRADER_DIR" || {
  echo "프로그램 폴더를 찾지 못했습니다: $TRADER_DIR"
  read -r -p "Enter를 누르면 닫힙니다."
  exit 1
}

if [ ! -f "$ENV_FILE" ]; then
  echo "설정 파일이 없습니다: $ENV_FILE"
  echo ".env.paper.example을 복사해 .env.paper를 만들고 값을 입력해 주세요."
  read -r -p "Enter를 누르면 닫힙니다."
  exit 1
fi

set -a
# shellcheck disable=SC1090
source "$ENV_FILE"
set +a

if [[ "${KIS_PAPER_APP_KEY:-}" == *"여기에"* ]] || \
   [[ "${KIS_PAPER_APP_SECRET:-}" == *"여기에"* ]] || \
   [[ ! "${KIS_PAPER_ACCOUNT_NO:-}" =~ ^[0-9]{8}$ ]]; then
  echo "한국투자증권 모의투자 설정을 먼저 입력해야 합니다."
  echo "지금 설정 파일을 TextEdit으로 엽니다. 값을 입력하고 저장한 뒤 아이콘을 다시 실행하세요."
  open -e "$ENV_FILE"
  read -r -p "Enter를 누르면 이 창이 닫힙니다."
  exit 1
fi

mkdir -p "$LOG_DIR"

echo "============================================================"
echo " 두산로보틱스 한국투자증권 모의매매를 시작합니다."
echo " 이 창을 닫거나 Ctrl+C를 누르면 프로그램이 종료됩니다."
echo " 로그: $LOG_FILE"
echo "============================================================"

python3 "$TRADER_DIR/paper_trader.py" 2>&1 | tee -a "$LOG_FILE"
EXIT_CODE=${PIPESTATUS[0]}

echo
echo "프로그램이 종료되었습니다. 종료 코드: $EXIT_CODE"
read -r -p "Enter를 누르면 창이 닫힙니다."
exit "$EXIT_CODE"
