"""
Telegram으로 CSV 파일을 Document로 전송한다. (D23~D24)

- 최대 3회 재시도, 지수 백오프
- 최종 실패 시 CSV 로컬 보관 + 로그 기록
"""
import time
from pathlib import Path

import requests
from loguru import logger

from config import (
    TELEGRAM_MAX_RETRIES,
    TELEGRAM_BACKOFF_BASE,
    TELEGRAM_API_URL,
)


def send_csv(
    bot_token: str,
    chat_id: str,
    csv_path: Path,
    month: str,
    total_rows: int,
) -> bool:
    """
    Args:
        bot_token: Telegram Bot Token
        chat_id: 전송 대상 Chat ID
        csv_path: 전송할 CSV 파일 경로
        month: 'YYYY-MM'
        total_rows: 총 데이터 행 수 (메시지 표시용)

    Returns:
        True(성공) / False(최종 실패)
    """
    caption = (
        f"[로지 월 취합 완료]\n"
        f"월: {month}\n"
        f"총 행수: {total_rows:,}\n"
        f"상태: SUCCESS"
    )

    url = TELEGRAM_API_URL.format(token=bot_token, method="sendDocument")

    for attempt in range(1, TELEGRAM_MAX_RETRIES + 1):
        try:
            with csv_path.open("rb") as f:
                resp = requests.post(
                    url,
                    data={"chat_id": chat_id, "caption": caption},
                    files={"document": (csv_path.name, f, "text/csv")},
                    timeout=60,
                )
            resp.raise_for_status()
            data = resp.json()
            if data.get("ok"):
                logger.info(f"Telegram 전송 성공: {csv_path.name} (시도 {attempt}회)")
                return True
            else:
                raise RuntimeError(f"Telegram API 오류: {data}")

        except Exception as e:
            wait = TELEGRAM_BACKOFF_BASE ** attempt
            logger.warning(f"Telegram 전송 실패 (시도 {attempt}/{TELEGRAM_MAX_RETRIES}): {e}")
            if attempt < TELEGRAM_MAX_RETRIES:
                logger.info(f"{wait}초 후 재시도...")
                time.sleep(wait)

    logger.error(f"Telegram 최종 전송 실패 — CSV 로컬 보관: {csv_path}")
    return False


# ── 단독 실행 테스트 ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent.parent))
    from utils.logger import setup_logger
    from utils.secrets import load_env, get_telegram_credentials

    setup_logger("TEST")
    load_env()

    token, chat_id = get_telegram_credentials()

    # 테스트용 더미 CSV
    test_csv = __import__("pathlib").Path("C:/RPA/logi_exports/csv/test.csv")
    test_csv.parent.mkdir(parents=True, exist_ok=True)
    test_csv.write_text("날짜,코드,성명,수신합계,발신합계,총합계\n2026-02-01,T001,테스트,10,5,15\n", encoding="utf-8-sig")

    ok = send_csv(token, chat_id, test_csv, "2026-02", 1)
    print(f"전송 결과: {'성공' if ok else '실패'}")
