"""
월 단위 실행 진행 상태를 로컬 JSON으로 저장/복구한다. (D25)

파일 위치: C:/RPA/logi_exports/logs/checkpoint_{YYYY-MM}.json
구조:
{
  "month": "2026-02",
  "done_dates": ["2026-02-01", "2026-02-03", ...],
  "failed_dates": ["2026-02-02", ...],
  "last_csv": "logi_calls_2026-02_20260219-0630.csv",
  "telegram_sent": false
}
"""
import json
from pathlib import Path
from loguru import logger

from config import LOG_DIR


def _checkpoint_path(month: str) -> Path:
    return LOG_DIR / f"checkpoint_{month}.json"


def load(month: str) -> dict:
    """체크포인트 로드. 없으면 빈 상태 반환."""
    path = _checkpoint_path(month)
    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            logger.info(f"체크포인트 로드: {path} (완료={len(data.get('done_dates', []))}일)")
            return data
        except Exception as e:
            logger.warning(f"체크포인트 파싱 실패, 초기화: {e}")

    return {
        "month": month,
        "done_dates": [],
        "failed_dates": [],
        "last_csv": None,
        "telegram_sent": False,
    }


def save(state: dict) -> None:
    """체크포인트 저장."""
    path = _checkpoint_path(state["month"])
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    logger.debug(f"체크포인트 저장: {path}")


def mark_done(state: dict, date_str: str) -> dict:
    """날짜를 완료 목록에 추가하고 실패 목록에서 제거."""
    if date_str not in state["done_dates"]:
        state["done_dates"].append(date_str)
    if date_str in state["failed_dates"]:
        state["failed_dates"].remove(date_str)
    save(state)
    return state


def mark_failed(state: dict, date_str: str) -> dict:
    """날짜를 실패 목록에 추가 (완료 목록에서는 제거하지 않음)."""
    if date_str not in state["failed_dates"]:
        state["failed_dates"].append(date_str)
    save(state)
    return state


def is_done(state: dict, date_str: str) -> bool:
    return date_str in state["done_dates"]


def pending_dates(all_dates: list[str], state: dict) -> list[str]:
    """완료되지 않은 날짜 목록 반환 (실패 포함)."""
    done = set(state["done_dates"])
    return [d for d in all_dates if d not in done]
