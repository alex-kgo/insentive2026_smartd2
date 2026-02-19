"""
loguru 기반 로거 설정.
스크린샷 저장 헬퍼 포함.
"""
import sys
from datetime import datetime
from pathlib import Path
from loguru import logger

from config import LOG_DIR, SCREEN_DIR


def setup_logger(month: str) -> None:
    """
    month: 'YYYY-MM' 형식
    - 콘솔: INFO 이상
    - 파일: DEBUG 이상, 월별 로그 파일
    """
    logger.remove()

    # 콘솔
    logger.add(
        sys.stdout,
        level="INFO",
        format="<green>{time:HH:mm:ss}</green> | <level>{level:<7}</level> | {message}",
        colorize=True,
    )

    # 파일 (월별)
    log_file = LOG_DIR / f"run_{month}.log"
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logger.add(
        str(log_file),
        level="DEBUG",
        format="{time:YYYY-MM-DD HH:mm:ss} | {level:<7} | {name}:{line} | {message}",
        encoding="utf-8",
        rotation="50 MB",
        retention="30 days",
    )

    logger.info(f"로거 초기화 완료 - 로그 파일: {log_file}")


def save_screenshot(month: str, label: str) -> Path | None:
    """
    현재 화면을 스크린샷으로 저장.
    pyautogui가 없거나 실패해도 예외 없이 None 반환.
    """
    try:
        import pyautogui  # 선택적 의존성

        screen_dir = SCREEN_DIR / month.replace("-", "")
        screen_dir.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%H%M%S")
        path = screen_dir / f"{label}_{ts}.png"
        pyautogui.screenshot(str(path))
        logger.debug(f"스크린샷 저장: {path}")
        return path
    except Exception as e:
        logger.debug(f"스크린샷 저장 실패(무시): {e}")
        return None
