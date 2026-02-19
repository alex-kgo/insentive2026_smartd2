"""
환경변수 로드 및 자격증명 관리.
.env 파일 또는 OS 환경변수에서 값을 읽는다.
"""
import os
from pathlib import Path
from dotenv import load_dotenv
from loguru import logger


def load_env(env_file: str | Path = ".env") -> None:
    """
    .env 파일 로드. 파일이 없으면 OS 환경변수만 사용.
    """
    env_path = Path(env_file)
    if env_path.exists():
        load_dotenv(env_path, override=False)
        logger.debug(f".env 로드: {env_path.resolve()}")
    else:
        logger.warning(f".env 파일 없음 — OS 환경변수만 사용: {env_path.resolve()}")


def _require(key: str) -> str:
    val = os.getenv(key)
    if not val:
        raise EnvironmentError(f"필수 환경변수 누락: {key}")
    return val


def get_logi_credentials() -> tuple[str, str]:
    """(id, password)"""
    return _require("LOGI_ID"), _require("LOGI_PW")


def get_spreadsheet_id() -> str:
    return _require("SPREADSHEET_ID")


def get_google_sa_json_path() -> Path:
    path = Path(_require("GOOGLE_SA_JSON_PATH"))
    if not path.exists():
        raise FileNotFoundError(f"Google SA JSON 없음: {path}")
    return path


def get_telegram_credentials() -> tuple[str, str]:
    """(bot_token, chat_id)"""
    return _require("TELEGRAM_BOT_TOKEN"), _require("TELEGRAM_CHAT_ID")
