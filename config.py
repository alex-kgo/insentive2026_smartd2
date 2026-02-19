"""
시스템 전역 상수/경로 설정
"""
from pathlib import Path

# ── 운영 폴더 경로 ────────────────────────────────────────────────────────────
BASE_DIR = Path(r"D:\_claude\인센티브_로지데이터")
RAW_DIR = BASE_DIR / "raw"
PROCESSED_DIR = BASE_DIR / "processed"
ERROR_DIR = BASE_DIR / "error"
CSV_DIR = BASE_DIR / "csv"
LOG_DIR = BASE_DIR / "logs"
SCREEN_DIR = LOG_DIR / "screens"

# ── 로지 UI 설정 ──────────────────────────────────────────────────────────────
LOGI_WINDOW_TITLE_RE = r".*아리랑.*|.*SMART.*|.*스마트D2.*"  # 메인 창 title_re
LOGI_MENU_EMPLOYEE   = "직원"             # 상단 메뉴명
LOGI_SCREEN_NAME     = "기간별수신콜수"   # 메뉴 클릭 후 진입할 화면명
LOGI_QUERY_WAIT_SEC  = 5                 # 조회 버튼 클릭 후 초기 대기(초)
LOGI_POLL_INTERVAL_SEC = 1.0             # 조회 완료 감지 폴링 간격
LOGI_POLL_MAX_SEC    = 60                # 조회 완료 최대 대기 시간
CHECKBOX_LABEL       = "전화받은건수기준" # 체크박스 레이블 (정확한 텍스트)
CHECKBOX_TARGET_STATE = True             # 체크박스 목표 상태 (True=체크)

# ── 기간 입력 형식 (D9) ───────────────────────────────────────────────────────
DATE_FMT = "%Y-%m-%d"                    # 날짜 문자열 포맷
PERIOD_FMT = "%Y-%m-%d 00:00"           # 로지 기간 필드 입력 포맷

# ── 엑셀 파싱 컬럼 인덱스 (0-based, A=0) ────────────────────────────────────
COL_CODE = 0    # A: 코드
COL_NAME = 1    # B: 성명
COL_C    = 2    # C: 고객(받음)
COL_D    = 3    # D: 기사(받음)
COL_E    = 4    # E: 고객(걸음)
COL_F    = 5    # F: 기사(걸음)
EXCEL_HEADER_ROWS = 1                    # 건너뛸 헤더 행 수

# ── Google Sheets 헤더 ───────────────────────────────────────────────────────
SHEET_HEADERS = ["날짜", "코드", "성명", "수신 합계", "발신 합계", "총합계"]

# ── CSV 파일명 패턴 ───────────────────────────────────────────────────────────
CSV_FILENAME_FMT = "logi_calls_{month}_{ts}.csv"

# ── Telegram 전송 ─────────────────────────────────────────────────────────────
TELEGRAM_MAX_RETRIES = 3
TELEGRAM_BACKOFF_BASE = 2               # 지수 백오프 밑수(초)
TELEGRAM_API_URL = "https://api.telegram.org/bot{token}/{method}"
