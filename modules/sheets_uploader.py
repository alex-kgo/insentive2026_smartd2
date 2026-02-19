"""
Google Sheets에 파싱 데이터를 Upsert한다. (D18~D-NEW3)

- 시트명: YYYY-MM
- 유니크 키: (날짜, 코드)
- 동일 키 존재 시 → 업데이트, 없으면 → append
- 멱등성 보장: 재실행해도 데이터 중복 없음
"""
from pathlib import Path
from loguru import logger

import gspread
from google.oauth2.service_account import Credentials

from config import SHEET_HEADERS

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _build_client(sa_json_path: Path) -> gspread.Client:
    creds = Credentials.from_service_account_file(str(sa_json_path), scopes=_SCOPES)
    return gspread.authorize(creds)


def _get_or_create_sheet(spreadsheet: gspread.Spreadsheet, month: str) -> gspread.Worksheet:
    """시트(탭) 가져오기. 없으면 생성 후 헤더 작성."""
    try:
        ws = spreadsheet.worksheet(month)
        logger.debug(f"기존 시트 사용: {month}")
        return ws
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=month, rows=5000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS, value_input_option="RAW")
        logger.info(f"새 시트 생성: {month}")
        return ws


def _row_to_values(row: dict) -> list:
    """dict → 시트 행 순서 리스트 [날짜, 코드, 성명, 수신합계, 발신합계, 총합계]."""
    return [
        row["날짜"],
        row["코드"],
        row["성명"],
        row["수신합계"],
        row["발신합계"],
        row["총합계"],
    ]


def upsert_rows(
    sa_json_path: Path,
    spreadsheet_id: str,
    month: str,
    rows: list[dict],
) -> int:
    """
    Args:
        sa_json_path: 서비스 계정 JSON 경로
        spreadsheet_id: Google Spreadsheet ID
        month: 'YYYY-MM'
        rows: excel_parser.parse_open_excel() 반환값

    Returns:
        upsert된 행 수
    """
    if not rows:
        logger.info(f"[{month}] upsert 대상 없음")
        return 0

    client = _build_client(sa_json_path)
    spreadsheet = client.open_by_key(spreadsheet_id)
    ws = _get_or_create_sheet(spreadsheet, month)

    # 현재 시트 전체 읽기 (헤더 제외)
    all_values = ws.get_all_values()
    if not all_values:
        existing_data: list[list] = []
        data_start_row = 2  # 헤더가 없는 경우
    else:
        existing_data = all_values[1:]  # 헤더 제외
        data_start_row = 2

    # (날짜, 코드) → 시트 행 인덱스(1-based) 맵 작성
    key_to_row: dict[tuple, int] = {}
    for i, row_vals in enumerate(existing_data):
        if len(row_vals) >= 2:
            key = (row_vals[0], row_vals[1])  # (날짜, 코드)
            key_to_row[key] = data_start_row + i

    batch_updates: list[dict] = []  # gspread batch_update용
    appends: list[list] = []

    for row in rows:
        key = (row["날짜"], row["코드"])
        values = _row_to_values(row)

        if key in key_to_row:
            sheet_row = key_to_row[key]
            # A열~F열 업데이트 (1-based col 1~6)
            cell_range = f"A{sheet_row}:F{sheet_row}"
            batch_updates.append({
                "range": cell_range,
                "values": [values],
            })
        else:
            appends.append(values)

    upserted = 0

    # 업데이트 배치 실행
    if batch_updates:
        ws.batch_update(batch_updates, value_input_option="RAW")
        upserted += len(batch_updates)
        logger.debug(f"[{month}] 업데이트 {len(batch_updates)}행")

    # 신규 append
    if appends:
        ws.append_rows(appends, value_input_option="RAW")
        upserted += len(appends)
        logger.debug(f"[{month}] 신규 추가 {len(appends)}행")

    logger.info(f"[{month}] Sheets upsert 완료 — 총 {upserted}행")
    return upserted


def read_all_rows(
    sa_json_path: Path,
    spreadsheet_id: str,
    month: str,
) -> list[list]:
    """
    월 시트의 전체 데이터(헤더 제외)를 반환.
    CSV export에서 사용.
    """
    client = _build_client(sa_json_path)
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        ws = spreadsheet.worksheet(month)
    except gspread.WorksheetNotFound:
        logger.warning(f"시트 없음: {month}")
        return []

    all_values = ws.get_all_values()
    if len(all_values) <= 1:
        return []
    return all_values[1:]  # 헤더 제외


# ── 단독 실행 테스트 ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent.parent))
    from utils.logger import setup_logger
    from utils.secrets import load_env, get_spreadsheet_id, get_google_sa_json_path

    setup_logger("TEST")
    load_env()

    test_rows = [
        {"날짜": "2026-02-01", "코드": "T001", "성명": "테스트", "수신합계": 10, "발신합계": 5, "총합계": 15},
        {"날짜": "2026-02-01", "코드": "T002", "성명": "홍길동", "수신합계": 8, "발신합계": 3, "총합계": 11},
    ]

    n = upsert_rows(
        sa_json_path=get_google_sa_json_path(),
        spreadsheet_id=get_spreadsheet_id(),
        month="2026-02",
        rows=test_rows,
    )
    print(f"upsert 행 수: {n}")
