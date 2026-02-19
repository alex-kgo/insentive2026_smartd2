"""
Google Sheets 월 시트 데이터를 CSV 파일로 내보낸다. (D21~D22)

파일명: logi_calls_{YYYY-MM}_{YYYYMMDD-HHMM}.csv
저장 위치: C:/RPA/logi_exports/csv/
인코딩: UTF-8 BOM (Excel 한글 호환)
"""
import csv
from datetime import datetime
from pathlib import Path
from loguru import logger

from config import CSV_DIR, CSV_FILENAME_FMT, SHEET_HEADERS


def export_csv(month: str, rows: list[list]) -> Path:
    """
    Args:
        month: 'YYYY-MM'
        rows: sheets_uploader.read_all_rows() 반환값 (헤더 제외 리스트)

    Returns:
        저장된 CSV 파일 경로
    """
    CSV_DIR.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d-%H%M")
    filename = CSV_FILENAME_FMT.format(month=month, ts=ts)
    filepath = CSV_DIR / filename

    with filepath.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(SHEET_HEADERS)
        writer.writerows(rows)

    logger.info(f"CSV 저장 완료: {filepath} ({len(rows)}행)")
    return filepath
