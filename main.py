"""
로지 월 취합 자동화 — 진입점

사용법:
    python main.py 2026-02                        # 월 전체 취합
    python main.py 2026-02 2026-02-15             # 단일 날짜 테스트
    python main.py 2026-02-01 2026-02-05          # 날짜 범위 지정

흐름:
    1. 로지 로그인
    2. 날짜 루프 (체크포인트로 재개 가능)
       a. 기간 설정 → 조회
       b. 엑셀로 보기
       c. Excel 파싱
       d. Excel 닫기
       e. Google Sheets upsert
       f. 체크포인트 갱신
    3. CSV Export
    4. Telegram 전송
"""
import sys
from calendar import monthrange
from datetime import date, timedelta
from pathlib import Path

# 프로젝트 루트를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent))

from loguru import logger
from utils.logger import setup_logger, save_screenshot
from utils.secrets import (
    load_env,
    get_logi_credentials,
    get_spreadsheet_id,
    get_google_sa_json_path,
    get_telegram_credentials,
)
from modules import checkpoint
from modules.logi_automation import LogiAutomation
from modules.excel_parser import parse_open_excel, close_excel_without_save
from modules.sheets_uploader import upsert_rows, read_all_rows
from modules.csv_exporter import export_csv
from modules.telegram_sender import send_csv


def _generate_dates(month: str) -> list[str]:
    """'YYYY-MM' → 해당 월 모든 날짜 리스트 ['YYYY-MM-DD', ...]."""
    year, mon = int(month[:4]), int(month[5:7])
    _, last_day = monthrange(year, mon)
    return [
        date(year, mon, d).isoformat()
        for d in range(1, last_day + 1)
    ]


def _date_range(start: str, end: str) -> list[str]:
    """'YYYY-MM-DD' 시작~종료(포함) 날짜 리스트 반환."""
    d = date.fromisoformat(start)
    end_d = date.fromisoformat(end)
    result = []
    while d <= end_d:
        result.append(d.isoformat())
        d += timedelta(days=1)
    return result


def run(month: str, dates: list[str], skip_export: bool = False) -> None:
    """
    Args:
        month: 'YYYY-MM' (체크포인트/시트명/CSV명에 사용)
        dates: 처리할 날짜 리스트 ['YYYY-MM-DD', ...]
        skip_export: True면 CSV/Telegram 단계 스킵 (단일 날짜 테스트 시)
    """
    setup_logger(month)
    load_env()

    logi_id, logi_pw       = get_logi_credentials()
    spreadsheet_id          = get_spreadsheet_id()
    sa_json_path            = get_google_sa_json_path()
    bot_token, chat_id      = get_telegram_credentials()

    all_dates = dates
    state = checkpoint.load(month)
    dates_to_process = checkpoint.pending_dates(all_dates, state)

    if not dates_to_process:
        logger.info(f"[{month}] 모든 날짜 이미 완료 - CSV/Telegram 단계로 진행")
    else:
        logger.info(f"[{month}] 처리 대상: {len(dates_to_process)}일 / 전체: {len(all_dates)}일")

        logi = LogiAutomation(logi_id, logi_pw)
        logi.login()

        for date_str in dates_to_process:
            logger.info(f"━━ [{date_str}] 처리 시작 ━━")
            try:
                # 조회
                logi.query_date(date_str)

                # 엑셀로 보기
                logi.open_excel()

                # Excel 파싱
                rows = parse_open_excel(date_str)

                # Excel 닫기
                close_excel_without_save()

                if not rows:
                    logger.warning(f"[{date_str}] 파싱 결과 없음 - 완료 처리")
                    checkpoint.mark_done(state, date_str)
                    continue

                # Sheets upsert
                upsert_rows(sa_json_path, spreadsheet_id, month, rows)

                checkpoint.mark_done(state, date_str)
                logger.info(f"[{date_str}] 완료 ({len(rows)}행)")

            except Exception as e:
                logger.error(f"[{date_str}] 처리 실패: {e}")
                save_screenshot(month, f"error_{date_str}")
                checkpoint.mark_failed(state, date_str)

    # ── CSV Export ────────────────────────────────────────────────────────────
    if skip_export:
        logger.info("테스트 모드 - CSV/Telegram 스킵")
        return

    failed = state.get("failed_dates", [])
    if failed:
        logger.warning(f"실패 날짜 {len(failed)}건 존재: {failed}")

    logger.info(f"[{month}] CSV Export 시작")
    try:
        all_rows = read_all_rows(sa_json_path, spreadsheet_id, month)
        csv_path = export_csv(month, all_rows)
        state["last_csv"] = csv_path.name
        checkpoint.save(state)
    except Exception as e:
        logger.error(f"CSV Export 실패: {e}")
        return

    # ── Telegram 전송 ─────────────────────────────────────────────────────────
    logger.info(f"[{month}] Telegram 전송 시작")
    ok = send_csv(bot_token, chat_id, csv_path, month, len(all_rows))
    state["telegram_sent"] = ok
    checkpoint.save(state)

    if ok:
        logger.info(f"[{month}] 전체 파이프라인 완료")
    else:
        logger.error(f"[{month}] Telegram 전송 실패 - CSV 로컬 보관: {csv_path}")


def main() -> None:
    if len(sys.argv) < 2:
        print("사용법:")
        print("  python main.py 2026-02                  # 월 전체 취합")
        print("  python main.py 2026-02 2026-02-15       # 단일 날짜 테스트")
        print("  python main.py 2026-02-01 2026-02-05    # 날짜 범위 지정")
        sys.exit(1)

    arg1 = sys.argv[1]
    arg2 = sys.argv[2] if len(sys.argv) >= 3 else None

    # ── 모드 판별 ──────────────────────────────────────────────────────────────
    # 날짜 범위 모드: arg1이 YYYY-MM-DD 형태 (길이 10)
    if len(arg1) == 10 and arg1[4] == "-" and arg1[7] == "-":
        if not arg2 or len(arg2) != 10:
            print(f"날짜 범위 모드: 종료 날짜를 지정하세요 (예: 2026-02-05)")
            sys.exit(1)
        try:
            start_d = date.fromisoformat(arg1)
            end_d   = date.fromisoformat(arg2)
        except ValueError as e:
            print(f"날짜 형식 오류: {e}")
            sys.exit(1)
        if start_d > end_d:
            print(f"시작({arg1})이 종료({arg2})보다 늦습니다.")
            sys.exit(1)
        month = arg1[:7]   # YYYY-MM (시작 날짜 기준)
        dates = _date_range(arg1, arg2)
        skip_export = True
        logger.info(f"날짜 범위 모드: {arg1} ~ {arg2} ({len(dates)}일)")

    # 월 전체 또는 단일 날짜 테스트 모드
    elif len(arg1) == 7 and arg1[4] == "-":
        month = arg1
        if arg2:
            # 단일 날짜 테스트
            try:
                date.fromisoformat(arg2)
            except ValueError:
                print(f"날짜 형식 오류: {arg2!r} (예: 2026-02-15)")
                sys.exit(1)
            dates = [arg2]
            skip_export = True
        else:
            # 월 전체
            dates = _generate_dates(month)
            skip_export = False

    else:
        print(f"인수 형식 오류: {arg1!r}")
        print("  월: 2026-02  /  날짜: 2026-02-15")
        sys.exit(1)

    run(month, dates, skip_export)


if __name__ == "__main__":
    main()
