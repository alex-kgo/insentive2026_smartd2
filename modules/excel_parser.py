"""
"엑셀로 보기"로 열린 Excel 인스턴스에서 데이터를 파싱한다. (D15~D17)

win32com.client.GetActiveObject("Excel.Application")로 현재 열린 Excel에 접근.
컬럼 구조:
  A(0): 코드   B(1): 성명
  C(2): 고객(받음)   D(3): 기사(받음)
  E(4): 고객(걸음)   F(5): 기사(걸음)
  G(6): 합계(참고용, 사용 안 함)

반환값: list[dict] - 파싱된 행 목록
  {
    "날짜": "2026-02-18",
    "코드": "A001",
    "성명": "홍길동",
    "수신합계": 12,
    "발신합계": 8,
    "총합계": 20,
  }
"""
import time
from typing import Any
from loguru import logger

from config import (
    COL_CODE, COL_NAME, COL_C, COL_D, COL_E, COL_F,
    EXCEL_HEADER_ROWS,
)


def _safe_int(value: Any, cell_ref: str = "") -> int:
    """빈칸/None/변환 실패 → 0 (WARN 로그)."""
    if value is None or str(value).strip() == "":
        return 0
    try:
        return int(float(str(value).replace(",", "")))
    except (ValueError, TypeError):
        logger.warning(f"숫자 변환 실패 → 0 처리: 셀={cell_ref!r}, 값={value!r}")
        return 0


def _get_excel_com():
    """
    현재 열린 Excel COM 인스턴스를 반환.
    백그라운드 스레드 안전을 위해 pythoncom.CoInitialize() 포함.
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass

    import win32com.client
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception as e:
        raise RuntimeError(f"Excel COM 연결 실패: {e}")


def _wait_for_excel(timeout_sec: float = 30.0) -> Any:
    """
    Excel이 완전히 로드될 때까지 대기.
    ActiveWorkbook이 생길 때까지 폴링.
    """
    deadline = time.time() + timeout_sec
    while time.time() < deadline:
        try:
            xl = _get_excel_com()
            wb = xl.ActiveWorkbook
            if wb is not None:
                logger.debug("Excel ActiveWorkbook 확인됨")
                return xl
        except Exception:
            pass
        time.sleep(0.5)
    raise TimeoutError(f"Excel이 {timeout_sec}초 내에 열리지 않았습니다.")


def _click_close_in_window(win) -> bool:
    """
    win 하위에서 '닫기'가 포함된 Button을 찾아 클릭.
    성공하면 True, 못 찾으면 False.
    """
    try:
        for btn in win.descendants(control_type="Button"):
            try:
                name = btn.window_text() or ""
                if "닫기" in name or "Close" in name:
                    btn.click_input()
                    time.sleep(0.4)
                    logger.info(f"인증 마법사 [닫기] 클릭 완료: '{name}'")
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def _dismiss_office_activation_dialog(timeout_sec: float = 10.0) -> None:
    """
    Excel 창을 전면으로 이동한 뒤 인증 마법사 다이얼로그를 닫는다.

    Excel 프로세스 PID로 소속 창 전체를 열거하고
    메인 Excel 창이 아닌 창에서 [닫기] 버튼을 찾아 클릭한다.
    버튼을 찾지 못해도 키보드 입력은 보내지 않는다 (엉뚱한 창 오작동 방지).
    """
    try:
        import win32process
        from pywinauto import findwindows, Application

        # 1. Excel 창 (메인) 탐색 및 전면 이동
        xl_handles = findwindows.find_windows(title_re=r".*Excel.*")
        if not xl_handles:
            logger.debug("Excel 창 없음 - 인증 마법사 스킵")
            return

        xl_handle = xl_handles[0]
        xl_app = Application(backend="uia").connect(handle=xl_handle)
        xl_win = xl_app.window(handle=xl_handle)
        xl_win.set_focus()
        time.sleep(0.5)
        logger.debug("Excel 창 전면 이동 완료")

        # 2. Excel 프로세스 PID
        _, excel_pid = win32process.GetWindowThreadProcessId(xl_handle)

        # 3. 해당 PID 소속 창 전체 열거 → 메인 창 제외 = 다이얼로그 후보
        deadline = time.time() + timeout_sec
        while time.time() < deadline:
            try:
                all_wins = findwindows.find_windows(process=excel_pid)
            except Exception:
                all_wins = []

            other_wins = [h for h in all_wins if h != xl_handle]

            if not other_wins:
                logger.debug("인증 마법사 없음 - 정상 진행")
                return

            for hwnd in other_wins:
                try:
                    dlg_app = Application(backend="uia").connect(handle=hwnd)
                    dlg = dlg_app.window(handle=hwnd)
                    dlg.set_focus()
                    time.sleep(0.2)
                    title = dlg.window_text() or "(제목없음)"
                    logger.debug(f"  다이얼로그 발견: '{title}'")

                    if _click_close_in_window(dlg):
                        logger.info("인증 마법사 닫기 완료")
                        time.sleep(0.5)
                        return

                    # 닫기 버튼 미발견 시 키보드 입력 없이 그냥 넘어감
                    logger.debug(f"  닫기 버튼 미발견(스킵): '{title}'")
                except Exception:
                    continue

            time.sleep(0.5)

        logger.debug("인증 마법사 대기 타임아웃 - 정상 진행")

    except Exception as e:
        logger.debug(f"인증 마법사 처리 예외(무시): {e}")


def parse_open_excel(date_str: str, timeout_sec: float = 30.0) -> list[dict]:
    """
    현재 열려있는 Excel ActiveSheet에서 데이터를 파싱한다.

    Args:
        date_str: 루프 날짜 (YYYY-MM-DD). 엑셀 날짜값 무시하고 이 값 사용.
        timeout_sec: Excel 인스턴스 대기 최대 시간(초).

    Returns:
        파싱된 행 목록. 빈 시트면 [].
    """
    logger.info(f"[{date_str}] Excel 파싱 시작")

    # 1. Excel 로드 대기
    _wait_for_excel(timeout_sec)

    # 2. 인증 마법사가 있으면 닫기
    _dismiss_office_activation_dialog()

    # 3. 인증 마법사 처리 후 fresh COM reference 재취득
    try:
        xl = _get_excel_com()
    except Exception as e:
        raise RuntimeError(f"Excel COM 재연결 실패: {e}")

    # 4. ActiveWorkbook / ActiveSheet 접근
    try:
        wb = xl.ActiveWorkbook
        if wb is None:
            raise RuntimeError("ActiveWorkbook이 None - 열린 통합문서가 없습니다.")
        ws = wb.ActiveSheet
        logger.debug(f"[{date_str}] 시트 접근 성공: '{ws.Name}'")
    except Exception as e:
        raise RuntimeError(f"ActiveSheet 접근 실패: {e}")

    # 5. 사용된 마지막 행 파악
    try:
        last_row = ws.UsedRange.Rows.Count
    except Exception as e:
        raise RuntimeError(f"UsedRange 접근 실패: {e}")

    logger.debug(f"[{date_str}] UsedRange 행 수: {last_row}")

    if last_row <= EXCEL_HEADER_ROWS:
        logger.warning(f"[{date_str}] 데이터 없음 (헤더만 존재, 총 {last_row}행)")
        return []

    # 6. 행 파싱
    rows: list[dict] = []
    skipped = 0

    for row_idx in range(EXCEL_HEADER_ROWS + 1, last_row + 1):  # 1-based
        try:
            # ri=row_idx: 루프 변수 클로저 캡처 방지
            def cell(col_0based: int, ri: int = row_idx) -> Any:
                return ws.Cells(ri, col_0based + 1).Value  # COM은 1-based

            code = str(cell(COL_CODE) or "").strip()
            name = str(cell(COL_NAME) or "").strip()

            # 코드/성명이 모두 비어있으면 합계행 등 → 스킵
            if not code and not name:
                skipped += 1
                continue

            col_c = _safe_int(cell(COL_C), f"C{row_idx}")
            col_d = _safe_int(cell(COL_D), f"D{row_idx}")
            col_e = _safe_int(cell(COL_E), f"E{row_idx}")
            col_f = _safe_int(cell(COL_F), f"F{row_idx}")

            수신합계 = col_c + col_d
            발신합계 = col_e + col_f
            총합계   = 수신합계 + 발신합계

            rows.append({
                "날짜":   date_str,
                "코드":   code,
                "성명":   name,
                "수신합계": 수신합계,
                "발신합계": 발신합계,
                "총합계":  총합계,
            })

        except Exception as e:
            logger.warning(f"[{date_str}] 행 {row_idx} 파싱 실패(스킵): {e}")
            skipped += 1
            continue

    logger.info(f"[{date_str}] 파싱 완료 - {len(rows)}행 (스킵 {skipped}행)")
    return rows


def close_excel_without_save() -> None:
    """열린 Excel을 저장 없이 닫는다."""
    try:
        xl = _get_excel_com()
        wb = xl.ActiveWorkbook
        if wb:
            wb.Close(SaveChanges=False)
            logger.debug("Excel 닫기 완료 (저장 안 함)")
    except Exception as e:
        logger.warning(f"Excel 닫기 실패(무시): {e}")


# ── 단독 실행 테스트 ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent.parent))
    from utils.logger import setup_logger
    setup_logger("TEST")

    test_date = "2026-02-18"
    try:
        result = parse_open_excel(test_date, timeout_sec=10)
        for r in result:
            print(r)
    except Exception as ex:
        print(f"오류: {ex}")
