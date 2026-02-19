"""
로지(스마트D2) UI 자동화 모듈.

실제 UI 구조 (스크린샷 확인):
  - 메인 창 제목: "아리향콜센터-알렉스,기수신[메인/1등급]" 형태
  - 내비게이션: 상단 메뉴 "직원" → "기간별수신콜수"
  - 기간 필드: DateTimePicker 2개 (시작/종료)
  - 체크박스: "전화받은건수기준"
  - 조회 버튼: "조회(V)"
  - 엑셀 내보내기: 그리드 우클릭 → "엑셀로보기"
  - 실행 파일: C:\\SmartD2\\update.exe
"""
import subprocess
import time
from datetime import date, timedelta
from loguru import logger
from pywinauto import Application, findwindows
from pywinauto.keyboard import send_keys

from config import (
    LOGI_WINDOW_TITLE_RE,
    LOGI_MENU_EMPLOYEE,
    LOGI_SCREEN_NAME,
    LOGI_QUERY_WAIT_SEC,
    LOGI_POLL_INTERVAL_SEC,
    LOGI_POLL_MAX_SEC,
    CHECKBOX_LABEL,
    CHECKBOX_TARGET_STATE,
    PERIOD_FMT,
)

LOGI_EXEC_PATH  = r"C:\SmartD2\update.exe"
LOGIN_WAIT_SEC  = 12    # 로그인 후 메인 화면 로드 대기
MENU_WAIT_SEC   = 1.5   # 메뉴 클릭 후 화면 전환 대기


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼
# ─────────────────────────────────────────────────────────────────────────────

def _find_logi_handles() -> list:
    try:
        return findwindows.find_windows(title_re=LOGI_WINDOW_TITLE_RE)
    except Exception:
        return []


def _connect_or_start() -> Application:
    """이미 실행 중인 로지에 연결하거나, 없으면 실행 후 연결."""
    handles = _find_logi_handles()
    if handles:
        logger.debug("기존 로지 창에 연결")
        return Application(backend="uia").connect(handle=handles[0])

    logger.info(f"로지 실행: {LOGI_EXEC_PATH}")
    subprocess.Popen([LOGI_EXEC_PATH])

    deadline = time.time() + 30
    while time.time() < deadline:
        time.sleep(1)
        handles = _find_logi_handles()
        if handles:
            break
    else:
        raise TimeoutError("로지 창이 30초 내에 열리지 않았습니다.")

    return Application(backend="uia").connect(handle=handles[0])


"""
컨트롤 트리 확인 결과 (debug_controls.py):
  날짜 입력 = Pane 타입, AutomationId '1204'(시작) / '1206'(종료)
  조회 버튼 = Button, name='조 회(V)'  (공백 있음)
  체크박스  = CheckBox, name='전화받은건수기준'
  그리드    = Table, name='Report', aid='1780'
"""

# AutomationId 상수 (컨트롤 트리 덤프 확인값)
_AID_DATE_START = "1204"
_AID_DATE_END   = "1206"
_AID_TABLE      = "1780"


def _set_datetime_field(win, field_index: int, value: str) -> None:
    """
    기간 입력 Pane 컨트롤에 값 세팅. (D10)
    DevExpress DateTimePicker는 파트별 순서대로 입력해야 함:
      클릭 → 년도 입력 → 월 입력 → 일 입력 → 시간 입력 → Enter
    (각 파트 입력 후 커서가 자동으로 다음 파트로 이동함)

    value 형식: "YYYY-MM-DD HH:MM"  (예: "2026-02-01 00:00")
    """
    aid = _AID_DATE_START if field_index == 0 else _AID_DATE_END

    # "2026-02-01 00:00" → year="2026", month="02", day="01", hour="00"
    date_part, time_part = value.split(" ")
    year, month, day = date_part.split("-")
    hour = time_part[:2]

    try:
        ctrl = win.child_window(auto_id=aid)
        ctrl.click_input()
        time.sleep(0.3)

        # 년도 → 월 → 일 → 시간 순서로 각 파트를 개별 입력
        send_keys(year)
        time.sleep(0.15)
        send_keys(month)
        time.sleep(0.15)
        send_keys(day)
        time.sleep(0.15)
        send_keys(hour)
        time.sleep(0.15)
        send_keys("{ENTER}")
        time.sleep(0.2)

        logger.info(f"  기간 필드[{field_index}] 입력 완료 (aid={aid}): {value}")
        return
    except Exception as e:
        logger.warning(f"  기간 필드[{field_index}] 입력 실패: {e}")

    raise RuntimeError(
        f"기간 필드[{field_index}] 입력 실패 (aid={aid}) - "
        "debug_controls.py 재실행 후 AutomationId를 확인하세요."
    )


def _force_checkbox(win, label: str, target: bool) -> None:
    """체크박스를 목표 상태로 강제 설정. (D11)"""
    try:
        chk = win.child_window(title=label, control_type="CheckBox")
        # toggle_state: 0=off, 1=on
        current = chk.get_toggle_state()
        is_checked = (current == 1)
        if is_checked != target:
            chk.click_input()
            time.sleep(0.1)
            logger.debug(f"체크박스 '{label}': {is_checked} → {target}")
        else:
            logger.debug(f"체크박스 '{label}' 이미 목표 상태({target})")
    except Exception as e:
        logger.warning(f"체크박스 '{label}' 처리 실패(무시): {e}")


def _wait_for_query_complete(query_win) -> None:
    """
    조회 완료 대기. (D12)
    Table(name='Report', aid='1780') row count 2회 연속 동일 → 완료 판정.
    """
    time.sleep(LOGI_QUERY_WAIT_SEC)

    prev_count = -1
    stable = 0
    deadline = time.time() + LOGI_POLL_MAX_SEC

    while time.time() < deadline:
        try:
            table = query_win.child_window(auto_id=_AID_TABLE, control_type="Table")
            # Custom(Report Row) 자식 수로 행 수 추정
            row_count = len(table.children(control_type="Custom"))
            if row_count == prev_count and row_count >= 0:
                stable += 1
                if stable >= 2:
                    logger.info(f"  조회 완료 감지 ({row_count}행)")
                    return
            else:
                stable = 0
                prev_count = row_count
        except Exception:
            pass
        time.sleep(LOGI_POLL_INTERVAL_SEC)

    logger.warning("조회 완료 감지 타임아웃 - 강제 진행")


# ─────────────────────────────────────────────────────────────────────────────
# 공개 인터페이스
# ─────────────────────────────────────────────────────────────────────────────

class LogiAutomation:
    """
    GUI 모드 (권장):
        logi = LogiAutomation()
        logi.connect_to_open_screen()   # 이미 열린 기간별수신콜수 화면에 연결
        logi.query_date("2026-02-18")
        logi.open_excel()

    CLI 모드 (기존):
        logi = LogiAutomation(logi_id, logi_pw)
        logi.login()
        logi.query_date("2026-02-18")
        logi.open_excel()
    """

    def __init__(self, logi_id: str = "", logi_pw: str = "") -> None:
        self._id = logi_id
        self._pw = logi_pw
        self._app: Application | None = None
        self._main_win = None
        self._query_win = None   # "기간별수신콜수" 패널/창

    # ── 0. GUI 모드 진입점 ────────────────────────────────────────────────────

    def connect_to_open_screen(self) -> None:
        """
        사용자가 이미 기간별수신콜수 화면을 열어둔 상태에서 연결.
        로그인/내비게이션 없이 실행 중인 로지 창에 바로 붙는다.
        """
        handles = _find_logi_handles()
        if not handles:
            raise RuntimeError(
                "로지(스마트D2) 창을 찾을 수 없습니다.\n"
                "로그인 후 [기간별수신콜수] 화면으로 이동했는지 확인하세요."
            )

        self._app = Application(backend="uia").connect(handle=handles[0])
        self._main_win = self._app.window(title_re=LOGI_WINDOW_TITLE_RE)
        self._main_win.wait("visible", timeout=10)

        # 기간별수신콜수 패널 탐색
        panel = self._find_query_panel()
        if panel is None:
            raise RuntimeError(
                "기간별수신콜수 화면을 찾을 수 없습니다.\n"
                "로지에서 [직원] -> [기간별수신콜수] 화면으로 이동한 후 다시 시도하세요."
            )

        self._query_win = panel
        logger.info("기간별수신콜수 화면 연결 완료")

    # ── 1. 로그인 ─────────────────────────────────────────────────────────────

    def login(self) -> None:
        """로지 실행 → 로그인 → 메인 화면 확인."""
        self._app = _connect_or_start()
        time.sleep(1)

        main_win = self._app.window(title_re=LOGI_WINDOW_TITLE_RE)
        main_win.wait("visible", timeout=20)

        # 로그인 창이 따로 뜨는 경우 처리
        try:
            # 아이디/비밀번호 Edit 컨트롤 탐색
            edits = main_win.children(control_type="Edit")
            if len(edits) >= 2:
                # 아이디
                edits[0].set_focus()
                send_keys("^a")
                edits[0].type_keys(self._id, with_spaces=True)
                # 비밀번호
                edits[1].set_focus()
                send_keys("^a")
                edits[1].type_keys(self._pw, with_spaces=True)
                # 로그인 버튼
                login_btn = main_win.child_window(title_re=".*로그인.*", control_type="Button")
                login_btn.click_input()
                logger.info("로그인 버튼 클릭")
                time.sleep(LOGIN_WAIT_SEC)
        except Exception:
            logger.debug("로그인 창 없음 또는 이미 로그인 상태")

        # 메인 창 재연결 (로그인 후 창 제목이 바뀔 수 있음)
        time.sleep(2)
        handles = _find_logi_handles()
        if handles:
            self._app = Application(backend="uia").connect(handle=handles[0])
            self._main_win = self._app.window(title_re=LOGI_WINDOW_TITLE_RE)
            self._main_win.wait("visible", timeout=20)
            logger.info("로지 메인 화면 확인 완료")
        else:
            raise RuntimeError("로그인 후 메인 창을 찾을 수 없습니다.")

        # ── 상단 메뉴 "직원" → "기간별수신콜수" 진입 ──────────────────────
        self._navigate_to_query_screen()

    # ── 2. 화면 내비게이션 ────────────────────────────────────────────────────

    def _find_query_panel(self):
        """
        기간별수신콜수 패널/창을 탐색. 찾으면 반환, 없으면 None.
        메인 창 하위 패널, 별도 창, title_re 순으로 시도.
        """
        win = self._main_win
        if win is None:
            return None

        # 방법 1: 메인 창 하위에서 control_type별 탐색
        for ct in ("Pane", "Window", "Custom", "Document", "Group"):
            try:
                panel = win.child_window(title=LOGI_SCREEN_NAME, control_type=ct)
                panel.wait("visible", timeout=1)
                return panel
            except Exception:
                continue

        # 방법 2: 앱 레벨 별도 창
        try:
            panel = self._app.window(title=LOGI_SCREEN_NAME)
            panel.wait("visible", timeout=1)
            return panel
        except Exception:
            pass

        # 방법 3: 부분 title_re 탐색
        try:
            panel = win.child_window(title_re=f".*{LOGI_SCREEN_NAME}.*")
            panel.wait("visible", timeout=1)
            return panel
        except Exception:
            pass

        # 방법 4: 메인 창 자체에 기간별수신콜수 컨트롤이 포함된 경우
        try:
            win.child_window(title_re=r".*기간.*수신.*|.*수신콜.*")
            return win
        except Exception:
            pass

        return None

    def _navigate_to_query_screen(self) -> None:
        """
        메인 창 상단 메뉴에서 "직원" -> "기간별수신콜수" 클릭.
        이미 해당 화면에 있으면 스킵.
        CLI 모드(login() 사용 시)에서 호출됨.
        """
        win = self._main_win

        # 이미 기간별수신콜수 패널이 열려있는지 확인
        panel = self._find_query_panel()
        if panel:
            logger.debug(f"'{LOGI_SCREEN_NAME}' 화면 이미 활성")
            self._query_win = panel
            return

        # "직원" 메뉴 클릭
        try:
            menu_employee = win.child_window(title=LOGI_MENU_EMPLOYEE, control_type="MenuItem")
            menu_employee.click_input()
            time.sleep(MENU_WAIT_SEC)
            logger.debug(f"메뉴 '{LOGI_MENU_EMPLOYEE}' 클릭")
        except Exception as e:
            logger.error(f"메뉴 '{LOGI_MENU_EMPLOYEE}' 클릭 실패: {e}")
            raise

        # "기간별수신콜수" 클릭
        try:
            sub_item = win.child_window(title=LOGI_SCREEN_NAME, control_type="MenuItem")
            sub_item.click_input()
            time.sleep(MENU_WAIT_SEC)
            logger.info(f"'{LOGI_SCREEN_NAME}' 화면 진입")
        except Exception as e:
            logger.error(f"'{LOGI_SCREEN_NAME}' 메뉴 클릭 실패: {e}")
            raise

        # 패널 로드 대기 후 저장
        deadline = time.time() + 10
        while time.time() < deadline:
            panel = self._find_query_panel()
            if panel:
                self._query_win = panel
                return
            time.sleep(0.5)

        # 찾지 못하면 메인 창 폴백
        logger.warning(f"'{LOGI_SCREEN_NAME}' 패널 탐색 실패 - 메인 창으로 폴백")
        self._query_win = win

    # ── 3. 날짜 조회 ──────────────────────────────────────────────────────────

    def query_date(self, date_str: str) -> None:
        """
        특정 날짜의 데이터 조회. (D8~D12)
        기간: 해당일 00:00 ~ 다음날 00:00
        """
        target   = date.fromisoformat(date_str)
        next_day = target + timedelta(days=1)
        start_val = target.strftime(PERIOD_FMT)    # "2026-02-18 00:00"
        end_val   = next_day.strftime(PERIOD_FMT)  # "2026-02-19 00:00"

        logger.info(f"[{date_str}] 기간 설정: {start_val} ~ {end_val}")

        win = self._query_win

        # 체크박스 강제 설정 (D11)
        _force_checkbox(win, CHECKBOX_LABEL, CHECKBOX_TARGET_STATE)

        # 시작 기간 입력 (field_index=0)
        _set_datetime_field(win, 0, start_val)

        # 종료 기간 입력 (field_index=1)
        _set_datetime_field(win, 1, end_val)

        # 조회 버튼 클릭 - 실제 name='조 회(V)' (공백 포함)
        try:
            query_btn = win.child_window(
                title_re=r"조\s*회.*",
                control_type="Button",
            )
            query_btn.click_input()
            logger.info(f"  '조회(V)' 버튼 클릭")
        except Exception as e:
            logger.error(f"[{date_str}] 조회 버튼 클릭 실패: {e}")
            raise

        # 완료 대기
        _wait_for_query_complete(win)
        logger.info(f"[{date_str}] 조회 완료")

    # ── 4. 엑셀로보기 ─────────────────────────────────────────────────────────

    def open_excel(self) -> None:
        """
        그리드 우클릭 → 컨텍스트 메뉴 → "엑셀로보기" 클릭. (D13)
        """
        win = self._query_win

        # 그리드 탐색 - Table name='Report', aid='1780'
        try:
            grid = win.child_window(auto_id=_AID_TABLE, control_type="Table")
            grid.wait("visible", timeout=5)
        except Exception:
            # 폴백: Table 타입 중 첫 번째
            try:
                grid = win.child_window(control_type="Table")
                grid.wait("visible", timeout=3)
            except Exception:
                raise RuntimeError("그리드 컨트롤을 찾을 수 없습니다. (aid=1780)")

        # 그리드 우클릭
        grid.click_input(button="right")
        time.sleep(0.5)

        # 컨텍스트 메뉴에서 "엑셀로보기" 클릭
        try:
            # 메뉴 아이템이 최상위 창에 뜨는 경우
            excel_item = self._app.top_window().child_window(
                title_re=r".*엑셀로보기.*|.*엑셀로 보기.*",
                control_type="MenuItem",
            )
            excel_item.click_input()
            logger.info("'엑셀로보기' 클릭 완료")
        except Exception as e:
            logger.error(f"'엑셀로보기' 메뉴 클릭 실패: {e}")
            raise

        # Excel이 열릴 때까지 대기 (excel_parser._wait_for_excel에서 추가 대기)
        time.sleep(2)
