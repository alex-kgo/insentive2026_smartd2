"""
로지 월 취합 자동화 - GUI 실행기

흐름:
  1. 취합 월 입력 후 [실행] 클릭
  2. 수동 진행 안내 표시 (로지 로그인 / 기간별수신콜수 이동)
  3. 완료 후 [자동 진행 시작] 클릭
  4. 자동 진행: 기간 설정 -> 조회 -> 엑셀로보기 -> 파싱 -> Sheets upsert 반복
  5. 완료 후 CSV Export -> Telegram 전송
"""
import queue
import sys
import threading
import tkinter as tk
from calendar import monthrange
from datetime import date, timedelta
from pathlib import Path
from tkinter import font, messagebox, scrolledtext, ttk

# 프로젝트 루트 경로 추가
sys.path.insert(0, str(Path(__file__).parent))

# ─────────────────────────────────────────────────────────────────────────────
# GUI 로그 라우터: 백그라운드 스레드 -> 큐 -> GUI 텍스트 위젯
# ─────────────────────────────────────────────────────────────────────────────

_log_queue: queue.Queue = queue.Queue()


def _queue_log(level: str, message: str) -> None:
    _log_queue.put((level, message))


def _add_queue_handler() -> None:
    """
    기존 핸들러를 유지한 채 GUI 큐 핸들러만 추가한다.
    setup_logger() 호출 이후에 실행해야 한다.
    """
    from loguru import logger
    logger.add(
        lambda msg: _queue_log(msg.record["level"].name, msg.record["message"]),
        format="{message}",
        level="INFO",
        colorize=False,
    )


# ─────────────────────────────────────────────────────────────────────────────
# 자동화 파이프라인 (백그라운드 스레드에서 실행)
# ─────────────────────────────────────────────────────────────────────────────

def _run_automation(month: str, done_callback, error_callback) -> None:
    """
    백그라운드 스레드 함수.
    기간별수신콜수 화면이 열려 있다는 전제로 자동 진행.
    """
    try:
        from loguru import logger
        from utils.secrets import load_env, get_spreadsheet_id, get_google_sa_json_path, get_telegram_credentials
        from modules import checkpoint
        from modules.logi_automation import LogiAutomation
        from modules.excel_parser import parse_open_excel, close_excel_without_save
        from modules.sheets_uploader import upsert_rows, read_all_rows
        from modules.csv_exporter import export_csv
        from modules.telegram_sender import send_csv

        # ── 환경 설정 로드 ────────────────────────────────────────────────────
        logger.info("환경 변수 로드 중...")
        load_env()
        spreadsheet_id     = get_spreadsheet_id()
        sa_json_path       = get_google_sa_json_path()
        bot_token, chat_id = get_telegram_credentials()
        logger.info("환경 변수 로드 완료")

        # ── 날짜 목록 ─────────────────────────────────────────────────────────
        year, mon = int(month[:4]), int(month[5:7])
        _, last_day = monthrange(year, mon)
        all_dates = [date(year, mon, d).isoformat() for d in range(1, last_day + 1)]

        state = checkpoint.load(month)
        dates_to_process = checkpoint.pending_dates(all_dates, state)

        if not dates_to_process:
            logger.info(f"[{month}] 모든 날짜 이미 완료 - CSV/Telegram 단계로 진행")
        else:
            logger.info(
                f"[{month}] 처리 대상: {len(dates_to_process)}일 / "
                f"전체: {len(all_dates)}일 "
                f"(이미 완료: {len(all_dates) - len(dates_to_process)}일)"
            )

            # ── 로지 화면 연결 ────────────────────────────────────────────────
            logger.info("로지 창 연결 중...")
            logi = LogiAutomation()
            logi.connect_to_open_screen()
            logger.info("기간별수신콜수 화면 연결 완료")

            total = len(dates_to_process)
            for idx, date_str in enumerate(dates_to_process, 1):
                logger.info(f"[{idx}/{total}] {date_str} 처리 시작")
                try:
                    # 1. 기간 설정 + 조회
                    logger.info(f"  [1/5] 기간 설정 및 조회 중...")
                    logi.query_date(date_str)
                    logger.info(f"  [1/5] 조회 완료")

                    # 2. 엑셀로보기
                    logger.info(f"  [2/5] 엑셀로보기 실행 중...")
                    logi.open_excel()
                    logger.info(f"  [2/5] Excel 열림")

                    # 3. Excel 파싱
                    logger.info(f"  [3/5] Excel 데이터 파싱 중...")
                    rows = parse_open_excel(date_str)
                    close_excel_without_save()

                    if not rows:
                        logger.warning(f"  [3/5] 데이터 없음 - 이 날짜는 스킵")
                        checkpoint.mark_done(state, date_str)
                        continue
                    logger.info(f"  [3/5] 파싱 완료 ({len(rows)}행)")

                    # 4. Google Sheets 업로드
                    logger.info(f"  [4/5] Google Sheets 업로드 중...")
                    upsert_rows(sa_json_path, spreadsheet_id, month, rows)
                    logger.info(f"  [4/5] 업로드 완료")

                    # 5. 체크포인트
                    checkpoint.mark_done(state, date_str)
                    logger.info(f"  [5/5] {date_str} 완료 ({len(rows)}행)")

                except Exception as e:
                    logger.error(f"  [오류] {date_str} 실패: {e}")
                    checkpoint.mark_failed(state, date_str)

            failed = state.get("failed_dates", [])
            logger.info(
                f"[{month}] 날짜 루프 완료 - "
                f"성공: {len(state.get('done_dates', []))}일, "
                f"실패: {len(failed)}일"
            )
            if failed:
                logger.warning(f"  실패 날짜: {', '.join(failed)}")

        # ── CSV Export ────────────────────────────────────────────────────────
        logger.info(f"[{month}] CSV Export 시작...")
        all_rows = read_all_rows(sa_json_path, spreadsheet_id, month)
        csv_path = export_csv(month, all_rows)
        state["last_csv"] = csv_path.name
        checkpoint.save(state)
        logger.info(f"[{month}] CSV 저장 완료: {csv_path.name} ({len(all_rows)}행)")

        # ── Telegram 전송 ─────────────────────────────────────────────────────
        logger.info(f"[{month}] Telegram 전송 중...")
        ok = send_csv(bot_token, chat_id, csv_path, month, len(all_rows))
        state["telegram_sent"] = ok
        checkpoint.save(state)

        if ok:
            logger.info(f"[{month}] Telegram 전송 완료")
        else:
            logger.error(f"[{month}] Telegram 전송 실패 - CSV 로컬 보관: {csv_path}")

        done_callback(month, len(all_rows))

    except Exception as e:
        import traceback
        error_callback(f"{e}\n{traceback.format_exc()}")


# ─────────────────────────────────────────────────────────────────────────────
# GUI 클래스
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("로지 월 취합 자동화")
        self.resizable(False, False)
        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self):
        PAD = 12

        # ── 헤더 ─────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg="#2c3e50", pady=10)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="로지 월 취합 자동화",
            bg="#2c3e50", fg="white",
            font=("맑은 고딕", 14, "bold"),
        ).pack()

        # ── 월 입력 ───────────────────────────────────────────────────────────
        row1 = tk.Frame(self, pady=PAD, padx=PAD)
        row1.pack(fill="x")
        tk.Label(row1, text="취합 월 (YYYY-MM):", font=("맑은 고딕", 10)).pack(side="left")
        self._month_var = tk.StringVar(value="2026-02")
        self._month_entry = tk.Entry(row1, textvariable=self._month_var, width=12, font=("맑은 고딕", 10))
        self._month_entry.pack(side="left", padx=6)
        self._run_btn = tk.Button(
            row1, text="실행", width=8,
            font=("맑은 고딕", 10, "bold"),
            bg="#27ae60", fg="white",
            activebackground="#219150",
            relief="flat", cursor="hand2",
            command=self._on_run_click,
        )
        self._run_btn.pack(side="left")

        ttk.Separator(self, orient="horizontal").pack(fill="x")

        # ── 수동 진행 안내 ─────────────────────────────────────────────────────
        manual_frame = tk.LabelFrame(
            self, text="  수동 진행  ",
            font=("맑은 고딕", 10, "bold"),
            padx=PAD, pady=PAD,
        )
        manual_frame.pack(fill="x", padx=PAD, pady=(PAD, 0))

        self._manual_label = tk.Label(
            manual_frame,
            text="[실행] 버튼을 눌러 취합 월을 확인하세요.",
            font=("맑은 고딕", 10),
            fg="#555",
            justify="left",
        )
        self._manual_label.pack(anchor="w")

        self._auto_btn = tk.Button(
            manual_frame,
            text="자동 진행 시작",
            font=("맑은 고딕", 11, "bold"),
            bg="#2980b9", fg="white",
            activebackground="#206090",
            relief="flat", cursor="hand2",
            pady=6,
            state="disabled",
            command=self._on_auto_click,
        )
        self._auto_btn.pack(fill="x", pady=(10, 0))

        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=(PAD, 0))

        # ── 진행 로그 ──────────────────────────────────────────────────────────
        log_frame = tk.Frame(self, padx=PAD)
        log_frame.pack(fill="both", expand=True, pady=(4, PAD))
        tk.Label(log_frame, text="진행 로그:", font=("맑은 고딕", 9, "bold")).pack(anchor="w")
        self._log_text = scrolledtext.ScrolledText(
            log_frame,
            width=70, height=18,
            font=("Consolas", 9),
            state="disabled",
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white",
        )
        self._log_text.pack(fill="both", expand=True)
        # 색상 태그
        self._log_text.tag_config("INFO",    foreground="#7ec8e3")
        self._log_text.tag_config("WARNING", foreground="#f0c040")
        self._log_text.tag_config("ERROR",   foreground="#f87171")
        self._log_text.tag_config("DEBUG",   foreground="#888888")
        self._log_text.tag_config("SUCCESS", foreground="#6ee7b7")

    # ── 이벤트 ────────────────────────────────────────────────────────────────

    def _on_run_click(self):
        month = self._month_var.get().strip()
        if not self._validate_month(month):
            return

        self._run_btn.config(state="disabled")
        self._month_entry.config(state="disabled")

        self._manual_label.config(
            text=(
                f"취합 월: {month}\n\n"
                "아래 작업을 직접 수행한 후\n"
                "[자동 진행 시작] 버튼을 클릭하세요.\n\n"
                "  1.  로지(스마트D2) 프로그램에 로그인\n"
                "  2.  상단 메뉴 [직원] 클릭\n"
                "  3.  [기간별수신콜수] 화면으로 이동"
            ),
            fg="#2c3e50",
        )
        self._auto_btn.config(state="normal")
        self._log("INFO", f"월 설정 완료: {month} - 수동 진행을 완료한 후 [자동 진행 시작]을 누르세요.")

    def _on_auto_click(self):
        month = self._month_var.get().strip()
        self._auto_btn.config(state="disabled")
        self._log("INFO", "자동 진행 시작...")

        # 1. 파일 로거 먼저 설정 (logger.remove() 포함)
        from utils.logger import setup_logger
        setup_logger(month)
        # 2. 그 위에 GUI 큐 핸들러 추가 (remove 없음)
        _add_queue_handler()

        # 백그라운드 스레드 실행
        t = threading.Thread(
            target=_run_automation,
            args=(month, self._on_done, self._on_error),
            daemon=True,
        )
        t.start()

    def _on_done(self, month: str, total_rows: int):
        self.after(0, lambda: self._log(
            "SUCCESS",
            f"[완료] {month} 취합 성공 - 총 {total_rows:,}행 / Telegram 전송 완료"
        ))
        self.after(0, lambda: self._run_btn.config(state="normal"))
        self.after(0, lambda: self._month_entry.config(state="normal"))

    def _on_error(self, msg: str):
        self.after(0, lambda: self._log("ERROR", f"[오류] {msg}"))
        self.after(0, lambda: self._run_btn.config(state="normal"))
        self.after(0, lambda: self._month_entry.config(state="normal"))
        self.after(0, lambda: self._auto_btn.config(state="normal"))

    # ── 헬퍼 ──────────────────────────────────────────────────────────────────

    def _validate_month(self, month: str) -> bool:
        if len(month) != 7 or month[4] != "-":
            messagebox.showerror("입력 오류", f"월 형식이 올바르지 않습니다.\n예: 2026-02\n\n입력값: {month}")
            return False
        try:
            int(month[:4]); int(month[5:])
        except ValueError:
            messagebox.showerror("입력 오류", f"월 형식이 올바르지 않습니다: {month}")
            return False
        return True

    def _log(self, level: str, message: str):
        self._log_text.config(state="normal")
        import datetime
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        tag = level if level in ("INFO", "WARNING", "ERROR", "DEBUG", "SUCCESS") else "INFO"
        self._log_text.insert("end", f"[{ts}] {message}\n", tag)
        self._log_text.see("end")
        self._log_text.config(state="disabled")

    def _poll_log_queue(self):
        """큐에서 로그를 꺼내 텍스트 위젯에 표시."""
        try:
            while True:
                level, message = _log_queue.get_nowait()
                self._log(level, message)
        except queue.Empty:
            pass
        self.after(100, self._poll_log_queue)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
