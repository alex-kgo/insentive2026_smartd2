# 로지 인센티브 데이터 자동화 - CLAUDE.md

## 프로젝트 목적

로지(스마트D2 v18.94) PC 프로그램에서 월별 콜 데이터를 자동 추출하여
Google Sheets에 누적 저장 → CSV 내보내기 → Telegram 전송하는 RPA 자동화 시스템.

API 미제공 → UI 자동화(pywinauto UIA 백엔드) 사용.

---

## 환경 요구사항

- Windows 10/11, Python 3.12+
- 로지(스마트D2) 실행 파일: `C:\SmartD2\update.exe`
- Microsoft Office Excel (인증 미완료 상태여도 동작, 인증 마법사 자동 처리)
- `pip install -r requirements.txt`

### .env 파일 필수 (`.env.example` 참고)

```
LOGI_ID=...
LOGI_PW=...
SPREADSHEET_ID=...
GOOGLE_SA_JSON_PATH=...
TELEGRAM_BOT_TOKEN=...
TELEGRAM_CHAT_ID=...
```

`.env`는 `.gitignore`에 포함 - 다른 PC 이동 시 직접 작성 필요.

---

## 실행 방법

### GUI 실행 (권장)

```bash
python gui.py
```

1. 취합 월 입력 (예: `2026-02`)
2. [실행] 클릭 → 수동 진행 안내 표시
3. 수동으로: 로지 로그인 → `[직원]` 메뉴 → `[기간별수신콜수]` 화면 이동
4. [자동 진행 시작] 클릭 → 날짜별 자동 반복

### CLI 실행 (보조)

```bash
python main.py 2026-02                    # 해당 월 전체
python main.py 2026-02-01 2026-02-05      # 날짜 범위
```

---

## 파일 구조

```
.
├── CLAUDE.md               # 이 파일
├── config.py               # 전역 상수/경로
├── gui.py                  # tkinter GUI 실행기 (메인 진입점)
├── main.py                 # CLI 실행기
├── requirements.txt
├── .env                    # 비밀값 (git 제외)
├── .env.example
├── debug_controls.py       # UIA 컨트롤 트리 덤프 진단 도구
├── modules/
│   ├── logi_automation.py  # UI 자동화 핵심 모듈
│   ├── excel_parser.py     # Excel COM 파싱
│   ├── sheets_uploader.py  # Google Sheets upsert
│   ├── csv_exporter.py     # CSV 내보내기
│   ├── telegram_sender.py  # Telegram 파일 전송
│   └── checkpoint.py       # JSON 진행 상태 저장
└── utils/
    ├── logger.py            # loguru 설정
    └── secrets.py           # .env 로딩
```

---

## 아키텍처 핵심 결정사항

| 항목 | 결정 | 이유 |
|------|------|------|
| UI 자동화 | pywinauto, backend='uia' | 로지가 WinForms/DevExpress 기반 |
| Excel 읽기 | win32com.client.GetActiveObject | "엑셀로보기"로 열린 인스턴스 재사용 |
| Google Sheets | gspread + 서비스 계정 | OAuth 불필요 |
| 로깅 | loguru | 파일 + GUI 큐 동시 출력 |
| 진행 저장 | checkpoint.py (JSON) | 중간 실패 후 재시작 가능 |

---

## 로지 UI 컨트롤 구조 (실제 확인값)

`debug_controls.py` 실행으로 확인한 실제 AutomationId:

```
로지 메인 창 title_re: r".*아리랑.*|.*SMART.*|.*스마트D2.*"

기간별수신콜수 패널:
  - 시작 기간 Pane:  AutomationId='1204'  (DevExpress DateTimePicker)
  - 종료 기간 Pane:  AutomationId='1206'
  - 체크박스:        CheckBox, name='전화받은건수기준'
  - 조회 버튼:       Button, name='조 회(V)'  ← 공백 있음, regex: r"조\s*회.*"
  - 그리드(Report):  Table, name='Report', AutomationId='1780'
```

---

## 날짜 입력 방법 (중요)

DevExpress DateTimePicker는 **파트별 순서 입력** 필요:
- 전체 문자열 한 번에 입력하면 오작동 ("2026-02-26 02:00" 같은 틀린 값 입력됨)

```
ctrl.click_input()
send_keys("2026")   # 년도
send_keys("02")     # 월 (자동 이동)
send_keys("01")     # 일 (자동 이동)
send_keys("00")     # 시간 (자동 이동)
send_keys("{ENTER}")
```

구현: `modules/logi_automation.py` → `_set_datetime_field()`

---

## Excel 파싱 컬럼 구조

| 컬럼 | 인덱스 | 내용 |
|------|--------|------|
| A | 0 (COL_CODE) | 코드 |
| B | 1 (COL_NAME) | 성명 |
| C | 2 (COL_C) | 고객(받음) |
| D | 3 (COL_D) | 기사(받음) |
| E | 4 (COL_E) | 고객(걸음) |
| F | 5 (COL_F) | 기사(걸음) |
| G | 6 | 합계(건) - 사용 안 함 |

- 헤더 행: 1행 (`EXCEL_HEADER_ROWS = 1`)
- 수신합계 = C + D, 발신합계 = E + F, 총합계 = 수신 + 발신

---

## GUI 흐름과 스레딩 구조

```
App (main thread, tkinter)
 └── [자동 진행 시작] 클릭
      ├── setup_logger(month)      # 먼저 호출 (logger.remove() 포함)
      ├── _add_queue_handler()     # 그 위에 GUI 큐 핸들러 추가
      └── threading.Thread(_run_automation)   # 백그라운드 스레드

_run_automation (background thread)
  ↓ COM 초기화: pythoncom.CoInitialize() in excel_parser._get_excel_com()
  ↓ logi.connect_to_open_screen()
  ↓ for each date:
      logi.query_date(date_str)
      logi.open_excel()
      parse_open_excel(date_str)      ← COM access
      close_excel_without_save()
      upsert_rows(...)
      checkpoint.mark_done(...)
  ↓ export_csv() → send_csv()
```

**핵심**: `setup_logger()` 반드시 `_add_queue_handler()` 전에 호출.
`setup_logger()`가 `logger.remove()`를 호출하므로 순서 바뀌면 GUI 로그 소실.

---

## 알려진 문제 및 해결 이력

### 1. Excel COM - 백그라운드 스레드 COM 미초기화
- **증상**: `parse_open_excel` 에서 `GetActiveObject` 실패
- **원인**: `threading.Thread`에서 `pythoncom.CoInitialize()` 미호출
- **해결**: `_get_excel_com()` 함수에 `pythoncom.CoInitialize()` 추가

### 2. Microsoft Office 인증 마법사 팝업
- **증상**: 엑셀로보기 후 인증 마법사 팝업이 파싱 전에 나타남
- **구현**: `_dismiss_office_activation_dialog()` - Excel PID 기준으로 소속 창 열거 후 닫기 버튼 클릭
- **주의**: 마법사가 없을 때 키보드 입력을 보내면 안 됨 (엉뚱한 창 오작동)
- **현황**: 사용자가 수동으로 닫아도 파싱이 작동하도록 개선됨

### 3. 날짜 입력 오류 ("2026-02-26 02:00" 등 엉뚱한 값)
- **원인**: DevExpress 피커에 전체 문자열을 한 번에 입력 (Ctrl+A 방식)
- **해결**: 년/월/일/시 파트별 send_keys 분리 입력

### 4. GUI 로그 미표시
- **원인**: `_add_queue_handler()` 후 `setup_logger()`가 `logger.remove()`로 핸들러 삭제
- **해결**: `setup_logger()` → `_add_queue_handler()` 순서로 고정

### 5. UnicodeEncodeError (em-dash `—`)
- **원인**: Windows cp949 인코딩에서 em-dash 처리 실패
- **해결**: 모든 em-dash를 하이픈 `-`으로 교체

---

## 현재 미완성/확인 필요 사항

1. **날짜 입력이 여전히 올바르게 들어가는지** - 파트별 입력 방식으로 변경했으나 실제 UI에서 검증 필요
2. **Excel 파싱 정상 동작 확인** - `pythoncom.CoInitialize()` 추가 후 재테스트 필요
3. **인증 마법사 자동 닫기 성공 여부** - 현재 사용자가 수동으로 처리 중
4. **Google Sheets upsert 완전 검증** - 아직 실제 업로드 테스트 미확인
5. **Telegram 전송 테스트** - 아직 미실행

---

## 진단 도구

### UIA 컨트롤 트리 덤프
```bash
python debug_controls.py
# → debug_controls_output.txt 생성
```
로지 창이 열린 상태에서 실행. AutomationId 확인에 사용.

### Excel 파싱 단독 테스트
```bash
# Excel에 "엑셀로보기"로 파일이 열린 상태에서:
python modules/excel_parser.py
```

---

## config.py 주요 상수

```python
BASE_DIR = Path(r"D:\_claude\인센티브_로지데이터")   # 프로젝트 루트
LOGI_WINDOW_TITLE_RE = r".*아리랑.*|.*SMART.*|.*스마트D2.*"
LOGI_SCREEN_NAME     = "기간별수신콜수"
CHECKBOX_LABEL       = "전화받은건수기준"
PERIOD_FMT = "%Y-%m-%d 00:00"           # 기간 필드 입력 포맷
EXCEL_HEADER_ROWS = 1                    # 헤더 1행 스킵
SHEET_HEADERS = ["날짜", "코드", "성명", "수신 합계", "발신 합계", "총합계"]
```

---

## Google Sheets 구조

- 탭명: `YYYY-MM` (월별 탭)
- upsert 키: `(날짜, 코드)` 조합 → 멱등성 보장
- 구현: `modules/sheets_uploader.py`

---

## requirements.txt

```
pywinauto==0.6.8
pywin32==311
gspread==6.1.2
google-auth==2.28.2
python-dotenv==1.0.1
loguru==0.7.2
requests==2.31.0
```
