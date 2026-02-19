"""
로지 창의 UIA 컨트롤 트리를 덤프해서 기간 입력 필드의 실제 타입을 확인하는 진단 스크립트.

실행 방법:
  1. 로지에 로그인하고 기간별수신콜수 화면을 열어둔다
  2. python debug_controls.py 실행
  3. debug_controls_output.txt 파일을 확인한다
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from pywinauto import Application, findwindows
from config import LOGI_WINDOW_TITLE_RE

OUTPUT_FILE = Path(__file__).parent / "debug_controls_output.txt"


def dump_tree(element, depth=0, lines=None, max_depth=8):
    if lines is None:
        lines = []
    if depth > max_depth:
        return lines

    indent = "  " * depth
    try:
        ct    = element.element_info.control_type
        name  = (element.element_info.name or "")[:60]
        aid   = (element.element_info.automation_id or "")[:40]
        rect  = element.element_info.rectangle
        lines.append(f"{indent}[{ct}] name='{name}' aid='{aid}' rect={rect}")
    except Exception as e:
        lines.append(f"{indent}[ERROR] {e}")
        return lines

    try:
        children = element.children()
        for child in children:
            dump_tree(child, depth + 1, lines, max_depth)
    except Exception:
        pass

    return lines


def main():
    print("로지 창 탐색 중...")
    handles = findwindows.find_windows(title_re=LOGI_WINDOW_TITLE_RE)
    if not handles:
        print(f"로지 창을 찾을 수 없습니다. (title_re={LOGI_WINDOW_TITLE_RE})")
        print("로지에 로그인하고 기간별수신콜수 화면을 열어두세요.")
        return

    app = Application(backend="uia").connect(handle=handles[0])
    win = app.window(title_re=LOGI_WINDOW_TITLE_RE)
    win.wait("visible", timeout=5)

    print(f"연결된 창: {win.element_info.name!r}")
    print(f"컨트롤 트리 덤프 중... (최대 8단계)")

    lines = dump_tree(win.wrapper_object())

    OUTPUT_FILE.write_text(
        "\n".join(lines),
        encoding="utf-8",
    )

    print(f"\n저장 완료: {OUTPUT_FILE}")
    print(f"총 컨트롤 수: {len(lines)}")
    print("\n상위 40줄 미리보기:")
    print("\n".join(lines[:40]))


if __name__ == "__main__":
    main()
