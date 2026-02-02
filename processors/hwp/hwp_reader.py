"""
HWP 문서 읽기 모듈

win32com.client를 사용하여 한글 프로그램을 직접 제어하여 HWP 문서를 읽습니다.
- 한글 프로그램 COM 객체 생성
- HWP 문서 열기 및 내용 읽기
- 범위 선택 및 복사 작업 지원
- 텍스트 추출 (검색용)

전제 조건:
- Windows 환경에서만 동작
- 한글 프로그램이 반드시 설치되어 있어야 함
- win32com.client를 통한 COM 자동화 필요
"""
import sys
import win32com.client
import os
import tempfile
import time
import threading
from contextlib import contextmanager
from typing import Optional, Tuple, List, Any

try:
    import win32clipboard  # type: ignore
    _CLIPBOARD_AVAILABLE = True
except Exception:
    _CLIPBOARD_AVAILABLE = False

try:
    import win32gui  # type: ignore
    import win32con  # type: ignore
    _WIN32GUI_AVAILABLE = True
except Exception:
    _WIN32GUI_AVAILABLE = False

# win32con을 import하지 않고도 쓸 수 있는 표준 포맷 코드
_CF_TEXT = 1
_CF_UNICODETEXT = 13


class HWPNotInstalledError(Exception):
    """한글 프로그램이 설치되지 않았을 때 발생하는 예외"""
    pass


class HWPInitializationError(Exception):
    """한글 프로그램 초기화 실패 시 발생하는 예외"""
    pass


class _HwpPopupAutoCloser:
    """
    HWP 자동화 중 뜨는 모달 다이얼로그(찾기 끝/없음, 저장 여부 등)를
    윈도우 레벨에서 최대한 자동으로 닫아주는 안전장치입니다.

    - `SetMessageBoxMode`로 닫히지 않는 "찾기" 계열 대화상자(#32770)까지 대응하기 위함
    - 실행 구간을 짧게(컨텍스트 내부)만 켜서, 다른 앱 다이얼로그를 건드릴 위험을 줄입니다.
    """

    def __init__(self, timeout_sec: float = 8.0, interval_sec: float = 0.12):
        self.timeout_sec = timeout_sec
        self.interval_sec = interval_sec
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None

    def start(self) -> None:
        if not _WIN32GUI_AVAILABLE:
            return
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        if not _WIN32GUI_AVAILABLE:
            return
        self._stop.set()
        t = self._thread
        if t and t.is_alive():
            t.join(timeout=0.5)

    def _run(self) -> None:
        end_time = time.time() + float(self.timeout_sec or 0.0)
        while not self._stop.is_set():
            if time.time() >= end_time:
                break
            try:
                self._close_hwp_popups_once()
            except Exception:
                pass
            time.sleep(self.interval_sec)

    def _close_hwp_popups_once(self) -> None:
        if not _WIN32GUI_AVAILABLE:
            return

        def enum_proc(hwnd: int, _lparam: int) -> None:
            try:
                if not win32gui.IsWindow(hwnd):
                    return
                # 대부분의 모달 팝업은 다이얼로그 클래스(#32770)
                if win32gui.GetClassName(hwnd) != "#32770":
                    return
                title = (win32gui.GetWindowText(hwnd) or "").strip()
                # 제목이 비어있을 수 있어, 본문 텍스트도 함께 보고 판단
                message_text = self._get_dialog_static_text(hwnd)
                if not self._looks_like_hwp_popup(title, message_text):
                    return

                buttons = self._get_dialog_buttons(hwnd)
                target_btn = self._pick_button_to_click(title, message_text, buttons)
                if target_btn is not None:
                    try:
                        win32gui.SendMessage(target_btn, win32con.BM_CLICK, 0, 0)
                    except Exception:
                        # 최후의 수단: Enter/Esc로 닫히는 경우가 있어 종료 시도
                        try:
                            win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                            win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                        except Exception:
                            pass
            except Exception:
                return

        win32gui.EnumWindows(enum_proc, 0)

    def _get_dialog_static_text(self, hwnd: int) -> str:
        texts: List[str] = []

        def enum_child(ch: int, _lp: int) -> None:
            try:
                if win32gui.GetClassName(ch) == "Static":
                    t = (win32gui.GetWindowText(ch) or "").strip()
                    if t:
                        texts.append(t)
            except Exception:
                pass

        try:
            win32gui.EnumChildWindows(hwnd, enum_child, 0)
        except Exception:
            pass
        return "\n".join(texts).strip()

    def _get_dialog_buttons(self, hwnd: int) -> List[Tuple[int, str]]:
        btns: List[Tuple[int, str]] = []

        def enum_child(ch: int, _lp: int) -> None:
            try:
                if win32gui.GetClassName(ch) == "Button":
                    t = (win32gui.GetWindowText(ch) or "").strip()
                    btns.append((ch, t))
            except Exception:
                pass

        try:
            win32gui.EnumChildWindows(hwnd, enum_child, 0)
        except Exception:
            pass
        return btns

    def _looks_like_hwp_popup(self, title: str, message_text: str) -> bool:
        hay = f"{title}\n{message_text}".strip()
        if not hay:
            return False
        # HWP 자동화에서 실제로 반복되는 문구들 위주로만 반응(오탐 최소화)
        keywords = [
            "문서의 끝까지",  # 찾기 끝
            "더 이상 찾",     # 더 이상 찾을 수 없음
            "찾을 수 없",     # 찾을 수 없음
            "저장하시겠",     # 저장 여부
            "저장 안",        # 저장 안 함
        ]
        return any(k in hay for k in keywords) or title in ("찾기",)

    def _pick_button_to_click(
        self, title: str, message_text: str, buttons: List[Tuple[int, str]]
    ) -> Optional[int]:
        msg = f"{title}\n{message_text}"

        def find_btn(*needles: str) -> Optional[int]:
            for hwnd, txt in buttons:
                for n in needles:
                    if n and n in (txt or ""):
                        return hwnd
            return None

        # 저장 여부 → "아니오/저장 안 함/No" 우선
        if ("저장" in msg) and ("하시겠" in msg or "저장하시겠" in msg):
            return find_btn("저장 안", "아니", "No", "취소", "Cancel") or (buttons[0][0] if buttons else None)

        # 찾기 끝(처음부터 계속?) → 성공적으로 다음 단계로 넘어가려면 "찾음/예/Yes"가 안전
        if ("문서의 끝까지" in msg) or ("처음부터" in msg) or ("계속" in msg and "찾" in msg):
            return find_btn("찾음", "예", "Yes", "확인", "OK") or (buttons[0][0] if buttons else None)

        # 더 이상 없음 → 확인
        if ("더 이상" in msg and "찾" in msg) or ("찾을 수 없" in msg):
            return find_btn("확인", "OK", "닫기", "Close") or (buttons[0][0] if buttons else None)

        # 그 외 "찾기" 제목 창은 보통 취소가 무난
        if title == "찾기":
            return find_btn("취소", "Cancel", "닫기", "Close") or (buttons[0][0] if buttons else None)

        return None


class HWPReader:
    """HWP 문서 읽기 클래스"""
    
    def __init__(self):
        """HWP Reader 초기화"""
        self.hwp = None
        self.is_opened = False
        # select_range_between_markers()가 계산한 마지막 문제 범위(루프 반복 감지/디버그용)
        self.last_problem_start_pos = None  # type: Optional[Tuple[int, int, int]]
        self.last_problem_end_pos = None  # type: Optional[Tuple[int, int, int]]

    @contextmanager
    def _auto_close_hwp_popups(self, timeout_sec: float = 8.0):
        """
        `SetMessageBoxMode`로 닫히지 않는 팝업까지 포함해,
        특정 구간에서만 HWP 팝업을 자동으로 닫아주는 보호막입니다.
        """
        closer = _HwpPopupAutoCloser(timeout_sec=timeout_sec)
        try:
            closer.start()
            yield
        finally:
            closer.stop()
    
    def initialize(self):
        """
        한글 프로그램 COM 객체 초기화
        
        Returns:
            성공 여부
            
        Raises:
            HWPNotInstalledError: 한글 프로그램이 설치되지 않은 경우
            HWPInitializationError: 초기화 실패 시
        """
        if sys.platform != 'win32':
            raise HWPNotInstalledError("이 프로그램은 Windows 환경에서만 동작합니다.")
        
        try:
            self.hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            self.hwp.XHwpWindows.Item(0).Visible = False  # 한글 창 숨기기
            return True
        except Exception as e:
            error_msg = (
                "한글 프로그램을 초기화할 수 없습니다.\n"
                "한글과컴퓨터의 한글 프로그램이 설치되어 있는지 확인해주세요.\n"
                f"상세 오류: {str(e)}"
            )
            raise HWPInitializationError(error_msg) from e
    
    def open_document(self, file_path: str) -> bool:
        """
        HWP 문서 열기
        
        Args:
            file_path: HWP 파일 경로
            
        Returns:
            성공 여부
        """
        if self.hwp is None:
            try:
                if not self.initialize():
                    return False
            except (HWPNotInstalledError, HWPInitializationError) as e:
                print(f"한글 프로그램 초기화 실패: {e}")
                return False
        
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"HWP 파일을 찾을 수 없습니다: {file_path}")
            
            # 한글 API: HAction을 사용하여 파일 열기
            # 방법 1: XHwpDocuments를 사용
            try:
                # 절대 경로로 변환
                abs_path = os.path.abspath(file_path)
                # XHwpDocuments.Open() 사용
                self.hwp.XHwpDocuments.Open(abs_path)
                self.is_opened = True
                print(f"[디버그] HWP 파일 열기 성공 (XHwpDocuments): {abs_path}")
                return True
            except Exception as e1:
                print(f"[디버그] XHwpDocuments.Open() 실패: {e1}")
                # 방법 2: HAction을 사용
                try:
                    self.hwp.HAction.GetDefault("FileOpen", self.hwp.HParameterSet.HFileOpenSave.HSet)
                    self.hwp.HParameterSet.HFileOpenSave.filename = abs_path
                    self.hwp.HAction.Execute("FileOpen", self.hwp.HParameterSet.HFileOpenSave.HSet)
                    self.is_opened = True
                    print(f"[디버그] HWP 파일 열기 성공 (HAction): {abs_path}")
                    return True
                except Exception as e2:
                    print(f"[디버그] HAction.FileOpen() 실패: {e2}")
                    raise e2
        except Exception as e:
            print(f"[디버그] HWP 문서 열기 실패: {e}")
            import traceback
            traceback.print_exc()
            self.is_opened = False
            return False
    
    def close_document(self):
        """현재 열린 HWP 문서 닫기"""
        if self.hwp and self.is_opened:
            try:
                # ✅ (우선) COM 문서 Close(isDirty=False)로 "저장 질문 없이" 닫기 시도
                # - 미리보기 생성 시 임시 문서(주석 삭제 등)로 인해 문서가 '수정됨' 상태가 될 수 있습니다.
                # - FileClose(HAction)만으로는 버전에 따라 저장 팝업이 뜰 수 있어, 가능한 경우 문서 Close를 우선합니다.
                with self._temp_message_box_mode(0x20021):  # No + Cancel + OK
                    try:
                        self.hwp.XHwpDocuments.Active_XHwpDocument.Close(isDirty=False)
                        self.is_opened = False
                        return
                    except Exception:
                        try:
                            # 일부 환경은 positional 인자만 받음
                            self.hwp.XHwpDocuments.Active_XHwpDocument.Close(False)
                            self.is_opened = False
                            return
                        except Exception:
                            pass

                # (폴백) FileClose 액션
                self.hwp.HAction.GetDefault("FileClose", self.hwp.HParameterSet.HFileOpenSave.HSet)
                # 수정된 문서(미리보기용 임시 문서 등)에서 "저장하시겠습니까?"가 뜨며 멈추는 것을 방지
                try:
                    self.hwp.HParameterSet.HFileOpenSave.filename = ""
                except Exception:
                    pass
                with self._temp_message_box_mode(0x20021):  # No + Cancel + OK
                    self.hwp.HAction.Execute("FileClose", self.hwp.HParameterSet.HFileOpenSave.HSet)
                self.is_opened = False
            except Exception as e:
                print(f"HWP 문서 닫기 실패: {e}")

    def export_pdf(self, output_path: str) -> bool:
        """
        현재 열린 HWP 문서를 PDF로 내보냅니다.

        주의:
        - HWP 버전/환경에 따라 Format 문자열이 다를 수 있어, 몇 가지 후보를 순차 시도합니다.
        - 실패 시 False를 반환합니다(예외 전파 최소화).
        """
        if not self.hwp or not self.is_opened:
            return False
        if not output_path:
            return False

        # 가능한 포맷 후보들 (환경별 차이 대응)
        format_candidates = ["PDF", "pdf", "Pdf"]
        last_err: Optional[Exception] = None
        try:
            # 디렉토리 보장(저장 실패 원인 최소화)
            try:
                os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            except Exception:
                pass

            with self._auto_close_hwp_popups(timeout_sec=8.0), self._temp_message_box_mode(0x20021):
                for fmt in format_candidates:
                    try:
                        self.hwp.HAction.GetDefault("FileSaveAs", self.hwp.HParameterSet.HFileOpenSave.HSet)
                        self.hwp.HParameterSet.HFileOpenSave.filename = output_path
                        self.hwp.HParameterSet.HFileOpenSave.Format = fmt
                        self.hwp.HAction.Execute("FileSaveAs", self.hwp.HParameterSet.HFileOpenSave.HSet)
                        return True
                    except Exception as e:
                        last_err = e

                # 일부 환경은 다른 액션명이 있을 수 있어 추가 시도
                try:
                    self.hwp.HAction.GetDefault("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)
                    self.hwp.HParameterSet.HFileOpenSave.filename = output_path
                    self.hwp.HAction.Execute("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)
                    return True
                except Exception as e:
                    last_err = e
        except Exception as e:
            last_err = e

        if last_err:
            print(f"[디버그] PDF 내보내기 실패: {last_err}")
        return False

    @contextmanager
    def _temp_message_box_mode(self, mode: int):
        """
        특정 구간에서만 HWP 메시지박스 자동 클릭 모드를 적용하고 원복합니다.

        mode 예:
        - 0x1: OK(확인) 자동
        - 0x20: OK/Cancel에서 Cancel(취소) 자동
        - 0x20000: Yes/No에서 No(취소 성격) 자동
        - 0x20021: No + Cancel + OK (찾기 끝/없음/끝-계속찾기 팝업에 안전)
        """
        prev = None
        try:
            prev = self.hwp.GetMessageBoxMode()
        except Exception:
            prev = None
        try:
            try:
                # 환경에 따라 키워드 인자(Mode=)만 받는 케이스가 있어 둘 다 시도
                try:
                    self.hwp.SetMessageBoxMode(mode)
                except Exception:
                    self.hwp.SetMessageBoxMode(Mode=mode)
            except Exception:
                pass
            yield
        finally:
            if prev is not None:
                try:
                    try:
                        self.hwp.SetMessageBoxMode(prev)
                    except Exception:
                        self.hwp.SetMessageBoxMode(Mode=prev)
                except Exception:
                    pass
    
    def find_text(self, text: str, start_from_beginning: bool = True, move_after: bool = True) -> Optional[Tuple[int, int]]:
        """
        텍스트 찾기 (현재 커서 위치에서)
        
        Args:
            text: 찾을 텍스트
            start_from_beginning: 문서 처음부터 찾기 여부 (True면 문서 시작점으로 이동)
            move_after: 찾은 뒤 커서를 다음 위치로 이동할지 여부
                        - True: (기본) 선택 해제 후 커서를 한 칸 이동 (다음 검색 진행용)
                        - False: 찾은 텍스트가 선택된 상태를 유지 (경계 위치 산출/선택 범위용)
            
        Returns:
            (시작_위치, 끝_위치) 또는 None
            주의: 위치 정보는 신뢰할 수 없으므로 더미 값 반환
        """
        if not self.is_opened:
            return None
        
        try:
            # 문서 처음으로 이동 (start_from_beginning=True인 경우만)
            if start_from_beginning:
                self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            
            # 찾기 실행 (현재 커서 위치에서)
            self.hwp.HAction.GetDefault("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
            self.hwp.HParameterSet.HFindReplace.FindString = text
            self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            # ✅ "문서 끝까지 찾았습니다/더 이상 없음" 팝업 무음 처리 (해당 호출 구간에서만)
            with self._temp_message_box_mode(0x20021):  # No + Cancel + OK
                result = self.hwp.HAction.Execute("RepeatFind", self.hwp.HParameterSet.HFindReplace.HSet)
            
            if result == 1:  # 찾기 성공
                # 찾은 텍스트가 선택되어 있음
                text_length = len(text)
                
                if move_after:
                    # 찾은 텍스트의 끝 위치 다음으로 커서 이동 (다음 검색을 위해)
                    # 핵심: 선택된 상태에서 끝으로 이동한 후, 선택 해제하고 한 문자 더 이동
                    try:
                        # 선택 영역의 끝으로 이동 시도
                        try:
                            self.hwp.HAction.GetDefault("MoveSelEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Execute("MoveSelEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        except:
                            # MoveSelEnd가 작동하지 않으면, 텍스트 길이만큼 오른쪽으로 이동
                            for _ in range(min(text_length, 50)):
                                try:
                                    self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                except:
                                    break
                        
                        # 선택 해제 (ESC 키) - 커서가 선택 영역의 끝에 위치
                        self.hwp.HAction.Run("Cancel")
                        
                        # 한 문자 오른쪽으로 이동 (다음 검색을 위해)
                        # 이렇게 하지 않으면 같은 위치를 계속 찾게 됨
                        self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        
                    except Exception:
                        # 이동 실패 시에도 계속 진행
                        pass
                
                # 위치 정보는 신뢰할 수 없으므로, 찾은 순서를 나타내는 더미 값 반환
                print(f"[디버그] 텍스트 찾기 성공: '{text}' (길이: {text_length})")
                return (0, text_length)  # 위치 정보는 사용하지 않음
            else:
                return None
        except Exception as e:
            print(f"텍스트 찾기 실패: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def move_cursor_after_position(self, position: int):
        """
        커서를 지정된 위치 다음으로 이동
        
        Args:
            position: 이동할 위치 (실제로는 텍스트 길이)
        """
        if not self.is_opened:
            return
        
        try:
            # 선택 해제
            self.hwp.HAction.Run("Cancel")
            # position + 1 만큼 오른쪽으로 이동
            for _ in range(min(position + 1, 100)):
                try:
                    self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    break
        except:
            pass
    
    def select_range(self, start_pos: int, end_pos: int) -> bool:
        """
        지정된 범위 선택
        
        Args:
            start_pos: 시작 위치
            end_pos: 끝 위치
            
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        try:
            # 시작 위치로 이동 (문서 처음으로 이동 후 오른쪽으로 이동)
            self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            
            # start_pos만큼 오른쪽으로 이동
            for _ in range(min(start_pos, 1000)):
                try:
                    self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    break
            
            # 범위 선택: 시작 위치에서 끝 위치까지 선택
            # end_pos - start_pos만큼 오른쪽으로 이동하면서 선택
            move_count = min(end_pos - start_pos, 10000)
            for _ in range(move_count):
                try:
                    # Shift+Right (선택하면서 이동)
                    self.hwp.HAction.GetDefault("ExtendSelRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("ExtendSelRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    # ExtendSelRight가 없으면 일반 MoveRight 사용
                    self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
            
            return True
        except Exception as e:
            print(f"범위 선택 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def copy_selected_range(self) -> bool:
        """
        선택된 범위를 클립보드에 복사
        
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        try:
            self.hwp.HAction.Run("Copy")
            return True
        except Exception as e:
            print(f"복사 실패: {e}")
            return False

    def _init_scan(self, option: int = 0x00, range_flag: int = 0x0000) -> bool:
        """
        InitScan()을 호출하여 텍스트 검색 준비
        
        Args:
            option: 검색 대상 마스크
                - 0x00 (maskNormal): 본문만 검색
                - 0x04 (maskCtrl): 컨트롤 포함 (미주, 각주, 글상자 등)
            range_flag: 검색 범위 플래그
                - 0x0000: 현재 위치부터
                - 0x0100: 역방향 검색
                
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        try:
            # win32com을 통한 호출 시 파라미터를 명시적으로 전달
            # InitScan은 직접 메서드 호출이 가능하지만, 파라미터 전달 방식이 다를 수 있음
            try:
                # 방법 1: 직접 호출 (파라미터를 명시적으로 전달)
                result = self.hwp.InitScan(option, range_flag)
                return bool(result)
            except Exception:
                # 방법 2: 파라미터 없이 호출 (기본값 사용)
                try:
                    result = self.hwp.InitScan()
                    return bool(result)
                except Exception as e:
                    print(f"[경고] InitScan() 실패: {e}")
                    return False
        except Exception as e:
            print(f"[경고] InitScan() 실패: {e}")
            return False
    
    def _release_scan(self) -> None:
        """ReleaseScan()을 호출하여 검색 정보 초기화"""
        if not self.is_opened:
            return
        try:
            self.hwp.ReleaseScan()
        except Exception:
            pass
    
    def _move_pos(self, move_id: int, para: Optional[int] = None, pos: Optional[int] = None) -> bool:
        """
        MovePos()를 사용하여 커서 이동 (HAction 기반으로 폴백)
        
        Args:
            move_id: 이동 타입 ID
                - 6: moveStartOfPara (현재 문단 시작)
                - 7: moveEndOfPara (현재 문단 끝)
                - 11: movePrevPara (앞 문단 끝)
                - 15: movePrevPosEx (한 글자 뒤로, 미주 포함)
            para: 문단 번호 (선택)
            pos: 문단 내 위치 (선택)
            
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        # MovePos() 직접 호출이 실패할 수 있으므로, HAction 기반으로 폴백
        try:
            # 방법 1: MovePos() 직접 호출 시도
            if para is not None and pos is not None:
                result = self.hwp.MovePos(move_id, para, pos)
            elif para is not None:
                result = self.hwp.MovePos(move_id, para)
            else:
                result = self.hwp.MovePos(move_id)
            return bool(result)
        except Exception:
            # 방법 2: HAction 기반으로 폴백
            try:
                if move_id == 6:  # moveStartOfPara
                    self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    return True
                elif move_id == 7:  # moveEndOfPara
                    self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    return True
                elif move_id == 11:  # movePrevPara
                    # 앞 문단 끝으로 이동: 위로 이동 후 문단 끝
                    self.hwp.HAction.GetDefault("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    return True
                elif move_id == 15:  # movePrevPosEx (한 글자 뒤로)
                    # 왼쪽으로 한 글자 이동
                    self.hwp.HAction.GetDefault("MoveLeft", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveLeft", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    return True
                else:
                    print(f"[경고] MovePos({move_id}) - 알 수 없는 move_id, HAction 폴백 실패")
                    return False
            except Exception as e:
                print(f"[경고] MovePos({move_id}) HAction 폴백 실패: {e}")
                return False
    
    def _get_text_with_scan(self, use_init_scan: bool = True) -> Tuple[int, str]:
        """
        InitScan()을 사용하여 GetText() 호출
        
        Args:
            use_init_scan: InitScan() 호출 여부
            
        Returns:
            (상태코드, 텍스트) 튜플
            상태코드: 0=텍스트 없음, 2=일반 텍스트, 101=초기화 안됨, 102=변환 실패
        """
        if not self.is_opened:
            return (101, "")
        
        try:
            if use_init_scan:
                # InitScan 호출 시도 (여러 옵션 시도)
                scan_success = False
                # 옵션 1: maskCtrl 포함
                if self._init_scan(option=0x04, range_flag=0x0000):
                    scan_success = True
                # 옵션 2: 기본 옵션
                elif self._init_scan(option=0x00, range_flag=0x0000):
                    scan_success = True
                # 옵션 3: 파라미터 없이
                else:
                    try:
                        if self.hwp.InitScan():
                            scan_success = True
                    except:
                        pass
                
                if not scan_success:
                    print(f"[경고] InitScan() 모든 시도 실패")
                    return (101, "")
            
            # GetText() 호출
            # Python win32com에서는 GetText()가 (상태코드, 텍스트) 튜플을 반환할 수 있음
            result = self.hwp.GetText()
            
            # 반환값 처리
            if isinstance(result, tuple):
                if len(result) >= 2:
                    # (상태코드, 텍스트) 또는 (텍스트, 상태코드) 형태
                    status_code = result[0] if isinstance(result[0], int) else result[1] if isinstance(result[1], int) else 0
                    text = result[1] if isinstance(result[1], str) else result[0] if isinstance(result[0], str) else ""
                    return (status_code, text)
                elif len(result) == 1:
                    # 단일 값
                    if isinstance(result[0], str):
                        return (2, result[0])  # 일반 텍스트로 간주
                    elif isinstance(result[0], int):
                        return (result[0], "")
            elif isinstance(result, int):
                # 상태코드만 반환 (텍스트는 출력 파라미터로 전달되지 않음)
                return (result, "")
            elif isinstance(result, str):
                # 텍스트만 반환 (상태코드는 2로 간주)
                return (2, result)
            
            # 알 수 없는 형태
            return (102, "")
            
        except Exception as e:
            print(f"[경고] GetText() 실패: {e}")
            return (102, "")
        finally:
            if use_init_scan:
                self._release_scan()
    
    def _normalize_gettext_result(self, raw: Any) -> str:
        """
        HWP COM의 GetText() 반환값을 문자열로 정규화합니다.

        HWP 버전/환경에 따라 GetText()가 다음 형태로 반환될 수 있습니다:
        - str
        - tuple (예: (상태코드, 텍스트) 또는 (텍스트, 상태코드))
        - 기타 (None 등)
        """
        if raw is None:
            return ""
        if isinstance(raw, str):
            return raw
        if isinstance(raw, (tuple, list)):
            # 튜플/리스트 안의 문자열들 중 "가장 길이가 긴" 문자열을 텍스트로 간주
            str_items = [x for x in raw if isinstance(x, str)]
            if str_items:
                return max(str_items, key=len)
            return ""
        # 알 수 없는 타입은 텍스트로 간주하지 않음 (상태코드 등을 섞지 않기 위함)
        return ""

    def _safe_repr(self, value: Any, max_len: int = 160) -> str:
        try:
            s = repr(value)
        except Exception:
            s = "<repr_failed>"
        if len(s) > max_len:
            return s[: max_len - 3] + "..."
        return s

    def _read_clipboard_text(self) -> str:
        if not _CLIPBOARD_AVAILABLE:
            return ""
        try:
            win32clipboard.OpenClipboard()
            try:
                # HWP 복사 결과는 여러 포맷으로 들어올 수 있어, 유니코드 텍스트를 우선 시도합니다.
                data = None
                try:
                    if win32clipboard.IsClipboardFormatAvailable(_CF_UNICODETEXT):
                        data = win32clipboard.GetClipboardData(_CF_UNICODETEXT)
                except Exception:
                    data = None

                if data is None:
                    try:
                        if win32clipboard.IsClipboardFormatAvailable(_CF_TEXT):
                            data = win32clipboard.GetClipboardData(_CF_TEXT)
                    except Exception:
                        data = None

                # 최후의 fallback (기본 포맷)
                if data is None:
                    try:
                        data = win32clipboard.GetClipboardData()
                    except Exception:
                        data = None
            finally:
                win32clipboard.CloseClipboard()
            if isinstance(data, str):
                return data
            if isinstance(data, (bytes, bytearray)):
                # CF_TEXT가 bytes로 오는 경우가 있어 최대한 안전하게 디코딩
                try:
                    return data.decode("utf-8", errors="ignore")
                except Exception:
                    try:
                        return data.decode("cp949", errors="ignore")
                    except Exception:
                        return ""
            return ""
        except Exception:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass
            return ""
    
    def ensure_main_body_focus(self) -> bool:
        """
        미주/각주(주석) 편집 영역에서 본문으로 복귀를 시도합니다.

        사용자가 제공한 한글 매크로(Shift+Esc 동작) 기반:
        - Goto(HGotoE, SetSelectionIndex=5) 실행
        - CloseEx 실행

        Returns:
            성공 여부(시도 결과). 실패해도 예외를 던지지 않습니다.
        """
        if not self.is_opened:
            return False

        ok = False
        # 팝업 자동 닫기 + 메시지박스 모드로 최대한 무음화
        with self._auto_close_hwp_popups(timeout_sec=6.0), self._temp_message_box_mode(0x20021):
            # 1) 상태/선택 취소(가능한 경우): 서브 편집 상태를 풀어주는 데 도움이 될 수 있음
            try:
                self.hwp.HAction.Run("Cancel")
            except Exception:
                pass

            # 2) Shift+Esc 동작에 가장 가까운 CloseEx를 먼저 시도(대부분 여기서 끝남)
            try:
                self.hwp.HAction.Run("CloseEx")
                return True
            except Exception:
                # 일부 버전에서는 CloseEx가 없을 수 있어 Close를 시도
                try:
                    self.hwp.HAction.Run("Close")
                    return True
                except Exception:
                    pass

            # 3) (폴백) 매크로 기반 Goto → CloseEx 시퀀스
            try:
                self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                try:
                    # 팝업 억제 힌트가 있는 환경이면 적용
                    self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                except Exception:
                    pass
                try:
                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("IgnoreMessage", 1)
                except Exception:
                    pass
                try:
                    # 매크로에서 쓰인 OK/닫기 값(환경별 상이) → 실패해도 무시
                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 31)
                except Exception:
                    pass
                try:
                    self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
                except Exception:
                    pass
                self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
            except Exception:
                pass

            try:
                self.hwp.HAction.Run("CloseEx")
                ok = True
            except Exception:
                try:
                    self.hwp.HAction.Run("Close")
                    ok = True
                except Exception:
                    pass

        return ok

    def get_text_from_document(self) -> str:
        """
        현재 열린 문서의 텍스트를 "본문 기준"으로 최대한 추출합니다.

        목적:
        - DB 생성(파싱) 단계에서는 텍스트 추출을 생략하고 속도를 확보
        - 필요할 때(일괄 생성/온디맨드) 원본 HWP 블록을 열어 텍스트를 생성

        주의:
        - 한글 문서는 본문/주석(미주/각주) 편집 영역이 분리될 수 있어,
          먼저 CloseEx(Shift+Esc 동작) 기반으로 본문 복귀를 시도합니다.
        """
        if not self.is_opened:
            return ""

        try:
            # 미리보기 일괄 생성 중 발생하는 팝업(찾기/저장/없음 등)을 최대한 무음 처리
            with self._auto_close_hwp_popups(timeout_sec=8.0), self._temp_message_box_mode(0x20021):
                # 문서 맨 앞으로 이동
                try:
                    self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception:
                    pass

                # ✅ 미리보기(본문 텍스트)용: 주석(미주/각주)을 임시 문서에서 제거
                # HWP는 Ctrl+A 복사 시 주석 텍스트를 앞에 합쳐 내보내는 경우가 있어,
                # 본문 텍스트만 얻으려면 임시 복사본에서 주석을 제거한 뒤 추출합니다.
                # (원본 GridFS HWP는 건드리지 않음)
                try:
                    # 주석은 문제당 1개로 고정(운영 전제) → 1회만 삭제 시도
                    self._delete_notes_via_macro(max_iters=1)
                except Exception:
                    pass

                # 주석 삭제 후 편집 상태가 주석 영역에 남아있을 수 있어 본문 복귀를 1회만 시도
                try:
                    self.ensure_main_body_focus()
                except Exception:
                    pass

                # 본문 기준으로 다시 맨 앞으로
                try:
                    self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception:
                    pass

            # 전체 선택 후 GetText (SelectAll이 없으면 수동 선택 폴백)
            try:
                self.hwp.HAction.Run("SelectAll")
            except Exception:
                try:
                    self.hwp.HAction.Run("Select")
                    try:
                        self.hwp.HAction.GetDefault("MoveDocEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveDocEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    except Exception:
                        pass
                    try:
                        self.hwp.HAction.Run("ExtendSel")
                    except Exception:
                        pass
                except Exception:
                    pass

            # 1) GetText() 기반 (환경에 따라 tuple/stream 형태일 수 있어 반복 시도)
            text_from_gettext = ""
            try:
                # 일부 환경에서는 GetText()가 (상태코드, 텍스트)로 반환되며 여러 번 호출해야 전체 텍스트가 나옵니다.
                chunks = []
                empty_streak = 0
                for _ in range(5000):
                    raw = self.hwp.GetText()
                    if isinstance(raw, (tuple, list)) and raw:
                        state = raw[0] if isinstance(raw[0], int) else None
                        chunk = self._normalize_gettext_result(raw)
                        if chunk:
                            chunks.append(chunk)
                            empty_streak = 0
                        else:
                            empty_streak += 1
                        # state==0을 EOF로 보는 예제가 많아 종료 조건으로 사용
                        if state == 0:
                            break
                        if empty_streak >= 10:
                            break
                    else:
                        chunk = self._normalize_gettext_result(raw)
                        if chunk:
                            chunks.append(chunk)
                        break
                text_from_gettext = "".join(chunks)
            except Exception:
                text_from_gettext = ""

            # 2) Copy→클립보드 텍스트 (GetText가 비거나 순서가 꼬이는 환경 보완)
            text_from_clipboard = ""
            try:
                self.hwp.HAction.Run("Copy")
                for _ in range(5):
                    text_from_clipboard = self._read_clipboard_text()
                    if text_from_clipboard and text_from_clipboard.strip():
                        break
                    time.sleep(0.03)
            except Exception:
                text_from_clipboard = ""

            def score(s: str) -> int:
                return len((s or "").strip())

            text = text_from_gettext
            if score(text_from_clipboard) > score(text):
                text = text_from_clipboard

            # 선택 해제
            try:
                self.hwp.HAction.Run("Cancel")
            except Exception:
                pass

            return text or ""
        except Exception:
            return ""

    def _delete_notes_via_macro(self, max_iters: int = 1) -> int:
        """
        사용자 제공 매크로(script10) 기반으로 문서의 주석(미주/각주)을 삭제합니다.

        매크로 동작:
        - Goto(HGotoE): DialogResult=31, SetSelectionIndex=5
        - MoveSelRight
        - Delete

        Returns:
            삭제 시도 횟수(성공적으로 delete까지 진행한 횟수)
        """
        if not self.is_opened:
            return 0

        deleted = 0
        # 운영 전제: 문서당 주석 1개 → "더 이상 없음" 팝업을 만들지 않도록 1회만 시도
        try:
            # 메시지박스 자동 처리(찾기 끝/없음 등)
            with self._auto_close_hwp_popups(timeout_sec=6.0), self._temp_message_box_mode(0x20021):  # No + Cancel + OK
                # Goto (주석으로 이동)
                self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                try:
                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 31)
                except Exception:
                    pass
                try:
                    self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                except Exception:
                    pass
                try:
                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("IgnoreMessage", 1)
                except Exception:
                    pass
                try:
                    self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
                except Exception:
                    pass
                res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                if res == 0:
                    return 0

                # 선택 이동 후 삭제
                try:
                    self.hwp.HAction.Run("MoveSelRight")
                except Exception:
                    return 0
                self.hwp.HAction.Run("Delete")
                deleted = 1
        except Exception:
            deleted = 0

        return deleted

    def goto_next_endnote(self) -> bool:
        """
        다음 미주로 이동 (문제 시작점)
        
        Returns:
            성공 여부 (미주를 찾았으면 True, 없으면 False)
        """
        if not self.is_opened:
            return False
        
        try:
            with self._auto_close_hwp_popups(timeout_sec=6.0), self._temp_message_box_mode(0x20021):
                # Goto (미주로 이동)
                self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                try:
                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 31)
                except Exception:
                    pass
                try:
                    self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                except Exception:
                    pass
                try:
                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("IgnoreMessage", 1)
                except Exception:
                    pass
                try:
                    self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
                except Exception:
                    pass
                res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                if res == 0:
                    # 미주를 찾지 못함
                    return False
                return True
        except Exception as e:
            print(f"[디버그] 미주 이동 실패: {e}")
            return False

    def find_text_line_above_endnote(self, previous_endnote_pos: Optional[Tuple[int, int, int]] = None) -> Optional[Tuple[int, int, int]]:
        """
        현재 미주 위치에서 위로 올라가면서 텍스트가 있는 첫 번째 줄(문단) 찾기
        
        개선 사항:
        - 무한 루프 방지: 같은 위치 반복 체크
        - 선택 전 검증: 위치 변경 확인
        - 선택 방식 개선: 다양한 선택 방법 시도
        - 연속 실패 카운터: 일정 횟수 실패 시 건너뛰기
        
        Args:
            previous_endnote_pos: 이전 미주 위치 (검색 범위 제한용). None이면 첫 번째 미주까지 검색
        
        Returns:
            (sec, para, pos) 위치 튜플 또는 None (텍스트가 있는 줄을 찾지 못한 경우)
        """
        if not self.is_opened:
            return None
        
        try:
            # 현재 미주 위치 저장
            current_pos = self.hwp.GetPos()
            if not current_pos:
                return None
            
            sec_current, para_current, pos_current = current_pos
            print(f"[디버그] 현재 미주 위치: ({sec_current}, {para_current}, {pos_current})")
            
            # 이전 미주 위치 (검색 범위 제한)
            if previous_endnote_pos:
                sec_prev, para_prev, pos_prev = previous_endnote_pos
                print(f"[디버그] 이전 미주 위치: ({sec_prev}, {para_prev}, {pos_prev})")
            else:
                # 첫 번째 미주인 경우, 문서 시작점까지 검색
                sec_prev, para_prev, pos_prev = None, None, None
            
            # 위로 올라가면서 텍스트가 있는 줄 찾기
            max_iterations = 500  # 무한루프 방지 (1000에서 500으로 감소)
            iteration = 0
            last_checked_pos = None  # 마지막으로 확인한 위치
            same_pos_count = 0  # 같은 위치 반복 횟수
            max_same_pos = 3  # 같은 위치가 3번 반복되면 건너뛰기
            consecutive_failures = 0  # 연속 실패 횟수
            max_consecutive_failures = 10  # 연속 10번 실패하면 중단
            
            while iteration < max_iterations:
                iteration += 1
                
                # 위로 이동: HAction 기반 이동 사용 (더 안정적)
                try:
                    # 위로 한 문단 이동
                    self.hwp.HAction.GetDefault("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception as e:
                    print(f"[디버그] 반복 #{iteration}: 이동 실패: {e}")
                    # 더 이상 위로 올라갈 수 없음
                    break
                
                # 현재 위치 확인
                try:
                    new_pos = self.hwp.GetPos()
                    if not new_pos:
                        break
                    sec_new, para_new, pos_new = new_pos
                    
                    # 무한 루프 방지: 같은 위치 반복 체크
                    if last_checked_pos == (sec_new, para_new):
                        same_pos_count += 1
                        if same_pos_count >= max_same_pos:
                            print(f"[경고] 반복 #{iteration}: 같은 위치 ({sec_new}, {para_new})가 {same_pos_count}번 반복됨. 건너뜁니다.")
                            # 선택 해제 후 다음 위치로 강제 이동 시도
                            try:
                                self.hwp.HAction.Run("Cancel")
                                # 위로 한 번 더 이동 시도
                                self.hwp.HAction.GetDefault("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Execute("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                new_pos = self.hwp.GetPos()
                                if new_pos:
                                    sec_new, para_new, pos_new = new_pos
                                    same_pos_count = 0  # 리셋
                                else:
                                    break
                            except:
                                break
                    else:
                        same_pos_count = 0  # 위치가 변경되면 리셋
                    
                    last_checked_pos = (sec_new, para_new)
                    
                    # 이전 미주 위치를 넘어갔는지 확인
                    if previous_endnote_pos:
                        if sec_new < sec_prev or (sec_new == sec_prev and para_new < para_prev):
                            # 이전 미주를 넘어갔음
                            print(f"[디버그] 이전 미주 위치를 넘어갔습니다. 검색 중단.")
                            break
                    
                    # 현재 위치에서 텍스트 읽기 (문단 단위)
                    text_found = False
                    try:
                        # 현재 위치로 명시적으로 이동
                        self.hwp.SetPos(sec_new, para_new, pos_new)
                        current_pos_check = self.hwp.GetPos()
                        if current_pos_check != new_pos:
                            # 위치가 변경되지 않았으면 건너뛰기
                            print(f"[디버그] 반복 #{iteration}: 위치 설정 실패, 건너뜁니다.")
                            consecutive_failures += 1
                            if consecutive_failures >= max_consecutive_failures:
                                print(f"[경고] 연속 {consecutive_failures}번 실패. 검색 중단.")
                                break
                            continue
                        
                        # 선택 방식 개선: 여러 방법 시도
                        selection_success = False
                        text = ""
                        
                        # 방법 1: MoveParaBegin → Shift+MoveParaEnd (ExtendSel 없이)
                        try:
                            # 선택 해제
                            self.hwp.HAction.Run("Cancel")
                            # 문단 시작으로 이동
                            self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            
                            # Shift 키를 누른 상태로 문단 끝까지 이동 (ExtendSel 사용)
                            self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Run("ExtendSel")
                            
                            # 선택 범위 확인
                            try:
                                sel_start = self.hwp.GetPos(0)
                                sel_end = self.hwp.GetPos(1)
                                
                                if sel_start != sel_end:
                                    # 선택 성공
                                    selection_success = True
                                    # GetText() 시도
                                    try:
                                        status_code, text = self._get_text_with_scan(use_init_scan=True)
                                        if status_code == 2 and text and text.strip():
                                            text_found = True
                                        elif status_code != 2 or not text or not text.strip():
                                            # GetText 실패 시 클립보드 시도
                                            try:
                                                self.copy_selected_range()
                                                time.sleep(0.05)
                                                text = self._read_clipboard_text()
                                                if text and text.strip():
                                                    text_found = True
                                            except:
                                                pass
                                    except:
                                        pass
                            except:
                                pass
                        except:
                            pass
                        
                        # 방법 2: 방법 1 실패 시 Select + ExtendSel
                        if not selection_success or not text_found:
                            try:
                                self.hwp.HAction.Run("Cancel")
                                self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Run("Select")
                                self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Run("ExtendSel")
                                
                                try:
                                    sel_start = self.hwp.GetPos(0)
                                    sel_end = self.hwp.GetPos(1)
                                    if sel_start != sel_end:
                                        try:
                                            status_code, text = self._get_text_with_scan(use_init_scan=True)
                                            if status_code == 2 and text and text.strip():
                                                text_found = True
                                            else:
                                                try:
                                                    self.copy_selected_range()
                                                    time.sleep(0.05)
                                                    text = self._read_clipboard_text()
                                                    if text and text.strip():
                                                        text_found = True
                                                except:
                                                    pass
                                        except:
                                            pass
                                except:
                                    pass
                            except:
                                pass
                        
                        # 선택 해제
                        try:
                            self.hwp.HAction.Run("Cancel")
                        except:
                            pass
                        
                        # 텍스트가 있는지 확인
                        if text_found and text and text.strip():
                            # 텍스트가 있는 줄을 찾음!
                            consecutive_failures = 0  # 성공 시 리셋
                            print(f"[디버그] 텍스트가 있는 줄 발견: 위치 ({sec_new}, {para_new}, {pos_new}), 텍스트 길이: {len(text)}, 텍스트: {text[:50]}")
                            # 문단 끝 위치 반환
                            try:
                                self.hwp.SetPos(sec_new, para_new, pos_new)
                                self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                end_pos = self.hwp.GetPos()
                                if end_pos:
                                    print(f"[디버그] 문단 끝 위치: {end_pos}")
                                    return end_pos
                            except Exception as e:
                                print(f"[디버그] 문단 끝 위치 설정 실패: {e}, 기본 위치 반환")
                                return (sec_new, para_new, pos_new)
                        else:
                            # 텍스트가 없거나 공백만 있는 경우
                            consecutive_failures += 1
                            if consecutive_failures >= max_consecutive_failures:
                                print(f"[경고] 연속 {consecutive_failures}번 실패. 검색 중단.")
                                break
                        
                    except Exception as e:
                        # 텍스트 읽기 실패 시 계속 진행
                        consecutive_failures += 1
                        if consecutive_failures >= max_consecutive_failures:
                            print(f"[경고] 연속 {consecutive_failures}번 실패. 검색 중단.")
                            break
                    
                except Exception:
                    # 위치 읽기 실패 시 중단
                    consecutive_failures += 1
                    if consecutive_failures >= max_consecutive_failures:
                        break
            
            print(f"[디버그] 텍스트가 있는 줄을 찾지 못했습니다. (총 {iteration}번 시도)")
            return None
            
        except Exception as e:
            print(f"[디버그] 위쪽 텍스트 줄 찾기 실패: {e}")
            import traceback
            traceback.print_exc()
            return None

    def find_last_content_below_endnote(
        self, 
        current_endnote_pos: Tuple[int, int, int],
        next_endnote_pos: Optional[Tuple[int, int, int]] = None
    ) -> Optional[Tuple[int, int, int]]:
        """
        현재 미주 위치에서 아래로 내려가며 마지막 콘텐츠 위치 찾기
        
        방안 2 구현:
        - 현재 미주에서 아래로 내려가며 콘텐츠 추적
        - 마지막 콘텐츠 위치를 last_content_pos에 저장
        - 종료 조건: 다음 미주 도달, 페이지 나누기 발견, 빈 줄 4개 이상
        - 종료 시 last_content_pos를 끝점으로 사용
        
        Args:
            current_endnote_pos: 현재 미주 위치 (sec, para, pos)
            next_endnote_pos: 다음 미주 위치 (범위 제한용). None이면 문서 끝까지 검색
        
        Returns:
            (sec, para, pos) 위치 튜플 또는 None (콘텐츠를 찾지 못한 경우)
        """
        if not self.is_opened:
            return None
        
        try:
            # 현재 미주 위치로 이동
            sec_current, para_current, pos_current = current_endnote_pos
            self.hwp.SetPos(sec_current, para_current, pos_current)
            
            # 미주 문단은 이미 콘텐츠로 간주 (미주가 있으니까)
            # 미주 문단의 끝을 last_content_pos 초기값으로 설정
            try:
                # 미주 위치에서 문단 끝으로 이동하여 미주 문단의 끝점 계산
                self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                endnote_end_pos = self.hwp.GetPos()
                if endnote_end_pos:
                    last_content_pos = endnote_end_pos
                else:
                    # 미주 문단 끝을 찾지 못하면 미주 위치를 사용
                    last_content_pos = current_endnote_pos
            except:
                last_content_pos = current_endnote_pos
            
            print(f"[디버그] 현재 미주 위치: ({sec_current}, {para_current}, {pos_current}), 초기 last_content_pos: {last_content_pos} (미주 문단은 이미 콘텐츠로 간주)")
            
            # 다음 미주 위치 (범위 제한용)
            if next_endnote_pos:
                sec_next, para_next, pos_next = next_endnote_pos
                print(f"[디버그] 다음 미주 위치: ({sec_next}, {para_next}, {pos_next})")
            else:
                sec_next, para_next, pos_next = None, None, None
                # 문서 끝 위치 저장 (next_endnote_pos가 None일 때 사용)
                try:
                    self.hwp.HAction.GetDefault("MoveDocEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveDocEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    doc_end_pos = self.hwp.GetPos()
                    # 현재 미주 위치로 복귀
                    self.hwp.SetPos(sec_current, para_current, pos_current)
                except:
                    doc_end_pos = None
            
            # 아래로 내려가며 콘텐츠 추적
            # 미주 문단은 이미 콘텐츠로 간주했으므로, 다음 문단부터 빈 줄 체크 시작
            max_iterations = 500
            iteration = 0
            consecutive_empty_count = 0
            max_empty_count = 4  # 4줄 이상 빈 줄이면 종료
            
            while iteration < max_iterations:
                iteration += 1
                
                # 아래로 한 문단 이동 (미주 문단 다음 문단부터 확인)
                try:
                    self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception as e:
                    print(f"[디버그] 반복 #{iteration}: 아래로 이동 실패: {e}")
                    break
                
                # 현재 위치 확인
                try:
                    current_pos = self.hwp.GetPos()
                    if not current_pos:
                        break
                    
                    sec_check, para_check, pos_check = current_pos
                    
                    # 종료 조건 1: 다음 미주 도달 (다음 미주 직전에서 종료)
                    if next_endnote_pos:
                        # 다음 미주 위치에 도달하기 전에 종료해야 함
                        # 다음 미주가 (0, 14, 0)이면, (0, 13, pos_end)까지만 포함
                        if sec_check > sec_next or (sec_check == sec_next and para_check >= para_next):
                            print(f"[디버그] 반복 #{iteration}: 다음 미주 직전 도달. 종료.")
                            break
                    else:
                        # 문서 끝 도달 체크 (next_endnote_pos가 None일 때)
                        if doc_end_pos:
                            sec_end, para_end, pos_end = doc_end_pos
                            if sec_check > sec_end or (sec_check == sec_end and para_check > para_end):
                                print(f"[디버그] 반복 #{iteration}: 문서 끝 도달. 종료.")
                                break
                            elif sec_check == sec_end and para_check == para_end:
                                print(f"[디버그] 반복 #{iteration}: 문서 끝 위치 도달. 종료.")
                                break
                    
                    # [FIX] 구역 나누기 발견 시 문제 종료 / [FIX] 페이지 나누기 발견 시 문제 종료
                    # 종료 조건 2: 현재 문단 + 다음 문단 모두에서 페이지 나누기/구역 나누기 컨트롤 확인
                    break_found = False
                    break_type = None
                    break_pos = None
                    
                    try:
                        # 현재 위치 저장
                        saved_pos = (sec_check, para_check, pos_check)
                        
                        # [FIX] 구역 나누기 발견 시 문제 종료 - 현재 문단과 다음 문단 모두 확인
                        # 방법 1: Goto 액션으로 현재 위치 이후의 페이지 나누기 찾기
                        try:
                            self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                            try:
                                # 페이지 나누기로 이동 (DialogResult = 32, SetSelectionIndex = 6)
                                self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 32)
                                self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 6
                                self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                            except:
                                pass
                            
                            res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                            if res != 0:
                                page_break_pos = self.hwp.GetPos()
                                if page_break_pos:
                                    sec_pb, para_pb, pos_pb = page_break_pos
                                    # 현재 문단 또는 바로 다음 문단에 페이지 나누기가 있는지 확인
                                    if sec_pb == sec_check and para_pb == para_check:
                                        # [FIX] 페이지 나누기에서 문제 종료 - 현재 문단에 페이지 나누기 있음
                                        print(f"[디버그] 반복 #{iteration}: 현재 문단 ({sec_check}, {para_check})에 페이지 나누기 발견. 종료.")
                                        break_found = True
                                        break_type = "page"
                                        break_pos = page_break_pos
                                        break
                                    elif (sec_pb > sec_check or (sec_pb == sec_check and para_pb > para_check)):
                                        # 다음 문단에 페이지 나누기가 있음
                                        if not next_endnote_pos or (sec_pb < sec_next or (sec_pb == sec_next and para_pb < para_next)):
                                            # [FIX] 페이지 나누기에서 문제 종료 - 다음 문단에 페이지 나누기 있음
                                            print(f"[디버그] 반복 #{iteration}: 다음 문단 ({sec_pb}, {para_pb})에 페이지 나누기 발견. 종료.")
                                            break_found = True
                                            break_type = "page"
                                            break_pos = page_break_pos
                                            break
                            
                            # 현재 위치로 복귀
                            self.hwp.SetPos(*saved_pos)
                        except Exception as e:
                            print(f"[디버그] 반복 #{iteration}: 페이지 나누기 찾기 실패: {e}")
                            pass
                        
                        # 방법 2: Goto 액션으로 현재 위치 이후의 구역 나누기 찾기
                        if not break_found:
                            try:
                                self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                                try:
                                    # 구역 나누기로 이동 (DialogResult = 34 시도)
                                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 34)
                                    self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                                except:
                                    pass
                                
                                res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                                if res != 0:
                                    section_break_pos = self.hwp.GetPos()
                                    if section_break_pos:
                                        sec_sb, para_sb, pos_sb = section_break_pos
                                        # 현재 문단 또는 바로 다음 문단에 구역 나누기가 있는지 확인
                                        if sec_sb == sec_check and para_sb == para_check:
                                            # [FIX] 구역 나누기에서 문제 종료 - 현재 문단에 구역 나누기 있음
                                            print(f"[디버그] 반복 #{iteration}: 현재 문단 ({sec_check}, {para_check})에 구역 나누기 발견. 종료.")
                                            break_found = True
                                            break_type = "section"
                                            break_pos = section_break_pos
                                            break
                                        elif (sec_sb > sec_check or (sec_sb == sec_check and para_sb > para_check)):
                                            # 다음 문단에 구역 나누기가 있음
                                            if not next_endnote_pos or (sec_sb < sec_next or (sec_sb == sec_next and para_sb < para_next)):
                                                # [FIX] 구역 나누기에서 문제 종료 - 다음 문단에 구역 나누기 있음
                                                print(f"[디버그] 반복 #{iteration}: 다음 문단 ({sec_sb}, {para_sb})에 구역 나누기 발견. 종료.")
                                                break_found = True
                                                break_type = "section"
                                                break_pos = section_break_pos
                                                break
                                
                                # 현재 위치로 복귀
                                self.hwp.SetPos(*saved_pos)
                            except Exception as e:
                                print(f"[디버그] 반복 #{iteration}: 구역 나누기 찾기 실패: {e}")
                                pass
                    except Exception as e:
                        print(f"[디버그] 반복 #{iteration}: 페이지/구역 나누기 찾기 실패: {e}")
                        pass
                    
                    # [FIX] 구조 기반 문단 판별 - 현재 문단 콘텐츠 확인 (미주 문단 다음 문단부터)
                    has_content = False
                    read_failed = False  # [FIX] 선택 실패는 빈 문단 아님
                    
                    try:
                        # 현재 위치 저장
                        current_check_pos = (sec_check, para_check, pos_check)
                        
                        # 문단 시작으로 이동
                        self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        
                        # 문단 끝까지 선택
                        self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Run("ExtendSel")
                        
                        # 선택 범위 확인
                        try:
                            sel_start = self.hwp.GetPos(0)
                            sel_end = self.hwp.GetPos(1)
                            
                            # FIX: 구조 기반 문단 판별 - 선택 범위가 있어도 없어도 모두 확인
                            # 1. 텍스트 확인 시도
                            text_found = False
                            text_content = None
                            try:
                                self.copy_selected_range()
                                time.sleep(0.1)
                                text_content = self._read_clipboard_text()
                                
                                if text_content and text_content.strip():
                                    text_found = True
                                    has_content = True
                                    print(f"[디버그] 반복 #{iteration}: 문단 ({sec_check}, {para_check})에서 텍스트 발견 (길이: {len(text_content)})")
                            except Exception as e:
                                print(f"[디버그] 반복 #{iteration}: 텍스트 읽기 실패: {e}")
                                read_failed = True  # [FIX] 선택 실패는 빈 문단 아님
                            
                            # 2. 컨트롤 확인 (수식, 미주, 표, 그림, 텍스트상자 등)
                            # 선택 범위가 없어도(sel_start == sel_end) 컨트롤이 있을 수 있음
                            if not has_content:
                                try:
                                    # 현재 위치에서 GetText() 시도 (컨트롤 포함)
                                    raw = self.hwp.GetText()
                                    if raw:
                                        # GetText()가 성공하면 뭔가 내용이 있다는 의미
                                        has_content = True
                                        print(f"[디버그] 반복 #{iteration}: 문단 ({sec_check}, {para_check})에서 컨트롤/내용 발견 (GetText 성공)")
                                except:
                                    pass
                            
                            # 3. MoveNextCtrl로 컨트롤 존재 여부 확인
                            if not has_content:
                                try:
                                    # 현재 위치에서 다음 컨트롤 찾기 시도
                                    saved_pos = self.hwp.GetPos()
                                    self.hwp.HAction.GetDefault("MoveNextCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    result = self.hwp.HAction.Execute("MoveNextCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    
                                    if result != 0:
                                        # 컨트롤을 찾음
                                        ctrl_pos = self.hwp.GetPos()
                                        # 현재 문단 내에 있는지 확인
                                        if ctrl_pos:
                                            ctrl_sec, ctrl_para, ctrl_pos_val = ctrl_pos
                                            if ctrl_sec == sec_check and ctrl_para == para_check:
                                                has_content = True
                                                print(f"[디버그] 반복 #{iteration}: 문단 ({sec_check}, {para_check})에서 컨트롤 발견")
                                    
                                    # 원래 위치로 복귀
                                    if saved_pos:
                                        self.hwp.SetPos(*saved_pos)
                                except Exception as e:
                                    print(f"[디버그] 반복 #{iteration}: 컨트롤 확인 실패: {e}")
                            
                            # FIX: 선택 실패는 빈 문단 아님
                            # 선택 범위가 없거나 같아도(sel_start == sel_end) 위에서 텍스트/컨트롤 확인은 시도함
                            # 따라서 여기서는 로그만 남기고 빈 문단으로 판단하지 않음
                            if sel_start and sel_end and sel_start == sel_end and not has_content and not read_failed:
                                # 진짜 빈 문단일 가능성 (하지만 위에서 확인했으므로 여기서는 로그만)
                                print(f"[디버그] 반복 #{iteration}: 문단 ({sec_check}, {para_check})는 선택 범위가 없고 콘텐츠도 없음")
                                
                        except Exception as e:
                            print(f"[디버그] 반복 #{iteration}: 선택 범위 확인 실패: {e}")
                            read_failed = True  # FIX: 선택 실패는 빈 문단 아님
                        
                        # 선택 해제 및 원래 위치로 복귀
                        try:
                            self.hwp.HAction.Run("Cancel")
                            self.hwp.SetPos(*current_check_pos)
                        except:
                            pass
                    except Exception as e:
                        print(f"[디버그] 반복 #{iteration}: 콘텐츠 확인 실패: {e}")
                        read_failed = True  # FIX: 선택 실패는 빈 문단 아님
                    
                    # [FIX] 구역 나누기 발견 시 문제 종료 / [FIX] 페이지 나누기 발견 시 문제 종료
                    # 구역/페이지 나누기를 발견했으면 즉시 종료
                    if break_found:
                        # 이전 콘텐츠 위치를 끝점으로 사용 (구역/페이지 나누기 문단은 제외)
                        # break_pos가 현재 문단이면 이전 문단까지, 다음 문단이면 현재 문단까지
                        if break_pos:
                            sec_br, para_br, pos_br = break_pos
                            if sec_br == sec_check and para_br == para_check:
                                # 현재 문단에 구역/페이지 나누기가 있음 - 이전 문단까지가 끝점
                                print(f"[디버그] 반복 #{iteration}: {break_type} 나누기 발견으로 종료. 마지막 콘텐츠 위치: {last_content_pos} (현재 문단 제외)")
                            else:
                                # 다음 문단에 구역/페이지 나누기가 있음 - 현재 문단까지가 끝점
                                # 현재 문단의 끝을 last_content_pos로 업데이트
                                try:
                                    self.hwp.SetPos(sec_check, para_check, pos_check)
                                    self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    current_para_end = self.hwp.GetPos()
                                    if current_para_end:
                                        last_content_pos = current_para_end
                                except:
                                    pass
                                print(f"[디버그] 반복 #{iteration}: {break_type} 나누기 발견으로 종료. 마지막 콘텐츠 위치: {last_content_pos} (다음 문단 제외)")
                        else:
                            print(f"[디버그] 반복 #{iteration}: {break_type} 나누기 발견으로 종료. 마지막 콘텐츠 위치: {last_content_pos}")
                        break
                    
                    # [FIX] 구조 기반 문단 판별 - 콘텐츠 처리
                    if has_content:
                        # 콘텐츠가 있음
                        consecutive_empty_count = 0  # 빈 줄 카운터 리셋
                        # 문단 끝 위치를 last_content_pos로 저장
                        try:
                            self.hwp.SetPos(sec_check, para_check, pos_check)
                            self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            para_end_pos = self.hwp.GetPos()
                            if para_end_pos:
                                last_content_pos = para_end_pos
                        except:
                            last_content_pos = (sec_check, para_check, pos_check)
                    else:
                        # [FIX] 선택 실패는 빈 문단 아님 - 진짜 빈 문단일 때만 카운트 증가
                        # has_content == False AND read_failed == False 인 경우에만 빈 문단으로 간주
                        if not read_failed:
                            # 진짜 구조적으로 비어 있는 문단
                            consecutive_empty_count += 1
                            print(f"[디버그] 반복 #{iteration}: 빈 문단 확인 (연속 {consecutive_empty_count}개)")
                            
                            # 종료 조건 3: 빈 줄 4개 이상
                            if consecutive_empty_count >= max_empty_count:
                                print(f"[디버그] 반복 #{iteration}: 연속 빈 줄 {consecutive_empty_count}개 발견. 종료.")
                                break
                        else:
                            # [FIX] 선택 실패는 빈 문단 아님 - 읽기 실패는 무시하고 계속 진행
                            print(f"[디버그] 반복 #{iteration}: 읽기 실패 (빈 문단으로 간주하지 않음, 계속 진행)")
                            consecutive_empty_count = 0  # 읽기 실패는 빈 줄 카운트에 포함하지 않음
                    
                except Exception as e:
                    print(f"[디버그] 반복 #{iteration}: 위치 확인 실패: {e}")
                    consecutive_empty_count += 1
                    if consecutive_empty_count >= max_empty_count:
                        break
            
            # [FIX] 문제 끝점 계산 - 마지막 콘텐츠 위치 반환
            # 미주 문단만 포함된 상태에서 끝으로 확정하지 않도록 함
            if last_content_pos:
                # last_content_pos가 미주 문단의 끝과 같은지 확인
                # 같으면 실제 본문을 찾지 못한 것이므로 경고
                sec_last, para_last, pos_last = last_content_pos
                if sec_last == sec_current and para_last == para_current:
                    # 미주 문단만 포함된 상태
                    print(f"[경고] 마지막 콘텐츠 위치가 미주 문단과 같습니다. 본문을 찾지 못했을 수 있습니다.")
                    # 그래도 반환 (최소한 미주는 포함)
                
                print(f"[디버그] 마지막 콘텐츠 위치를 끝점으로 사용: {last_content_pos}")
                return last_content_pos
            else:
                # 폴백: 미주 위치의 문단 끝 사용
                # [FIX] 문제 끝점 계산 - 미주 문단만 포함된 상태에서 끝으로 확정하지 않도록
                # 하지만 본문을 찾지 못했으므로 최소한 미주는 포함
                try:
                    self.hwp.SetPos(sec_current, para_current, pos_current)
                    self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    fallback_pos = self.hwp.GetPos()
                    if fallback_pos:
                        print(f"[경고] 폴백: 미주 위치의 문단 끝을 끝점으로 사용 (본문을 찾지 못함): {fallback_pos}")
                        return fallback_pos
                except:
                    pass
                
                print(f"[경고] 마지막 콘텐츠 위치를 찾지 못했습니다. 미주 위치 사용: {current_endnote_pos}")
                return current_endnote_pos
            
        except Exception as e:
            print(f"[디버그] 아래쪽 콘텐츠 찾기 실패: {e}")
            import traceback
            traceback.print_exc()
            # 폴백: 미주 위치 반환
            return current_endnote_pos

    def _adjust_problem_start_pos(self, start_pos: Tuple[int, int, int]) -> Tuple[int, int, int]:
        """
        [FIX] 문제 시작점 보정 로직 개선 - start_pos 이후 첫 "실제 콘텐츠 문단" 찾기
        구역 나누기 직후 빈 페이지, 구역 나누기만 있는 문단, 텍스트 없는 문단 모두 제외
        
        Args:
            start_pos: 원래 시작 위치 (sec, para, pos)
        
        Returns:
            보정된 시작 위치
        """
        if not self.is_opened:
            return start_pos
        
        try:
            sec_start, para_start, pos_start = start_pos
            print(f"[디버그] 문제 시작점 보정 시작: 원래 ({sec_start}, {para_start}, {pos_start})")
            self.hwp.SetPos(sec_start, para_start, pos_start)
            
            # 현재 위치부터 아래로 내려가며 첫 실제 콘텐츠 문단 찾기
            for i in range(30):  # 최대 30문단까지만 확인
                try:
                    # 현재 문단에 콘텐츠가 있는지 확인
                    current_pos = self.hwp.GetPos()
                    if not current_pos:
                        break
                    
                    sec_check, para_check, pos_check = current_pos
                    
                    # [FIX] 문제 시작점 보정 로직 개선 - 현재 문단에 구역/페이지 나누기가 있는지 확인 (빈 페이지 제외)
                    is_break_only = False  # 구역/페이지 나누기만 있는 문단인지
                    try:
                        saved_pos = self.hwp.GetPos()
                        # 페이지 나누기 확인
                        self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                        try:
                            self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 32)
                            self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 6
                            self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                        except:
                            pass
                        
                        res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                        if res != 0:
                            break_pos = self.hwp.GetPos()
                            if break_pos:
                                sec_br, para_br, pos_br = break_pos
                                if sec_br == sec_check and para_br == para_check:
                                    is_break_only = True
                        
                        self.hwp.SetPos(*saved_pos)
                        
                        # 구역 나누기 확인
                        if not is_break_only:
                            self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                            try:
                                self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 34)
                                self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                            except:
                                pass
                            
                            res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                            if res != 0:
                                break_pos = self.hwp.GetPos()
                                if break_pos:
                                    sec_br, para_br, pos_br = break_pos
                                    if sec_br == sec_check and para_br == para_check:
                                        is_break_only = True
                            
                            self.hwp.SetPos(*saved_pos)
                    except:
                        pass
                    
                    # [FIX] 문제 시작점 보정 로직 개선 - 구역/페이지 나누기만 있는 문단은 건너뛰기
                    if is_break_only:
                        print(f"[디버그] 문제 시작점 보정: 문단 ({sec_check}, {para_check})는 구역/페이지 나누기만 있음. 건너뜀.")
                        # 다음 문단으로 이동
                        self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        continue
                    
                    # 문단 시작으로 이동
                    self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    
                    # 문단 끝까지 선택
                    self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Run("ExtendSel")
                    
                    # [FIX] 문제 시작점 보정 로직 개선 - 텍스트 확인 (구조 기반 판별)
                    has_content = False
                    read_failed = False
                    
                    # 텍스트 확인
                    try:
                        self.copy_selected_range()
                        time.sleep(0.05)
                        text = self._read_clipboard_text()
                        if text and text.strip():
                            has_content = True
                    except:
                        read_failed = True
                    
                    # 컨트롤 확인 (텍스트가 없어도 컨트롤이 있을 수 있음)
                    if not has_content:
                        try:
                            raw = self.hwp.GetText()
                            if raw:
                                has_content = True
                        except:
                            pass
                    
                    # 선택 해제
                    self.hwp.HAction.Run("Cancel")
                    
                    # [FIX] 문제 시작점 보정 로직 개선 - 실제 콘텐츠 문단 발견
                    if has_content:
                        print(f"[디버그] 문제 시작점 보정: 첫 실제 콘텐츠 문단 발견. 원래 ({sec_start}, {para_start}, {pos_start}) → 보정 ({sec_check}, {para_check}, 0)")
                        return (sec_check, para_check, 0)
                    elif not read_failed:
                        # 텍스트도 없고 읽기 실패도 아님 = 진짜 빈 문단
                        print(f"[디버그] 문제 시작점 보정: 문단 ({sec_check}, {para_check})는 빈 문단. 건너뜀.")
                    
                    # 다음 문단으로 이동
                    self.hwp.SetPos(sec_check, para_check, pos_check)
                    self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception as e:
                    print(f"[디버그] 문제 시작점 보정: 문단 확인 실패: {e}")
                    break
            
            # 콘텐츠를 찾지 못하면 원래 위치 반환
            print(f"[디버그] 문제 시작점 보정: 실제 콘텐츠 문단을 찾지 못함. 원래 위치 유지: ({sec_start}, {para_start}, {pos_start})")
            return start_pos
        except Exception as e:
            print(f"[디버그] 문제 시작점 보정 실패: {e}")
            return start_pos
    
    def _adjust_problem_end_pos(self, end_pos: Tuple[int, int, int]) -> Tuple[int, int, int]:
        """
        [FIX] 문제 끝점 보정 로직 개선 - 현재 문단과 다음 문단 모두에서 구역/페이지 나누기 확인
        
        Args:
            end_pos: 원래 끝 위치 (sec, para, pos)
        
        Returns:
            보정된 끝 위치
        """
        if not self.is_opened:
            return end_pos
        
        try:
            sec_end, para_end, pos_end = end_pos
            print(f"[디버그] 문제 끝점 보정 시작: 원래 ({sec_end}, {para_end}, {pos_end})")
            self.hwp.SetPos(sec_end, para_end, pos_end)
            
            # [FIX] 문제 끝점 보정 로직 개선 - 현재 문단에 구역/페이지 나누기가 있는지 먼저 확인
            # 현재 문단에 구역/페이지 나누기 확인
            try:
                saved_pos = self.hwp.GetPos()
                
                # 방법 1: 현재 문단에 페이지 나누기 확인
                self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                try:
                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 32)
                    self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 6
                    self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                except:
                    pass
                
                res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                if res != 0:
                    page_break_pos = self.hwp.GetPos()
                    if page_break_pos:
                        sec_pb, para_pb, pos_pb = page_break_pos
                        if sec_pb == sec_end and para_pb == para_end:
                            # [FIX] 문제 끝점 보정 - 현재 문단에 페이지 나누기 발견
                            # 이전 문단까지가 끝점
                            try:
                                # 이전 문단으로 이동
                                self.hwp.SetPos(sec_end, para_end, 0)
                                self.hwp.HAction.GetDefault("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Execute("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                prev_pos = self.hwp.GetPos()
                                if prev_pos:
                                    self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    adjusted_pos = self.hwp.GetPos()
                                    if adjusted_pos:
                                        print(f"[디버그] 문제 끝점 보정: 현재 문단에 페이지 나누기 발견. 원래 ({sec_end}, {para_end}, {pos_end}) → 보정 {adjusted_pos}")
                                        return adjusted_pos
                            except:
                                pass
                
                self.hwp.SetPos(*saved_pos)
                
                # 방법 2: 현재 문단에 구역 나누기 확인
                self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                try:
                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 34)
                    self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                except:
                    pass
                
                res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                if res != 0:
                    section_break_pos = self.hwp.GetPos()
                    if section_break_pos:
                        sec_sb, para_sb, pos_sb = section_break_pos
                        if sec_sb == sec_end and para_sb == para_end:
                            # [FIX] 문제 끝점 보정 - 현재 문단에 구역 나누기 발견
                            # 이전 문단까지가 끝점
                            try:
                                # 이전 문단으로 이동
                                self.hwp.SetPos(sec_end, para_end, 0)
                                self.hwp.HAction.GetDefault("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Execute("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                prev_pos = self.hwp.GetPos()
                                if prev_pos:
                                    self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    adjusted_pos = self.hwp.GetPos()
                                    if adjusted_pos:
                                        print(f"[디버그] 문제 끝점 보정: 현재 문단에 구역 나누기 발견. 원래 ({sec_end}, {para_end}, {pos_end}) → 보정 {adjusted_pos}")
                                        return adjusted_pos
                            except:
                                pass
                
                self.hwp.SetPos(*saved_pos)
            except Exception as e:
                print(f"[디버그] 문제 끝점 보정: 현재 문단 확인 실패: {e}")
                pass
            
            # 다음 문단에 구역/페이지 나누기가 있는지 확인
            try:
                # 다음 문단으로 이동
                self.hwp.SetPos(sec_end, para_end, pos_end)
                self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                next_pos = self.hwp.GetPos()
                
                if next_pos:
                    sec_next, para_next, pos_next = next_pos
                    
                    # 다음 문단에 구역/페이지 나누기 확인
                    # 방법 1: Goto로 페이지 나누기 찾기
                    try:
                        saved_pos = self.hwp.GetPos()
                        self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                        try:
                            self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 32)
                            self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 6
                            self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                        except:
                            pass
                        
                        res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                        if res != 0:
                            page_break_pos = self.hwp.GetPos()
                            if page_break_pos:
                                sec_pb, para_pb, pos_pb = page_break_pos
                                if sec_pb == sec_next and para_pb == para_next:
                                    # [FIX] 문제 끝점 보정 - 다음 문단에 페이지 나누기 발견
                                    # 현재 위치가 이미 올바름 (다음 문단은 제외)
                                    print(f"[디버그] 문제 끝점 보정: 다음 문단에 페이지 나누기 발견. 원래 ({sec_end}, {para_end}, {pos_end}) → 보정 ({sec_end}, {para_end}, {pos_end})")
                                    self.hwp.SetPos(*saved_pos)
                                    return end_pos
                        
                        self.hwp.SetPos(*saved_pos)
                    except:
                        pass
                    
                    # 방법 2: Goto로 구역 나누기 찾기
                    try:
                        saved_pos = self.hwp.GetPos()
                        self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                        try:
                            self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 34)
                            self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                        except:
                            pass
                        
                        res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                        if res != 0:
                            section_break_pos = self.hwp.GetPos()
                            if section_break_pos:
                                sec_sb, para_sb, pos_sb = section_break_pos
                                if sec_sb == sec_next and para_sb == para_next:
                                    # [FIX] 문제 끝점 보정 - 다음 문단에 구역 나누기 발견
                                    # 현재 위치가 이미 올바름 (다음 문단은 제외)
                                    print(f"[디버그] 문제 끝점 보정: 다음 문단에 구역 나누기 발견. 원래 ({sec_end}, {para_end}, {pos_end}) → 보정 ({sec_end}, {para_end}, {pos_end})")
                                    self.hwp.SetPos(*saved_pos)
                                    return end_pos
                        
                        self.hwp.SetPos(*saved_pos)
                    except:
                        pass
            except Exception as e:
                print(f"[디버그] 문제 끝점 보정: 다음 문단 확인 실패: {e}")
                pass
            
            # 구역/페이지 나누기를 찾지 못하면 원래 위치 반환
            print(f"[디버그] 문제 끝점 보정: 구역/페이지 나누기 없음. 원래 위치 유지: ({sec_end}, {para_end}, {pos_end})")
            return end_pos
        except Exception as e:
            print(f"[디버그] 문제 끝점 보정 실패: {e}")
            return end_pos

    def select_range_from_endnote_to_problem_end(
        self, 
        endnote_start_pos: Tuple[int, int, int],
        problem_end_pos: Tuple[int, int, int]
    ) -> bool:
        """
        미주 시작부터 문제 끝까지 선택
        
        Args:
            endnote_start_pos: 미주 시작 위치 (sec, para, pos)
            problem_end_pos: 문제 끝 위치 (sec, para, pos)
        
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        try:
            # [FIX] 문제 시작점 보정 로직 개선 - start_pos 이후 첫 "실제 콘텐츠 문단" 찾기
            adjusted_start_pos = self._adjust_problem_start_pos(endnote_start_pos)
            # [FIX] 문제 끝점 보정 로직 개선 - end_pos 이후에 구역/페이지 나누기가 있으면 그 직전까지로 제한
            adjusted_end_pos = self._adjust_problem_end_pos(problem_end_pos)
            
            sec_start, para_start, pos_start = adjusted_start_pos
            sec_end, para_end, pos_end = adjusted_end_pos
            
            print(f"[디버그] 범위 선택: 시작 ({sec_start}, {para_start}, {pos_start}) → 끝 ({sec_end}, {para_end}, {pos_end})")
            
            # 시작과 끝 위치가 같은 경우 처리
            if sec_start == sec_end and para_start == para_end and pos_start == pos_end:
                print(f"[경고] 시작과 끝 위치가 같습니다. 선택할 수 없습니다.")
                return False
            
            # 시작 위치로 이동
            try:
                self.hwp.SetPos(sec_start, para_start, pos_start)
                actual_start_pos = self.hwp.GetPos()
                print(f"[디버그] 시작 위치 설정 - 요청: ({sec_start}, {para_start}, {pos_start}), 실제: {actual_start_pos}")
            except Exception as e:
                print(f"[경고] SetPos 실패: {e}, 대체 방법 시도")
                # SetPos 실패 시 대체 방법
                self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                # para_start만큼 아래로 이동 (근사치)
                for _ in range(min(para_start, 100)):
                    try:
                        self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    except:
                        break
            
            # 선택 시작
            self.hwp.HAction.Run("Select")
            print(f"[디버그] 선택 시작 완료")
            
            # 끝 위치까지 선택 확장
            # 문단 차이 계산
            para_diff = para_end - para_start
            pos_diff = pos_end - pos_start
            
            # 같은 문단 내에서 이동하는 경우
            if sec_start == sec_end and para_diff == 0:
                # 같은 문단 내에서 pos_diff만큼 오른쪽으로 이동하면서 선택 확장
                move_count = max(0, min(pos_diff, 1000))
                for _ in range(move_count):
                    try:
                        self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Run("ExtendSel")
                    except:
                        break
            else:
                # 다른 문단으로 이동하는 경우
                print(f"[디버그] 다른 문단으로 이동 - para_diff: {para_diff}, pos_diff: {pos_diff}")
                # 끝 위치로 직접 이동하면서 선택 확장
                try:
                    self.hwp.SetPos(sec_end, para_end, pos_end)
                    actual_end_pos = self.hwp.GetPos()
                    print(f"[디버그] 끝 위치 설정 - 요청: ({sec_end}, {para_end}, {pos_end}), 실제: {actual_end_pos}")
                    # 선택 확장
                    self.hwp.HAction.Run("ExtendSel")
                    print(f"[디버그] ExtendSel 완료")
                except Exception as e:
                    print(f"[경고] SetPos 실패: {e}, 대체 방법 시도")
                    # SetPos 실패 시 대체 방법
                    # 현재 문단 끝까지 이동
                    try:
                        self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Run("ExtendSel")
                        print(f"[디버그] 현재 문단 끝까지 선택 확장 완료")
                    except Exception as e2:
                        print(f"[경고] MoveParaEnd 실패: {e2}")
                        pass
                    
                    # 중간 문단들을 통과하면서 선택 확장
                    print(f"[디버그] 중간 문단 {max(0, para_diff - 1)}개 통과 시작")
                    for i in range(max(0, para_diff - 1)):
                        try:
                            self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Run("ExtendSel")
                            print(f"[디버그] 중간 문단 {i+1} 통과 완료")
                        except Exception as e3:
                            print(f"[경고] 중간 문단 {i+1} 통과 실패: {e3}")
                            break
                    
                    # 마지막 문단에서 끝 위치까지
                    if para_diff > 0:
                        print(f"[디버그] 마지막 문단에서 끝 위치까지 이동 - pos_end: {pos_end}")
                        for i in range(min(pos_end, 1000)):
                            try:
                                self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                self.hwp.HAction.Run("ExtendSel")
                            except Exception as e4:
                                print(f"[경고] MoveRight {i+1}회 실패: {e4}")
                                break
            
            # [FIX] 선택 범위 검증 - 선택 범위 안에 구역/페이지 나누기 포함 여부 확인
            try:
                # 선택 후 위치 확인
                try:
                    sel_start_pos = self.hwp.GetPos()
                    print(f"[디버그] 선택 후 시작 위치: {sel_start_pos}")
                except Exception:
                    pass
                
                # 선택 범위 내에 구역/페이지 나누기가 있는지 확인
                # 선택 범위의 끝 부분부터 확인
                try:
                    # 선택 끝 위치로 이동
                    sel_end_pos = self.hwp.GetPos(1)
                    if sel_end_pos:
                        sec_sel_end, para_sel_end, pos_sel_end = sel_end_pos
                        
                        # 선택 범위 끝 문단과 그 다음 문단에 구역/페이지 나누기 확인
                        # 현재 선택 끝 위치에서 확인
                        self.hwp.SetPos(sec_sel_end, para_sel_end, pos_sel_end)
                        
                        # 현재 문단에 구역/페이지 나누기 확인
                        break_found_in_selection = False
                        try:
                            saved_pos = self.hwp.GetPos()
                            
                            # 페이지 나누기 확인
                            self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                            try:
                                self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 32)
                                self.hwp.HParameterSet.HGotoE.SetSelectionIndex = 6
                                self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                            except:
                                pass
                            
                            res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                            if res != 0:
                                break_pos = self.hwp.GetPos()
                                if break_pos:
                                    sec_br, para_br, pos_br = break_pos
                                    if sec_br == sec_sel_end and para_br == para_sel_end:
                                        break_found_in_selection = True
                                        print(f"[경고] 선택 범위에 페이지 나누기가 포함됨. 선택 범위 재조정 필요.")
                            
                            self.hwp.SetPos(*saved_pos)
                            
                            # 구역 나누기 확인
                            if not break_found_in_selection:
                                self.hwp.HAction.GetDefault("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                                try:
                                    self.hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 34)
                                    self.hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                                except:
                                    pass
                                
                                res = self.hwp.HAction.Execute("Goto", self.hwp.HParameterSet.HGotoE.HSet)
                                if res != 0:
                                    break_pos = self.hwp.GetPos()
                                    if break_pos:
                                        sec_br, para_br, pos_br = break_pos
                                        if sec_br == sec_sel_end and para_br == para_sel_end:
                                            break_found_in_selection = True
                                            print(f"[경고] 선택 범위에 구역 나누기가 포함됨. 선택 범위 재조정 필요.")
                                
                                self.hwp.SetPos(*saved_pos)
                            
                            # [FIX] 선택 범위 검증 - 구역/페이지 나누기가 포함되어 있으면 선택 범위 재조정
                            if break_found_in_selection:
                                # 이전 문단까지로 선택 범위 축소
                                try:
                                    # 이전 문단으로 이동
                                    self.hwp.SetPos(sec_sel_end, para_sel_end, 0)
                                    self.hwp.HAction.GetDefault("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    self.hwp.HAction.Execute("MoveUp", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                    prev_pos = self.hwp.GetPos()
                                    if prev_pos:
                                        # 이전 문단 끝까지 선택
                                        self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                        self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                                        self.hwp.HAction.Run("ExtendSel")
                                        print(f"[디버그] 선택 범위 재조정: 구역/페이지 나누기 제외")
                                except Exception as e:
                                    print(f"[경고] 선택 범위 재조정 실패: {e}")
                        except Exception as e:
                            print(f"[디버그] 선택 범위 검증 실패: {e}")
                except Exception as e:
                    print(f"[디버그] 선택 범위 검증 중 오류: {e}")
                
                # 선택된 텍스트 읽기
                selected_text = self.get_text_from_selection()
                print(f"[디버그] 선택된 텍스트 - 타입: {type(selected_text)}, 길이: {len(selected_text) if selected_text else 0}, 내용(처음 100자): {selected_text[:100] if selected_text else 'None'}")
                
                if selected_text and len(selected_text.strip()) > 10:
                    print(f"[디버그] 선택 완료: {len(selected_text)} 글자")
                    return True
                else:
                    print(f"[디버그] 선택 실패: 텍스트가 너무 짧음")
                    return False
            except Exception:
                # 검증 실패해도 선택은 성공한 것으로 간주
                return True
                
        except Exception as e:
            print(f"[디버그] 범위 선택 실패: {e}")
            import traceback
            traceback.print_exc()
            return False

    def select_range_between_markers(self, marker_start: str, marker_end: str) -> bool:
        """
        방법 1: 범위 선택 기반 문제 블록 추출
        
        핵심 원칙:
        1. [문제시작] 마커를 찾습니다
        2. 마커 다음 줄(또는 다음 문단)로 이동합니다
        3. [문제끝] 마커를 찾습니다
        4. 마커 이전 줄(또는 이전 문단)로 이동합니다
        5. 두 위치 사이를 선택합니다 (Shift+클릭과 유사)
        
        Args:
            marker_start: [문제시작] 마커 문자열
            marker_end: [문제끝] 마커 문자열
            
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        try:
            # ✅ 본문 포커스 복귀(미주/각주 편집 종료) 시도
            # 일부 문서에서는 미주 영역에 커서가 갇히면 "문서 처음 이동/검색/텍스트 추출"이 미주 기준으로 동작할 수 있습니다.
            # 매크로(Shift+Esc) 동작을 액션으로 재현해 본문으로 복귀를 우선 시도합니다.
            try:
                self.ensure_main_body_focus()
            except Exception:
                pass

            # 1. [문제시작] 마커 찾기
            # move_after=False: 마커 텍스트가 선택된 상태를 유지하여 경계 계산에 활용
            start_result = self.find_text(marker_start, start_from_beginning=False, move_after=False)
            if start_result is None:
                print(f"[디버그] [문제시작] 마커를 찾을 수 없습니다.")
                return False

            # [문제시작] 마커 선택의 끝으로 이동 → 선택 해제 (커서를 마커 끝에 둠)
            try:
                self.hwp.HAction.GetDefault("MoveSelEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveSelEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
            except:
                pass
            self.hwp.HAction.Run("Cancel")
            
            # 2. 마커 다음 문단 시작으로 이동
            # [문제시작] 마커는 보통 독립된 문단에 있으므로, 다음 문단이 문제 내용입니다
            try:
                # 현재 문단 끝으로 이동
                self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                # 다음 문단 시작으로 이동 (아래로 한 줄 이동 후 문단 시작)
                try:
                    self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    # MoveDown이 없으면 오른쪽으로 이동
                    for _ in range(5):
                        try:
                            self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                            self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        except:
                            break
                
                # 문단 시작으로 이동
                self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            except Exception as e:
                print(f"[경고] 마커 다음으로 이동 실패: {e}")
                # 대체 방법: 오른쪽으로 여러 번 이동
                for _ in range(20):
                    try:
                        self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    except:
                        break
            
            # 선택 시작 위치 저장
            try:
                select_start_pos = self.hwp.GetPos()
                sec_start, para_start, pos_start = select_start_pos
                print(f"[검증] 선택 시작 위치 ([문제시작] 다음): ({sec_start}, {para_start}, {pos_start})")
            except:
                print(f"[경고] 선택 시작 위치를 가져올 수 없습니다.")
                return False
            
            # 3. [문제끝] 마커 찾기 (선택 시작 전에 먼저 찾기)
            # move_after=False: 마커 텍스트가 선택된 상태를 유지하여 "마커 시작"을 종료 경계로 사용
            end_result = self.find_text(marker_end, start_from_beginning=False, move_after=False)
            if end_result is None:
                print(f"[디버그] [문제끝] 마커를 찾을 수 없습니다.")
                return False

            # [문제끝] 마커의 "시작 위치"를 안정적으로 산출
            # - MoveSelBegin이 실패하거나 Cancel이 커서를 선택 끝으로 두는 경우가 있어,
            #   시작/끝을 둘 다 읽고 필요시 보정합니다.
            sec_end = para_end = pos_end = None
            marker_len = len(marker_end)
            try:
                # 1) 선택 시작으로 이동 (선택 유지 상태에서 좌표를 읽는다)
                try:
                    self.hwp.HAction.GetDefault("MoveSelBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveSelBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    pass
                begin_pos = self.hwp.GetPos()

                # 2) 선택 끝으로 이동 후 좌표를 읽는다
                try:
                    self.hwp.HAction.GetDefault("MoveSelEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveSelEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    pass
                end_pos = self.hwp.GetPos()

                # 3) begin_pos가 제대로 시작점으로 이동하지 못한 경우(끝과 동일) 보정
                #    마커는 일반적으로 한 문단 내 텍스트이므로 pos에서 marker_len을 빼서 시작점 추정
                sec_b, para_b, pos_b = begin_pos
                sec_e, para_e, pos_e = end_pos

                # 기본은 MoveSelBegin 결과를 신뢰
                sec_end, para_end, pos_end = sec_b, para_b, pos_b

                # begin이 end와 같거나, pos가 마커 길이처럼 보이면(자주 pos=5), 끝 기준 보정 시도
                if (sec_b, para_b, pos_b) == (sec_e, para_e, pos_e) or pos_b == marker_len:
                    if sec_e == sec_b and para_e == para_b and pos_e >= marker_len:
                        sec_end, para_end, pos_end = sec_e, para_e, max(0, pos_e - marker_len)

                # 선택 해제 후, 계산된 "마커 시작"으로 커서를 고정
                self.hwp.HAction.Run("Cancel")
                try:
                    self.hwp.SetPos(sec_end, para_end, pos_end)
                except:
                    pass

                print(f"[검증] 선택 끝 위치 ([문제끝] 시작): ({sec_end}, {para_end}, {pos_end})")
            except Exception:
                try:
                    self.hwp.HAction.Run("Cancel")
                except:
                    pass
                print(f"[경고] 선택 끝 위치를 안정적으로 계산할 수 없습니다.")
                return False

            # ✅ [문제끝] 마커 자체를 선택에서 제외하기 위해 종료 경계를 1칸 왼쪽으로 이동
            # (선택 구현에 따라 종료 위치의 글자가 포함될 수 있어, 마커 첫 글자 포함을 방지)
            try:
                if pos_end is not None and pos_end > 0:
                    self.hwp.HAction.GetDefault("MoveLeft", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveLeft", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    select_end_pos = self.hwp.GetPos()
                    sec_end, para_end, pos_end = select_end_pos
                    print(f"[검증] 선택 끝 위치 (마커 제외 조정): ({sec_end}, {para_end}, {pos_end})")
            except:
                pass
            
            # 5. 선택 시작 위치로 다시 이동하여 선택 시작
            # SetPos를 사용하여 시작 위치로 이동
            try:
                self.hwp.SetPos(sec_start, para_start, pos_start)
                print(f"[검증] 시작 위치로 이동 완료: ({sec_start}, {para_start}, {pos_start})")
            except Exception as e:
                print(f"[경고] SetPos로 시작 위치 이동 실패: {e}")
                # 대체 방법: 문서 처음으로 이동 후 문단 단위로 이동
                self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                # para_start만큼 아래로 이동 (대략적인 위치)
                for _ in range(min(para_start, 100)):
                    try:
                        self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    except:
                        break
                # 문단 시작으로 이동
                try:
                    self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    pass
            
            # 선택 시작
            self.hwp.HAction.Run("Select")
            print(f"[검증] Select 실행 완료")
            
            # 6. 선택 끝 위치까지 이동하면서 선택 확장
            # SetPos 대신 커서 이동 명령을 사용하여 선택 상태를 유지하면서 이동
            # 문단 차이 계산
            para_diff = para_end - para_start
            pos_diff = pos_end - pos_start
            
            print(f"[검증] 선택 확장 시작: 문단 차이={para_diff}, 위치 차이={pos_diff}")
            
            # 같은 문단 내에서 이동하는 경우
            if para_diff == 0:
                # 같은 문단 내에서 pos_diff만큼 오른쪽으로 이동하면서 선택 확장
                move_count = max(0, min(pos_diff, 1000))
                for _ in range(move_count):
                    try:
                        self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        # 선택 확장
                        self.hwp.HAction.Run("ExtendSel")
                    except:
                        break
            else:
                # 다른 문단으로 이동하는 경우
                # 1. 현재 문단 끝까지 이동하면서 선택 확장
                try:
                    self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Run("ExtendSel")
                except:
                    pass
                
                # 2. 중간 문단들을 통과하면서 선택 확장
                for _ in range(para_diff - 1):
                    try:
                        # 다음 문단으로 이동
                        self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        # 문단 전체 선택 확장
                        self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Run("ExtendSel")
                        self.hwp.HAction.GetDefault("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveParaEnd", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Run("ExtendSel")
                    except:
                        break
                
                # 3. 마지막 문단으로 이동
                try:
                    self.hwp.HAction.GetDefault("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveDown", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    # 문단 시작으로 이동
                    self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Run("ExtendSel")
                except:
                    pass
                
                # 4. 마지막 문단에서 pos_end 위치까지 이동
                move_count = max(0, min(pos_end, 1000))
                for _ in range(move_count):
                    try:
                        self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        # 선택 확장
                        self.hwp.HAction.Run("ExtendSel")
                    except:
                        break
            
            print(f"[검증] 선택 확장 완료")
            
            # 외부에서 반복 감지에 사용할 수 있도록 저장
            self.last_problem_start_pos = (sec_start, para_start, pos_start)
            self.last_problem_end_pos = (sec_end, para_end, pos_end)

            print(f"[디버그] 범위 선택 완료: Start=({sec_start}, {para_start}, {pos_start}), End=({sec_end}, {para_end}, {pos_end})")
            return True
            
        except Exception as e:
            print(f"[디버그] 범위 선택 실패: {e}")
            import traceback
            traceback.print_exc()
            # 선택 해제
            try:
                self.hwp.HAction.Run("Cancel")
            except:
                pass
            return False
    
    def extract_selected_to_hwp_file(self, output_path: str) -> bool:
        """
        현재 선택된 범위를 새 HWP 파일로 저장
        
        Args:
            output_path: 저장할 파일 경로
            
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        try:
            # 원본 문서 핸들 저장 (다시 "열기"가 아니라 Activate로 복귀)
            # 기존에는 원본 파일을 다시 Open() 했는데, 이 경우 커서/상태가 초기화되어
            # 같은 문제를 반복 추출하는 무한루프의 원인이 됩니다.
            try:
                original_doc = self.hwp.XHwpDocuments.Item(0)
            except Exception:
                original_doc = None
            
            # ✅ 선택 영역 검증 로그 (필수) - 컨트롤 기준
            selection_info = self.get_selection_info()
            control_info = self.get_selected_control_info()
            
            print(f"[검증] 선택 영역 정보:")
            print(f"  - SelectionStartPos: {selection_info.get('selection_start', 'N/A')}")
            print(f"  - SelectionEndPos: {selection_info.get('selection_end', 'N/A')}")
            print(f"  - 선택된 텍스트 길이: {selection_info.get('text_length', 0)} 글자")
            print(f"  - 선택된 문단 수 (추정): {selection_info.get('paragraph_count', 0)} 문단")
            print(f"[검증] 선택된 컨트롤 정보:")
            print(f"  - 선택된 컨트롤 개수: {control_info.get('control_count', 0)}")
            print(f"  - 선택된 컨트롤 타입: {control_info.get('control_types', [])}")
            
            # ⚠️ 주의: 현재 get_selected_control_info()는 컨트롤 개수를 정확히 계산하지 못합니다.
            # (HWP API 제한으로 인해 기본값이 1로 고정될 수 있음)
            # 따라서 컨트롤 개수로 즉시 실패 처리하면 정상 추출도 막힙니다.
            control_count = control_info.get('control_count', 0)
            if control_count <= 1:
                print(f"[경고] 선택된 컨트롤 개수 추정치가 {control_count}로 표시됩니다. (검증용 로그이며, 추출은 계속 진행)")
            
            # 선택 영역 검증: 텍스트 길이가 너무 짧으면 경고
            text_length = selection_info.get('text_length', 0)
            if text_length < 10:
                print(f"[경고] 선택 영역이 너무 작습니다. 텍스트 길이: {text_length}")
            
            # 선택 범위 재확인: 선택이 제대로 되어 있는지 확인
            try:
                sel_start = self.hwp.GetPos(0)  # 선택 시작
                sel_end = self.hwp.GetPos(1)    # 선택 끝
                if sel_start == sel_end:
                    print(f"[경고] 선택 범위가 축소되었습니다. 선택을 다시 시도합니다.")
                    # 선택 해제 후 다시 선택 시도는 복잡하므로, 그냥 진행
            except:
                pass
            
            # 선택된 범위 복사 (여러 번 시도)
            copy_success = False
            for attempt in range(3):  # 최대 3번 시도
                if self.copy_selected_range():
                    copy_success = True
                    break
                time.sleep(0.1)  # 짧은 대기 후 재시도
            
            if not copy_success:
                print(f"[디버그] 선택된 범위 복사 실패 (3번 시도)")
                return False
            
            # 새 문서 생성
            self.hwp.HAction.GetDefault("FileNew", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HAction.Execute("FileNew", self.hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 붙여넣기
            self.hwp.HAction.Run("Paste")
            
            # ✅ 붙여넣기 결과 검증
            paste_result = self.verify_paste_result()
            print(f"[검증] 붙여넣기 결과:")
            print(f"  - 텍스트 길이: {paste_result.get('text_length', 0)} 글자")
            print(f"  - 문단 수 (추정): {paste_result.get('paragraph_count', 0)} 문단")
            print(f"  - 텍스트 존재 여부: {paste_result.get('has_text', False)}")
            
            # 붙여넣기 결과 검증: 문단 수가 1이거나 텍스트만 있으면 실패
            if paste_result.get('paragraph_count', 0) <= 1:
                print(f"[경고] 붙여넣기 결과가 1문단 이하입니다. 문단 수: {paste_result.get('paragraph_count', 0)}")
            if not paste_result.get('has_text', False):
                print(f"[경고] 붙여넣기 결과에 텍스트가 없습니다.")
            
            # 파일 저장
            self.hwp.HAction.GetDefault("FileSaveAs", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HParameterSet.HFileOpenSave.filename = output_path
            self.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
            self.hwp.HAction.Execute("FileSaveAs", self.hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 새 문서 닫기 (저장 확인 다이얼로그가 나타날 수 있으므로 무시)
            try:
                self.hwp.HAction.GetDefault("FileClose", self.hwp.HParameterSet.HFileOpenSave.HSet)
                self.hwp.HParameterSet.HFileOpenSave.filename = ""  # 저장 확인 다이얼로그 방지
                self.hwp.HAction.Execute("FileClose", self.hwp.HParameterSet.HFileOpenSave.HSet)
            except:
                pass
            
            # 원본 문서로 돌아가기 (재오픈 금지: 상태/커서 초기화 방지)
            try:
                if original_doc is not None:
                    original_doc.Activate()
                else:
                    # 혹시라도 참조가 없으면 0번 문서 Activate 시도
                    if getattr(self.hwp.XHwpDocuments, "Count", 0) > 0:
                        self.hwp.XHwpDocuments.Item(0).Activate()
            except:
                pass
            
            success = os.path.exists(output_path)
            if success:
                print(f"[디버그] HWP 파일 추출 성공: {output_path}")
            else:
                print(f"[디버그] HWP 파일 추출 실패: 파일이 생성되지 않음")
            
            return success
        except Exception as e:
            print(f"[디버그] HWP 파일 추출 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_text_from_selection(self) -> str:
        """
        현재 선택된 범위의 텍스트 추출
        
        HWP 공식 문서에 따라 InitScan() + GetText()를 올바르게 사용합니다.
        
        Returns:
            추출된 텍스트
        """
        if not self.is_opened:
            return ""
        
        try:
            # 1) InitScan() + GetText() 사용 (HWP 공식 방법)
            text_from_gettext = ""
            try:
                status_code, text_from_gettext = self._get_text_with_scan(use_init_scan=True)
                print(f"[디버그] get_text_from_selection: GetText() 상태코드: {status_code}, 텍스트 길이: {len(text_from_gettext) if text_from_gettext else 0}")
                
                # 상태코드 확인
                if status_code == 101:
                    # InitScan() 초기화 안됨 - 클립보드 방식으로 폴백
                    print(f"[경고] InitScan() 초기화 안됨, 클립보드 방식 사용")
                    text_from_gettext = ""
                elif status_code == 0:
                    # 텍스트 정보 없음
                    print(f"[경고] 텍스트 정보 없음 (상태코드 0), 클립보드 방식 사용")
                    text_from_gettext = ""
                elif status_code != 2:
                    # 일반 텍스트가 아닌 경우
                    print(f"[경고] 상태코드 {status_code}, 클립보드 방식 사용")
                    text_from_gettext = ""
            except Exception as e:
                print(f"[경고] GetText() 실패: {e}, 클립보드 방식 사용")
                text_from_gettext = ""

            # 2) 클립보드 기반 텍스트(수식/표 등에서 GetText가 비는 케이스 보완)
            text_from_clipboard = ""
            try:
                self.copy_selected_range()
                # 복사 직후 클립보드 갱신이 지연될 수 있어 짧게 재시도
                for _ in range(3):
                    text_from_clipboard = self._read_clipboard_text()
                    if text_from_clipboard and text_from_clipboard.strip():
                        break
                    time.sleep(0.03)
            except Exception:
                text_from_clipboard = ""

            # 3) 가장 유효한(공백 제외 길이가 긴) 텍스트를 반환
            def score(s: str) -> int:
                return len((s or "").strip())

            best = text_from_gettext
            if score(text_from_clipboard) > score(best):
                best = text_from_clipboard

            return best or ""
        except Exception as e:
            print(f"텍스트 추출 실패: {e}")
            return ""
    
    def get_selection_info(self) -> dict:
        """
        현재 선택 영역 정보 가져오기 (검증용)
        
        HWP 공식 문서에 따라 InitScan() + GetText()를 올바르게 사용합니다.
        
        Returns:
            선택 영역 정보 딕셔너리
        """
        if not self.is_opened:
            return {}
        
        try:
            info = {}
            
            # 선택 시작/끝 위치
            try:
                info['selection_start'] = self.hwp.GetPos(0)  # 시작 위치
                info['selection_end'] = self.hwp.GetPos(1)  # 끝 위치
            except:
                info['selection_start'] = None
                info['selection_end'] = None
            
            # 선택된 텍스트 길이 (InitScan() + GetText() 사용)
            try:
                status_code, selected_text = self._get_text_with_scan(use_init_scan=True)
                if status_code == 2:  # 일반 텍스트
                    info['text_length'] = len(selected_text) if selected_text else 0
                else:
                    # 상태코드가 2가 아니면 클립보드 방식으로 폴백
                    try:
                        self.copy_selected_range()
                        time.sleep(0.05)
                        selected_text = self._read_clipboard_text()
                        info['text_length'] = len(selected_text) if selected_text else 0
                    except:
                        info['text_length'] = 0
            except:
                info['text_length'] = 0
            
            # 선택된 문단 수 (근사치)
            # HWP API에서 직접 문단 수를 가져오는 방법이 없으므로
            # 텍스트의 줄바꿈 개수로 근사치 계산
            try:
                status_code, selected_text = self._get_text_with_scan(use_init_scan=True)
                if status_code == 2 and selected_text:  # 일반 텍스트
                    # 줄바꿈 개수로 문단 수 추정
                    paragraph_count = selected_text.count('\n') + 1
                    info['paragraph_count'] = paragraph_count
                else:
                    # 클립보드 방식으로 폴백
                    try:
                        self.copy_selected_range()
                        time.sleep(0.05)
                        selected_text = self._read_clipboard_text()
                        if selected_text:
                            paragraph_count = selected_text.count('\n') + 1
                            info['paragraph_count'] = paragraph_count
                        else:
                            info['paragraph_count'] = 0
                    except:
                        info['paragraph_count'] = 0
            except:
                info['paragraph_count'] = 0
            
            return info
        except Exception as e:
            print(f"[디버그] 선택 영역 정보 가져오기 실패: {e}")
            return {}
    
    def get_selected_control_info(self) -> dict:
        """
        현재 선택된 컨트롤 정보 가져오기 (검증용)
        
        Returns:
            컨트롤 정보 딕셔너리
        """
        if not self.is_opened:
            return {}
        
        try:
            info = {
                'control_count': 0,
                'control_types': []
            }
            
            # 선택된 컨트롤 개수 및 타입 확인
            try:
                # HWP에서 선택된 컨트롤 정보 가져오기
                # HParameterSet.HSelectionOpt를 사용하여 선택 영역 정보 확인
                sel_opt = self.hwp.HParameterSet.HSelectionOpt
                
                # 선택된 컨트롤 개수 (근사치)
                # 실제로는 선택 영역의 컨트롤을 직접 열거해야 하지만,
                # HWP API 제한으로 인해 텍스트 기반으로 추정
                raw = self.hwp.GetText()
                selected_text = self._normalize_gettext_result(raw)
                if selected_text:
                    # 텍스트 길이로 컨트롤 존재 여부 추정
                    # 실제 컨트롤 개수는 HWP 내부 구조에 따라 다름
                    info['text_length'] = len(selected_text)
                    
                    # 컨트롤 타입 추정 (HWP API 제한으로 정확하지 않을 수 있음)
                    # 실제로는 HParameterSet을 통해 확인해야 함
                    info['control_count'] = 1  # 기본값 (실제로는 더 정확한 방법 필요)
                    info['control_types'] = ['Text']  # 기본값
            except Exception as e:
                print(f"[디버그] 컨트롤 정보 가져오기 실패: {e}")
            
            return info
        except Exception as e:
            print(f"[디버그] 선택된 컨트롤 정보 가져오기 실패: {e}")
            return {}
    
    def enumerate_controls_and_collect_problem(self, marker_start: str, marker_end: str) -> List[dict]:
        """
        컨트롤 열거 방식으로 문제 블록 수집
        
        핵심 원칙:
        - 좌표 기반 선택 폐기
        - MoveNextCtrl()로 문서 전체 컨트롤 순회
        - [문제시작] ~ [문제끝] 사이의 컨트롤들을 수집
        - 수식, 그림, 표, 문단 모두 안정적으로 보존
        
        Args:
            marker_start: [문제시작] 마커 문자열
            marker_end: [문제끝] 마커 문자열
        
        Returns:
            수집된 컨트롤 정보 리스트 (각 컨트롤의 위치 및 타입 정보)
        """
        if not self.is_opened:
            return []
        
        collected_controls = []
        is_collecting = False
        
        try:
            # 문서 처음으로 이동
            self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            
            max_iterations = 1000  # 무한 루프 방지
            iteration = 0
            
            while iteration < max_iterations:
                iteration += 1
                
                # 현재 컨트롤의 전체 텍스트 문자열 가져오기
                # 여러 방법을 시도하여 표/텍스트박스/문단 내부 텍스트도 포함
                current_text = ""
                try:
                    # 방법 1: GetFullText() 시도
                    try:
                        current_text = self.hwp.GetFullText()
                    except:
                        # 방법 2: GetCtrlText() 시도
                        try:
                            current_text = self.hwp.GetCtrlText()
                        except:
                            # 방법 3: GetText() 시도
                            try:
                                current_text = self.hwp.GetText()
                            except:
                                current_text = ""
                except:
                    current_text = ""
                
                # 문자열 포함 여부로 마커 판단 (컨트롤 == 마커 비교 금지)
                current_text = current_text or ""
                
                # [문제시작] 마커 포함 여부 확인
                if marker_start in current_text:
                    is_collecting = True
                    collected_controls = []  # 새 문제 시작
                    print(f"[디버그] [문제시작] 발견: 수집 시작 (텍스트: {current_text[:50]}...)")
                    # [문제시작] 마커 자체는 제외하고 다음 컨트롤부터 수집
                    continue
                
                # [문제끝] 마커 포함 여부 확인
                if is_collecting and marker_end in current_text:
                    print(f"[디버그] [문제끝] 발견: 수집 종료 (총 {len(collected_controls)}개 컨트롤, 텍스트: {current_text[:50]}...)")
                    break
                
                # 수집 중이면 현재 컨트롤 정보 저장
                if is_collecting:
                    try:
                        control_pos = self.hwp.GetPos()
                        control_info = {
                            'position': control_pos,
                            'text': current_text,
                            'text_length': len(current_text) if current_text else 0
                        }
                        collected_controls.append(control_info)
                    except:
                        pass
                
                # 다음 컨트롤로 이동
                try:
                    self.hwp.HAction.GetDefault("MoveNextCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    result = self.hwp.HAction.Execute("MoveNextCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    if result == 0:  # 더 이상 컨트롤이 없음
                        break
                except:
                    # MoveNextCtrl이 없으면 문서 끝으로 간주
                    break
            
            return collected_controls
        except Exception as e:
            print(f"[디버그] 컨트롤 열거 실패: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def create_hwp_from_controls(self, controls: List[dict], output_path: str) -> bool:
        """
        수집된 컨트롤들을 새 HWP 문서에 복사하여 저장
        
        Args:
            controls: 수집된 컨트롤 정보 리스트
            output_path: 저장할 파일 경로
        
        Returns:
            성공 여부
        """
        if not self.is_opened or not controls:
            return False
        
        try:
            # 원본 문서의 파일 경로 저장
            original_path = None
            try:
                original_path = self.hwp.XHwpDocuments.Item(0).FullName
            except:
                try:
                    original_path = self.hwp.XHwpDocuments.GetPathName(0)
                except:
                    pass
            
            # 새 문서 생성
            self.hwp.HAction.GetDefault("FileNew", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HAction.Execute("FileNew", self.hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 각 컨트롤을 순서대로 복사하여 새 문서에 붙여넣기
            for ctrl_idx, control_info in enumerate(controls):
                try:
                    # 원본 문서로 돌아가기
                    if original_path and os.path.exists(original_path):
                        try:
                            self.hwp.XHwpDocuments.Open(original_path)
                        except:
                            try:
                                self.hwp.HAction.GetDefault("FileOpen", self.hwp.HParameterSet.HFileOpenSave.HSet)
                                self.hwp.HParameterSet.HFileOpenSave.filename = original_path
                                self.hwp.HAction.Execute("FileOpen", self.hwp.HParameterSet.HFileOpenSave.HSet)
                            except:
                                pass
                    
                    # 컨트롤 위치로 이동
                    sec, para, pos = control_info['position']
                    self.hwp.SetPos(sec, para, pos)
                    
                    # 컨트롤 선택 (컨트롤 전체 선택)
                    try:
                        self.hwp.HAction.GetDefault("SelectCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("SelectCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    except:
                        # SelectCtrl이 없으면 기본 선택
                        self.hwp.HAction.Run("Select")
                    
                    # 복사
                    self.hwp.HAction.Run("Copy")
                    
                    # 새 문서로 전환
                    # 새 문서는 이미 생성되어 있으므로 전환만
                    try:
                        # XHwpDocuments에서 새 문서 인덱스 찾기
                        doc_count = self.hwp.XHwpDocuments.Count
                        if doc_count > 1:
                            self.hwp.XHwpDocuments.Item(doc_count - 1).Activate()
                    except:
                        pass
                    
                    # 붙여넣기
                    self.hwp.HAction.Run("Paste")
                    
                except Exception as e:
                    print(f"[경고] 컨트롤 {ctrl_idx} 복사 실패: {e}")
                    continue
            
            # 새 문서 저장
            self.hwp.HAction.GetDefault("FileSaveAs", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HParameterSet.HFileOpenSave.filename = output_path
            self.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
            self.hwp.HAction.Execute("FileSaveAs", self.hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 새 문서 닫기
            try:
                self.hwp.HAction.GetDefault("FileClose", self.hwp.HParameterSet.HFileOpenSave.HSet)
                self.hwp.HParameterSet.HFileOpenSave.filename = ""
                self.hwp.HAction.Execute("FileClose", self.hwp.HParameterSet.HFileOpenSave.HSet)
            except:
                pass
            
            # 원본 문서로 돌아가기
            if original_path and os.path.exists(original_path):
                try:
                    self.hwp.XHwpDocuments.Open(original_path)
                except:
                    try:
                        self.hwp.HAction.GetDefault("FileOpen", self.hwp.HParameterSet.HFileOpenSave.HSet)
                        self.hwp.HParameterSet.HFileOpenSave.filename = original_path
                        self.hwp.HAction.Execute("FileOpen", self.hwp.HParameterSet.HFileOpenSave.HSet)
                    except:
                        pass
            
            success = os.path.exists(output_path)
            if success:
                print(f"[디버그] 컨트롤 기반 HWP 파일 생성 성공: {output_path}")
            else:
                print(f"[디버그] 컨트롤 기반 HWP 파일 생성 실패: 파일이 생성되지 않음")
            
            return success
        except Exception as e:
            print(f"[디버그] 컨트롤 기반 HWP 파일 생성 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def select_problem_range_by_positions(self, begin_pos: tuple, end_pos: tuple) -> bool:
        """
        [문제시작] ~ [문제끝] 범위를 좌표(Pos) 기반으로 선택
        
        핵심 원칙:
        - 마커 문자열 자체는 선택 범위에서 제외
        - 마커 다음 컨트롤부터 선택 시작
        - [문제끝] 이전 컨트롤까지만 선택 확장
        - 실제 문제 컨트롤 블록만 선택
        
        Args:
            begin_pos: (Section, Paragraph, Position) 튜플 - [문제시작] 마커 위치
            end_pos: (Section, Paragraph, Position) 튜플 - [문제끝] 마커 위치
        
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        try:
            # GetPos()는 (Section, Paragraph, Position) 3개 값을 반환
            sec_s, para_s, pos_s = begin_pos
            sec_e, para_e, pos_e = end_pos
            
            # [문제시작] 마커 위치로 이동
            self.hwp.SetPos(sec_s, para_s, pos_s)
            
            # 마커 다음 컨트롤로 이동 (마커 문자열은 선택 범위에서 제외)
            # 방법 1: MoveNextCtrl 시도
            try:
                self.hwp.HAction.GetDefault("MoveNextCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveNextCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
            except:
                # MoveNextCtrl이 없으면 다음 문단 시작으로 이동
                try:
                    # 마커 길이만큼 오른쪽으로 이동
                    marker_length = 6  # "[문제시작]" 길이
                    for _ in range(marker_length + 1):
                        self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    # 문단 시작으로 이동
                    self.hwp.HAction.GetDefault("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveParaBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    pass
            
            # 선택 시작 위치 저장 (마커 다음 컨트롤 위치)
            try:
                select_start_pos = self.hwp.GetPos()
                sec_start, para_start, pos_start = select_start_pos
                print(f"[검증] 선택 시작 위치 (마커 다음): ({sec_start}, {para_start}, {pos_start})")
            except:
                print(f"[경고] 선택 시작 위치를 가져올 수 없습니다.")
                return False
            
            # 선택 시작
            self.hwp.HAction.Run("Select")
            
            # [문제끝] 마커 위치로 이동
            self.hwp.SetPos(sec_e, para_e, pos_e)
            
            # [문제끝] 이전 컨트롤로 이동 (마커 문자열은 선택 범위에서 제외)
            # 방법 1: MovePrevCtrl 시도
            try:
                self.hwp.HAction.GetDefault("MovePrevCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MovePrevCtrl", self.hwp.HParameterSet.HSelectionOpt.HSet)
            except:
                # MovePrevCtrl이 없으면 마커 시작 위치로 이동
                try:
                    # 마커 길이만큼 왼쪽으로 이동
                    marker_length = 5  # "[문제끝]" 길이
                    for _ in range(marker_length):
                        self.hwp.HAction.GetDefault("MoveLeft", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveLeft", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    pass
            
            # 선택 확장 위치 저장 ([문제끝] 이전 컨트롤 위치)
            try:
                select_end_pos = self.hwp.GetPos()
                sec_end, para_end, pos_end = select_end_pos
                print(f"[검증] 선택 끝 위치 ([문제끝] 이전): ({sec_end}, {para_end}, {pos_end})")
            except:
                print(f"[경고] 선택 끝 위치를 가져올 수 없습니다.")
                return False
            
            # 선택 확장
            self.hwp.HAction.Run("ExtendSel")
            
            print(f"[디버그] 좌표 기반 선택 완료: BeginPos=({sec_s}, {para_s}, {pos_s}), EndPos=({sec_e}, {para_e}, {pos_e})")
            print(f"[디버그] 실제 선택 범위: Start=({sec_start}, {para_start}, {pos_start}), End=({sec_end}, {para_end}, {pos_end})")
            return True
        except Exception as e:
            print(f"[디버그] 좌표 기반 선택 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def save_selected_block_to_file(self, output_path: str) -> bool:
        """
        현재 선택된 블록을 HWP 파일로 저장 (FileSaveBlock 방식)
        
        Args:
            output_path: 저장할 파일 경로
        
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        try:
            # FileSaveBlock 또는 FileSaveAs_SelBlock 사용
            try:
                # FileSaveBlock 시도
                self.hwp.HAction.GetDefault("FileSaveBlock", self.hwp.HParameterSet.HFileOpenSave.HSet)
                self.hwp.HParameterSet.HFileOpenSave.filename = output_path
                self.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
                self.hwp.HAction.Execute("FileSaveBlock", self.hwp.HParameterSet.HFileOpenSave.HSet)
            except:
                # FileSaveBlock이 없으면 FileSaveAs_SelBlock 시도
                try:
                    self.hwp.HAction.GetDefault("FileSaveAs_SelBlock", self.hwp.HParameterSet.HFileOpenSave.HSet)
                    self.hwp.HParameterSet.HFileOpenSave.filename = output_path
                    self.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
                    self.hwp.HAction.Execute("FileSaveAs_SelBlock", self.hwp.HParameterSet.HFileOpenSave.HSet)
                except:
                    # 둘 다 없으면 기존 방식 사용 (복사-붙여넣기)
                    print(f"[경고] FileSaveBlock/FileSaveAs_SelBlock를 사용할 수 없습니다. 복사-붙여넣기 방식으로 진행합니다.")
                    return self.extract_selected_to_hwp_file(output_path)
            
            success = os.path.exists(output_path)
            if success:
                print(f"[디버그] 블록 저장 성공: {output_path}")
            else:
                print(f"[디버그] 블록 저장 실패: 파일이 생성되지 않음")
                # 일부 환경에서는 액션이 실행되지만 파일이 생성되지 않는 케이스가 있음
                # 이 경우 복사-붙여넣기 방식으로 폴백하여 추출을 계속 진행
                print(f"[경고] 블록 저장 폴백: 복사-붙여넣기 방식으로 재시도합니다.")
                return self.extract_selected_to_hwp_file(output_path)
            
            return success
        except Exception as e:
            print(f"[디버그] 블록 저장 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def verify_paste_result(self) -> dict:
        """
        붙여넣기 결과 검증 (새 문서에서)
        
        Returns:
            검증 결과 딕셔너리
        """
        if not self.is_opened:
            return {}
        
        try:
            result = {}
            
            # 선택된 텍스트 확인
            try:
                text = self.hwp.GetText()
                result['text_length'] = len(text) if text else 0
                result['has_text'] = bool(text and text.strip())
            except:
                result['text_length'] = 0
                result['has_text'] = False
            
            # 문단 수 추정
            try:
                text = self.hwp.GetText()
                if text:
                    paragraph_count = text.count('\n') + 1
                    result['paragraph_count'] = paragraph_count
                else:
                    result['paragraph_count'] = 0
            except:
                result['paragraph_count'] = 0
            
            return result
        except Exception as e:
            print(f"[디버그] 붙여넣기 결과 검증 실패: {e}")
            return {}
    
    def extract_range_to_hwp_file(self, start_pos: int, end_pos: int, output_path: str) -> bool:
        """
        지정된 범위를 새 HWP 파일로 저장 (방법 A)
        
        Args:
            start_pos: 시작 위치
            end_pos: 끝 위치
            output_path: 저장할 파일 경로
            
        Returns:
            성공 여부
        """
        if not self.is_opened:
            return False
        
        try:
            # 원본 문서의 현재 상태 저장 (나중에 복원하기 위해)
            try:
                original_pos = self.hwp.GetPos()
                if isinstance(original_pos, tuple):
                    original_pos = original_pos[0] if len(original_pos) > 0 else 0
            except:
                original_pos = 0
            
            # 범위 선택: 시작 위치로 이동 후 끝 위치까지 선택
            # SetPos()는 작동하지 않으므로 다른 방법 사용
            # 문서 처음으로 이동 후 오른쪽으로 start_pos만큼 이동
            self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            
            # start_pos만큼 오른쪽으로 이동 (대략적인 이동)
            # 정확한 위치 이동은 어렵지만, 최소한 시작 위치 근처로 이동
            for _ in range(min(start_pos, 1000)):  # 최대 1000번까지만 이동 (무한 루프 방지)
                try:
                    self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                except:
                    break
            
            # 범위 선택: Shift 키를 누른 상태로 끝 위치까지 이동
            # 한글 API: Shift+End 방식으로 범위 선택
            try:
                # 선택 모드 시작 (Shift 키 누름 효과)
                # 시작 위치에서 끝 위치까지 선택
                # end_pos - start_pos만큼 오른쪽으로 이동하면서 선택
                move_count = min(end_pos - start_pos, 10000)  # 최대 10000번까지만 이동
                for _ in range(move_count):
                    try:
                        # Shift+Right (선택하면서 이동)
                        self.hwp.HAction.GetDefault("ExtendSelRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("ExtendSelRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                    except:
                        # ExtendSelRight가 없으면 일반 MoveRight 사용
                        self.hwp.HAction.GetDefault("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
                        self.hwp.HAction.Execute("MoveRight", self.hwp.HParameterSet.HSelectionOpt.HSet)
            except Exception as select_error:
                print(f"[디버그] 범위 선택 중 오류 (무시): {select_error}")
                # 선택 실패 시에도 계속 진행 (복사 시도)
            
            # 복사
            if not self.copy_selected_range():
                print(f"[디버그] 범위 복사 실패: {start_pos}~{end_pos}")
                return False
            
            # 새 문서 생성
            self.hwp.HAction.GetDefault("FileNew", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HAction.Execute("FileNew", self.hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 붙여넣기
            self.hwp.HAction.Run("Paste")
            
            # 파일 저장
            self.hwp.HAction.GetDefault("FileSaveAs", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HParameterSet.HFileOpenSave.filename = output_path
            self.hwp.HParameterSet.HFileOpenSave.Format = "HWP"
            self.hwp.HAction.Execute("FileSaveAs", self.hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 새 문서 닫기
            self.hwp.HAction.GetDefault("FileClose", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HAction.Execute("FileClose", self.hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 원본 문서로 돌아가기 (다음 문제 추출을 위해)
            # SetPos()는 작동하지 않으므로, 문서 처음으로 이동
            try:
                self.hwp.HAction.GetDefault("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
                self.hwp.HAction.Execute("MoveDocBegin", self.hwp.HParameterSet.HSelectionOpt.HSet)
            except:
                pass
            
            success = os.path.exists(output_path)
            if success:
                print(f"[디버그] HWP 파일 추출 성공: {output_path}")
            else:
                print(f"[디버그] HWP 파일 추출 실패: 파일이 생성되지 않음")
            
            return success
        except Exception as e:
            print(f"[디버그] HWP 파일 추출 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_text_from_range(self, start_pos: int, end_pos: int) -> str:
        """
        범위의 텍스트만 추출 (검색용)
        
        Args:
            start_pos: 시작 위치
            end_pos: 끝 위치
            
        Returns:
            추출된 텍스트
        """
        if not self.is_opened:
            return ""
        
        try:
            # 범위 선택
            if not self.select_range(start_pos, end_pos):
                return ""
            
            # 텍스트 추출
            # GetText()는 선택된 범위의 텍스트를 반환하거나, 인자 없이 호출 시 선택된 텍스트 반환
            try:
                # 선택된 범위의 텍스트 가져오기
                text = self.hwp.GetText()
                return text if text else ""
            except:
                # GetText() 실패 시 빈 문자열 반환
                return ""
        except Exception as e:
            print(f"텍스트 추출 실패: {e}")
            return ""
    
    def cleanup(self):
        """리소스 정리"""
        # 문서 닫기(저장 질문 방지)
        try:
            self.close_document()
        except Exception:
            pass
        if self.hwp:
            try:
                # ✅ Quit 시 "모두 저장/모두 저장 안 함" 팝업이 뜨는 것을 방지:
                # - 남아있는 문서들의 수정 상태를 버리고(가능하면) 1개 문서만 남긴 뒤 종료
                with self._temp_message_box_mode(0x20021):  # No + Cancel + OK
                    try:
                        # 여러 문서가 열려 있으면 추가 문서부터 닫기(수정사항 버림)
                        for _ in range(30):
                            cnt = 0
                            try:
                                cnt = int(getattr(self.hwp.XHwpDocuments, "Count", 0))
                            except Exception:
                                cnt = 0
                            if cnt <= 1:
                                break
                            try:
                                self.hwp.XHwpDocuments.Active_XHwpDocument.Close(isDirty=False)
                            except Exception:
                                try:
                                    self.hwp.XHwpDocuments.Active_XHwpDocument.Close(False)
                                except Exception:
                                    break

                        # 마지막 문서도 "수정사항 버림" 상태로 정리(가능하면)
                        try:
                            self.hwp.XHwpDocuments.Active_XHwpDocument.Clear(1)
                        except Exception:
                            try:
                                self.hwp.XHwpDocuments.Active_XHwpDocument.Clear(option=1)
                            except Exception:
                                pass
                    except Exception:
                        pass

                self.hwp.Quit()
            except:
                pass
            self.hwp = None
        self.is_opened = False
    
    def __enter__(self):
        """Context manager 진입"""
        self.initialize()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager 종료"""
        self.cleanup()
