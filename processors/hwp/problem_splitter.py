"""
문제 분리 모듈

HWP 문서에서 [문제시작] ~ [문제끝] 마커를 기준으로 문제를 분리합니다.
- 마커 검색 및 위치 파악
- 각 문제 영역 추출 (방법 A: 범위 선택 → 복사 → 새 HWP 저장)
- 문제별 메타데이터 생성
- 분리된 문제 리스트 반환

미주 기반 파싱은 pyhwpx를 사용합니다 (endnote_save.py 방식).
"""
import os
import tempfile
import time
from typing import List, Tuple, Optional, Any
import json
try:
    import win32clipboard
    CLIPBOARD_AVAILABLE = True
except:
    CLIPBOARD_AVAILABLE = False

try:
    import win32com.client
    import win32gui
    import win32con
    import win32api
    WIN32_AVAILABLE = True
except:
    WIN32_AVAILABLE = False

try:
    from pyhwpx import Hwp
    PYHWPX_AVAILABLE = True
except ImportError:
    PYHWPX_AVAILABLE = False
    print("[경고] pyhwpx가 설치되지 않았습니다. pip install pyhwpx를 실행하세요.")


class ProblemSplitter:
    """문제 분리 클래스"""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        ProblemSplitter 초기화
        
        Args:
            config_path: config.json 파일 경로
        """
        if config_path is None:
            config_path = os.path.join(
                os.path.dirname(os.path.dirname(os.path.dirname(__file__))),
                'config', 'config.json'
            )
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        hwp_config = config.get('hwp', {})
        self.marker_start = hwp_config.get('marker_start', '[문제시작]')
        self.marker_end = hwp_config.get('marker_end', '[문제끝]')
        # DB 생성 속도 최우선: 파싱 중 텍스트 추출을 끄면(원본 HWP 블록 저장만) 훨씬 빨라집니다.
        self.extract_text_during_parse = bool(hwp_config.get('extract_text_during_parse', False))
    
    def find_problem_ranges(self, hwp_reader: HWPReader) -> List[Tuple[int, int]]:
        """
        [문제시작] ~ [문제끝] 마커 위치 찾기
        
        핵심 원칙:
        - 현재 커서 이후에서만 검색
        - 문서 전체 기준 검색 금지
        - 중복 체크는 보조 안전장치로만 사용
        
        Args:
            hwp_reader: HWPReader 인스턴스
            
        Returns:
            [(시작_위치, 끝_위치), ...] 리스트
        """
        ranges = []
        
        # 문서 처음으로 이동 (한 번만)
        hwp_reader.hwp.HAction.GetDefault("MoveDocBegin", hwp_reader.hwp.HParameterSet.HSelectionOpt.HSet)
        hwp_reader.hwp.HAction.Execute("MoveDocBegin", hwp_reader.hwp.HParameterSet.HSelectionOpt.HSet)
        
        print(f"[디버그] 마커 검색 시작: 시작='{self.marker_start}', 끝='{self.marker_end}'")
        
        # 이전에 찾은 [문제시작] 텍스트를 저장하여 중복 체크 (보조 안전장치)
        previous_start_text = None
        
        # while True 루프 구조로 문서 끝까지 반복 탐색
        while True:
            # [문제시작] 찾기 (현재 커서 이후에서만 검색)
            start_result = hwp_reader.find_text(self.marker_start, start_from_beginning=False)
            if start_result is None:
                print(f"[디버그] [문제시작] 마커를 더 이상 찾을 수 없습니다.")
                break  # 더 이상 찾을 수 없음
            
            # 중복 체크: 보조 안전장치 (동일 위치에서 다시 찾으면 중단)
            if CLIPBOARD_AVAILABLE:
                try:
                    # 현재 선택된 [문제시작] 텍스트를 클립보드에 복사
                    hwp_reader.hwp.HAction.Run("Copy")
                    win32clipboard.OpenClipboard()
                    try:
                        current_start_text = win32clipboard.GetClipboardData()
                    except:
                        current_start_text = None
                    finally:
                        win32clipboard.CloseClipboard()
                    
                    # 이전과 동일한 텍스트를 찾았으면 같은 위치를 반복 찾는 것 (안전장치)
                    if current_start_text == previous_start_text and previous_start_text is not None:
                        print(f"[디버그] 중복 발견: 같은 위치에서 [문제시작]을 다시 찾았습니다. 루프를 중단합니다.")
                        break
                    
                    previous_start_text = current_start_text
                except:
                    pass  # 클립보드 확인 실패 시 계속 진행
            
            # start_pos 저장 (현재 커서 위치 기준)
            start_pos, _ = start_result
            
            # [문제끝] 찾기 (start_pos 이후, 현재 커서 위치에서)
            end_result = hwp_reader.find_text(self.marker_end, start_from_beginning=False)
            if end_result is None:
                print(f"[디버그] [문제끝] 마커를 찾을 수 없습니다.")
                # 파싱 오류: 문제시작은 찾았지만 문제끝을 찾을 수 없음
                raise ValueError(f"[문제시작]을 찾았지만 [문제끝]을 찾을 수 없습니다.")
            
            # end_pos 저장
            _, end_pos = end_result
            ranges.append((start_pos, end_pos))
            print(f"[디버그] 문제 범위 추가: {start_pos}~{end_pos} (총 {len(ranges)}개)")
            
            # end_pos + 1 로 커서 이동 (다음 문제를 찾기 위해)
            # find_text에서 이미 커서가 이동했지만, 확실히 하기 위해 추가 이동
            try:
                # 선택 해제
                hwp_reader.hwp.HAction.Run("Cancel")
                # [문제끝] 마커 길이 + 1 만큼 오른쪽으로 이동
                end_marker_length = len(self.marker_end)
                for _ in range(end_marker_length + 1):
                    hwp_reader.hwp.HAction.GetDefault("MoveRight", hwp_reader.hwp.HParameterSet.HSelectionOpt.HSet)
                    hwp_reader.hwp.HAction.Execute("MoveRight", hwp_reader.hwp.HParameterSet.HSelectionOpt.HSet)
            except:
                pass  # 이동 실패 시에도 계속 진행
        
        print(f"[디버그] 총 {len(ranges)}개의 문제 범위를 찾았습니다.")
        return ranges
    
    def extract_problem_block(
        self, 
        hwp_reader: HWPReader, 
        start_pos: int, 
        end_pos: int,
        temp_dir: Optional[str] = None
    ) -> Optional[bytes]:
        """
        지정된 범위의 HWP 원본 블록 추출 (방법 A)
        
        Args:
            hwp_reader: HWPReader 인스턴스
            start_pos: 시작 위치
            end_pos: 끝 위치
            temp_dir: 임시 파일 저장 디렉토리
            
        Returns:
            HWP 바이너리 데이터 또는 None
        """
        if temp_dir is None:
            temp_dir = tempfile.gettempdir()
        
        # 임시 파일 경로 생성
        temp_file = os.path.join(temp_dir, f"problem_{os.urandom(8).hex()}.hwp")
        
        try:
            # 범위를 새 HWP 파일로 저장
            if not hwp_reader.extract_range_to_hwp_file(start_pos, end_pos, temp_file):
                return None
            
            # 파일을 바이너리로 읽기
            if os.path.exists(temp_file):
                with open(temp_file, 'rb') as f:
                    hwp_bytes = f.read()
                
                # 임시 파일 삭제
                try:
                    os.remove(temp_file)
                except:
                    pass
                
                return hwp_bytes
            else:
                return None
        except Exception as e:
            print(f"문제 블록 추출 실패: {e}")
            # 임시 파일 정리
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except:
                    pass
            return None
    
    def extract_text_for_search(
        self, 
        hwp_reader: HWPReader, 
        start_pos: int, 
        end_pos: int
    ) -> str:
        """
        검색용 텍스트 추출 (보조 필드)
        
        Args:
            hwp_reader: HWPReader 인스턴스
            start_pos: 시작 위치
            end_pos: 끝 위치
            
        Returns:
            추출된 텍스트
        """
        return hwp_reader.get_text_from_range(start_pos, end_pos)
    
    def split_problems(
        self, 
        hwp_path: str,
        temp_dir: Optional[str] = None
    ) -> List[Tuple[bytes, str]]:
        """
        HWP 파일에서 모든 문제 블록 추출
        
        핵심 원칙:
        1. 현재 커서 이후에서만 검색 (문서 전체 기준 검색 금지)
        2. HWP 블록 자체를 복사 (수식, 그림, 표 포함)
        3. 텍스트 추출이 아닌 HWP 객체 복사
        4. 복사-붙여넣기 방식 사용 (Ctrl+C → Ctrl+V와 동일)
        
        Args:
            hwp_path: HWP 파일 경로
            temp_dir: 임시 파일 저장 디렉토리
            
        Returns:
            [(hwp_bytes, text), ...] 리스트
            - hwp_bytes: HWP 원본 바이너리 (수식/그림/표 포함)
            - text: 검색용 텍스트 (보조 필드)
        """
        results = []
        
        with HWPReader() as hwp_reader:
            # HWP 파일 열기
            if not hwp_reader.open_document(hwp_path):
                return results
            
            if temp_dir is None:
                temp_dir = tempfile.gettempdir()
            
            # 문서 처음으로 이동 (한 번만)
            hwp_reader.hwp.HAction.GetDefault("MoveDocBegin", hwp_reader.hwp.HParameterSet.HSelectionOpt.HSet)
            hwp_reader.hwp.HAction.Execute("MoveDocBegin", hwp_reader.hwp.HParameterSet.HSelectionOpt.HSet)
            
            print(f"[디버그] 마커 검색 시작: 시작='{self.marker_start}', 끝='{self.marker_end}'")
            
            # 이전에 찾은 [문제시작] 텍스트를 저장하여 중복 체크 (보조 안전장치)
            previous_start_text = None
            problem_index = 0
            last_advance_pos = None
            seen_problem_starts = set()
            
            # while True 루프: 방법 1 - 범위 선택 기반 문제 블록 추출
            while True:
                problem_index += 1
                
                # 방법 1: 범위 선택 기반 문제 블록 추출
                # 1. [문제시작] ~ [문제끝] 사이를 선택
                if not hwp_reader.select_range_between_markers(self.marker_start, self.marker_end):
                    print(f"[디버그] 문제 #{problem_index} 범위 선택 실패. 더 이상 문제를 찾을 수 없습니다. (총 {problem_index - 1}개 문제 발견)")
                    break

                # ✅ wrap-around/무한루프 방지: 같은 문제 시작 위치(문제 본문 시작)가 다시 나오면 종료
                current_start = getattr(hwp_reader, "last_problem_start_pos", None)
                if current_start is not None:
                    if current_start in seen_problem_starts:
                        print(f"[경고] 동일한 문제 시작 위치가 다시 감지되어 루프를 종료합니다. start={current_start}")
                        break
                    seen_problem_starts.add(current_start)
                
                print(f"[디버그] 문제 #{problem_index} 추출 시작...")
                
                # 2. 선택된 범위를 새 HWP 파일로 저장
                temp_file = os.path.join(temp_dir, f"problem_{os.urandom(8).hex()}.hwp")
                
                # 검색용 텍스트 추출 (선택된 상태에서, 새 문서 생성 전에)
                if self.extract_text_during_parse:
                    try:
                        text = hwp_reader.get_text_from_selection()
                    except Exception:
                        text = ""  # 텍스트 추출 실패 시 빈 문자열
                else:
                    text = ""  # 속도 우선: 파싱 중 텍스트 추출 생략
                
                # 선택된 범위를 새 HWP 파일로 저장
                # FileSaveBlock이 가능하면 그 경로가 미주/각주 포함에 더 유리합니다.
                if hwp_reader.save_selected_block_to_file(temp_file):
                    # 파일을 바이너리로 읽기
                    if os.path.exists(temp_file):
                        with open(temp_file, 'rb') as f:
                            hwp_bytes = f.read()
                        
                        results.append((hwp_bytes, text))
                        print(f"[디버그] 문제 #{problem_index} 추출 완료")
                        # ⚠️ 자동으로 파일을 열면 수백 개가 열릴 수 있어 비활성화합니다.
                        # 필요하면 problem_index == 1일 때만 열도록 바꾸세요.
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                    else:
                        print(f"[디버그] 문제 #{problem_index} 추출 실패: 파일이 생성되지 않음")
                else:
                    print(f"[디버그] 문제 #{problem_index} 추출 실패")
                
                # [문제끝] 다음 위치로 커서 이동 (다음 문제를 찾기 위해)
                try:
                    # ✅ (개선) [문제끝]을 "다시 찾기" 하면 문서 끝에서
                    # "끝까지 찾았습니다/더 이상 없음" 팝업이 발생할 수 있습니다.
                    # 선택을 해제하면 커서는 선택 끝(= [문제끝] 직전)에 위치하므로,
                    # 마커 길이만큼 오른쪽으로 이동해 "마커 뒤"로 진행합니다.
                    hwp_reader.hwp.HAction.Run("Cancel")
                except:
                    pass

                try:
                    end_marker_length = len(self.marker_end)
                    for _ in range(end_marker_length + 1):
                        hwp_reader.hwp.HAction.GetDefault("MoveRight", hwp_reader.hwp.HParameterSet.HSelectionOpt.HSet)
                        hwp_reader.hwp.HAction.Execute("MoveRight", hwp_reader.hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception as e:
                    print(f"[경고] 다음 문제로 커서 이동 실패: {e}")
                    break

                # 무한루프 안전장치: [문제끝] 뒤로 이동한 커서가 이전과 동일하면 중단
                try:
                    new_pos = hwp_reader.hwp.GetPos()
                except Exception:
                    new_pos = None
                if new_pos is not None and new_pos == last_advance_pos:
                    print(f"[경고] 커서가 [문제끝] 뒤로 진행되지 않아 루프를 중단합니다. pos={new_pos}")
                    break
                last_advance_pos = new_pos
        
        print(f"[디버그] 총 {len(results)}개의 문제 블록을 추출했습니다.")
        return results

    def split_problems_by_endnote(
        self, 
        hwp_path: str,
        temp_dir: Optional[str] = None,
        *,
        apply_style_to_blocks: bool = False
    ) -> List[Tuple[bytes, str]]:
        """
        HWP 파일에서 미주 기반으로 모든 문제 블록 추출 (pyhwpx 방식, endnote_save.py 기반)
        
        파싱 로직 (endnote_save.py 기반):
        1. HeadCtrl로 모든 미주 앵커 위치 수집 (ena)
        2. 각 미주의 끝점 계산: 다음 미주 직전 또는 "노블록" (ene)
        3. 미주 앵커 ~ 끝점 범위 선택 및 저장
        4. 뒤에서부터 처리하여 문서 구조 유지
        
        파싱 시에는 apply_style_to_blocks=False로 호출하여 스타일 코드를 전혀 타지 않음(잘림 방지).
        저장된 한 문제짜리 문서에 스타일을 넣고 싶을 때만 apply_style_to_blocks=True 사용.
        
        Args:
            hwp_path: HWP 파일 경로
            temp_dir: 임시 파일 저장 디렉토리
            apply_style_to_blocks: True면 각 블록 저장 시 Paste 직후 스타일 적용. 기본 False(파싱만).
            
        Returns:
            [(hwp_bytes, text), ...] 리스트
            - hwp_bytes: HWP 원본 바이너리 (수식/그림/표 포함)
            - text: 검색용 텍스트 (보조 필드)
        """
        if not PYHWPX_AVAILABLE:
            print("[오류] pyhwpx가 설치되지 않았습니다. pip install pyhwpx를 실행하세요.")
            return []
        
        results = []
        
        if temp_dir is None:
            temp_dir = tempfile.gettempdir()
        
        try:
            # pyhwpx로 HWP 파일 열기
            hwp = Hwp()
            hwp.open(hwp_path)
            
            print(f"[디버그] 미주 기반 파싱 시작 (pyhwpx 방식)")
            
            # 0단계: 문서 정리 (단나누기/쪽나누기 제거, 1단 설정)
            self._prepare_document_pyhwpx(hwp)
            
            # 0.5단계: 본문 빈줄 제거 (본문스캔만, 미주 본문 정리는 제외)
            # 미주 앵커를 먼저 수집해야 빈줄 제거 가능
            temp_anchors = self._collect_endnote_anchors_pyhwpx(hwp)
            if temp_anchors:
                self._remove_blank_lines_pyhwpx(hwp, temp_anchors)
            
            # 1단계: 모든 미주 앵커 위치 수집 (ena 함수) - 빈줄 제거 후 다시 수집
            anchors = self._collect_endnote_anchors_pyhwpx(hwp)
            if not anchors:
                print(f"[디버그] 미주를 찾을 수 없습니다.")
                hwp.Quit()
                return results
            
            print(f"[디버그] 총 {len(anchors)}개의 미주 앵커를 발견했습니다.")
            
            # 2단계: 각 미주의 끝점 계산 (ene 함수)
            end_positions = self._calculate_endnote_end_positions_pyhwpx(hwp, anchors)
            
            if len(end_positions) != len(anchors):
                print(f"[경고] 미주 앵커 수({len(anchors)})와 끝점 수({len(end_positions)})가 일치하지 않습니다.")
                hwp.Quit()
                return results
            
            # 3단계: 뒤에서부터 문제 블록 추출 (ext 함수 방식)
            # 뒤에서부터 처리하여 문서 구조 유지
            for i in range(len(anchors) - 1, -1, -1):
                problem_index = i + 1
                anchor_pos = anchors[i]
                end_pos = end_positions[i]
                
                print(f"[디버그] 문제 #{problem_index} 처리 시작 (미주 앵커: {anchor_pos}, 끝점: {end_pos})")
                
                # 범위 선택 및 저장
                temp_file = os.path.join(temp_dir, f"problem_{os.urandom(8).hex()}.hwp")
                
                if self._select_and_save_block_pyhwpx(hwp, anchor_pos, end_pos, temp_file, apply_style=apply_style_to_blocks):
                    if os.path.exists(temp_file):
                        with open(temp_file, 'rb') as f:
                            hwp_bytes = f.read()
                        
                        # 검색용 텍스트 추출 (선택된 상태에서)
                        if self.extract_text_during_parse:
                            try:
                                # pyhwpx에서 텍스트 추출
                                text = hwp.GetText() or ""
                            except Exception:
                                text = ""
                        else:
                            text = ""
                        
                        results.append((hwp_bytes, text))
                        print(f"[디버그] 문제 #{problem_index} 추출 완료")
                        
                        try:
                            os.remove(temp_file)
                        except:
                            pass
                    else:
                        print(f"[디버그] 문제 #{problem_index} 추출 실패: 파일이 생성되지 않음")
                else:
                    print(f"[디버그] 문제 #{problem_index} 추출 실패: 범위 선택 또는 저장 실패")
            
            # 결과를 순서대로 정렬 (뒤에서부터 처리했으므로 reverse)
            results.reverse()
            
            # HWP 닫기
            hwp.Quit()
                    
        except Exception as e:
            print(f"[오류] pyhwpx 파싱 실패: {e}")
            import traceback
            traceback.print_exc()
            try:
                hwp.Quit()
            except:
                pass
        
        print(f"[디버그] 총 {len(results)}개의 문제 블록을 추출했습니다.")
        return results
            
    def _prepare_document_pyhwpx(self, hwp: Any) -> None:
        """
        문서 파싱 전 준비 작업
        1. 단을 1단으로 설정
        2. 단나누기(18), 쪽나누기(19) 제거
        
        Args:
            hwp: pyhwpx Hwp 객체
        """
        try:
            print(f"[디버그] 문서 정리 시작: 단나누기/쪽나누기 제거, 1단 설정")
            
            # 1단 설정
            try:
                hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HPageSetup.HSet)
                hwp.HParameterSet.HPageSetup.ColumnCount = 1  # 1단
                hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HPageSetup.HSet)
                print(f"[디버그] 1단 설정 완료")
            except Exception as e:
                print(f"[경고] 1단 설정 실패: {e}")
            
            # 단나누기(18), 쪽나누기(19) 제거
            try:
                hwp.HAction.GetDefault("DeleteCtrls", hwp.HParameterSet.HDeleteCtrls.HSet)
                # 2개의 컨트롤 타입 제거 (단나누기, 쪽나누기)
                hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 2)
                hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 18)  # 단나누기
                hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(1, 19)  # 쪽나누기
                hwp.HAction.Execute("DeleteCtrls", hwp.HParameterSet.HDeleteCtrls.HSet)
                print(f"[디버그] 단나누기/쪽나누기 제거 완료")
            except Exception as e:
                print(f"[경고] 단나누기/쪽나누기 제거 실패: {e}")
                # 대체 방법: 각각 따로 제거 시도
                try:
                    # 단나누기만 제거
                    hwp.HAction.GetDefault("DeleteCtrls", hwp.HParameterSet.HDeleteCtrls.HSet)
                    hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
                    hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 18)
                    hwp.HAction.Execute("DeleteCtrls", hwp.HParameterSet.HDeleteCtrls.HSet)
                except:
                    pass
                try:
                    # 쪽나누기만 제거
                    hwp.HAction.GetDefault("DeleteCtrls", hwp.HParameterSet.HDeleteCtrls.HSet)
                    hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
                    hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 19)
                    hwp.HAction.Execute("DeleteCtrls", hwp.HParameterSet.HDeleteCtrls.HSet)
                except:
                    pass
        
        except Exception as e:
            print(f"[경고] 문서 정리 실패: {e}")
            # 실패해도 계속 진행
    
    def _remove_blank_lines_pyhwpx(self, hwp: Any, anchors: List[Tuple[int, int, int]]) -> None:
        """
        빈줄 제거 (본문스캔 + cln 함수, pyhwpx 버전)
        
        Args:
            hwp: pyhwpx Hwp 객체
            anchors: 미주 앵커 위치 리스트
        """
        try:
            print(f"[디버그] 빈줄 제거 시작")
            
            # 1) 빈 줄 길이 측정 (emp 함수)
            blank_len = self._measure_blank_line_length_pyhwpx(hwp)
            
            # 2) 본문스캔: 미주 기준 본문 빈줄 정리 (본문만, 미주 본문 정리는 제외)
            self._scan_and_remove_blank_lines_pyhwpx(hwp, anchors, blank_len)
            
            print(f"[디버그] 본문 빈줄 제거 완료")
        except Exception as e:
            print(f"[경고] 빈줄 제거 실패: {e}")
            import traceback
            traceback.print_exc()
    
    def _measure_blank_line_length_pyhwpx(self, hwp: Any) -> int:
        """
        빈 줄 길이 측정 (emp 함수, pyhwpx 버전)
        
        Args:
            hwp: pyhwpx Hwp 객체
            
        Returns:
            빈 줄 길이
        """
        try:
            # 문서 끝에 공백 넣고 선택 길이 측정 후 원복
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            
            hwp.Run("MoveTopLevelEnd")
            hwp.Run("BreakPara")
            hwp.Run("BreakPara")
            
            hwp.HParameterSet.HInsertText.Text = "  "  # 공백 2개
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            
            # 현재 줄 선택
            hwp.MoveLineBegin()
            hwp.MoveSelLineEnd()
            
            # 선택 길이 측정
            s = hwp.GetTextFile("HWP", "saveblock")
            n = len(s) if s else 0
            
            # 원복 (3번 backspace)
            hwp.Run("DeleteBack")
            hwp.Run("DeleteBack")
            hwp.Run("DeleteBack")
            
            return n
        except Exception as e:
            print(f"[경고] 빈 줄 길이 측정 실패: {e}")
            return 0
    
    def _is_blank_line_pyhwpx(self, hwp: Any, blank_len: int) -> bool:
        """
        현재 문단이 빈 줄인지 판정 (isb 함수, pyhwpx 버전).
        확실할 때만 True 반환. 불확실하면 False(삭제 안 함)로 잘림 방지.
        """
        try:
            # pyhwpx의 is_empty_para 메서드가 있으면 최우선 사용
            if hasattr(hwp, "is_empty_para"):
                try:
                    if hwp.is_empty_para():
                        return True
                except Exception:
                    pass
            
            # 줄 단위 선택: Run으로 통일 (pyhwpx에서 선택이 제대로 잡히도록)
            if hasattr(hwp, "Run"):
                hwp.Run("MoveLineBegin")
                hwp.Run("MoveSelLineEnd")
            else:
                hwp.MoveLineBegin()
                hwp.MoveSelLineEnd()
            s = hwp.GetTextFile("HWP", "saveblock")
            # GetTextFile 실패/None이면 삭제하지 않음(빈줄이 아닐 수 있음)
            if s is None:
                return False
            s = s.strip()
            # 눈에 보이는 문자가 하나도 없을 때만 빈 줄로 판정
            if s == "":
                return True
            # 길이만 맞고 내용이 있으면 빈 줄 아님
            return False
        except Exception:
            return False
    
    def _scan_and_remove_blank_lines_pyhwpx(
        self,
        hwp: Any,
        anchors: List[Tuple[int, int, int]],
        blank_len: int
    ) -> None:
        """
        본문 전체 텍스트 스캔: 연속 빈줄 삭제 (본문스캔 함수, pyhwpx 버전)
        
        Args:
            hwp: pyhwpx Hwp 객체
            anchors: 미주 앵커 위치 리스트
            blank_len: 빈 줄 길이
        """
        try:
            if not anchors:
                return
            
            # '노블록'을 가상의 마지막 앵커로 추가
            nob_pos = self._find_noblock_position_pyhwpx(hwp, "노블록")
            all_anchors = anchors + [nob_pos]
            
            # 첫 번째 미주 위치 저장 (예외 처리용)
            first_anchor = anchors[0] if anchors else None
            
            # 마지막 미주부터 역순 처리
            for a in reversed(all_anchors):
                # 미주 앵커 위치로 이동
                self._set_pos_pyhwpx(hwp, a)
                
                # 이전 문단으로 이동
                try:
                    if hasattr(hwp, "move_pos"):
                        moved = hwp.move_pos(11)  # movePrevPara
                    else:
                        moved = hwp.MovePos(11, 0, 0)
                    if not moved:
                        continue
                except:
                    continue
                
                deleted_any = False
                
                # 연속된 빈 문단 삭제
                while True:
                    # 표 안이면 종료
                    if hasattr(hwp, "is_cell") and hwp.is_cell():
                        break
                    
                    if self._is_blank_line_pyhwpx(hwp, blank_len):
                        hwp.Run("DeleteBack")
                        deleted_any = True
                        continue
                    else:
                        break
                
                # 첫 번째 미주는 BreakPara() 실행하지 않음 (불필요한 엔터 추가 방지)
                if deleted_any and first_anchor and a != first_anchor:
                    hwp.Run("MoveParaEnd")
                    hwp.BreakPara()
            
            hwp.Run("MoveTopLevelBegin")
        except Exception as e:
            print(f"[경고] 본문 빈줄 스캔 실패: {e}")
    
    def _clean_endnote_bodies_pyhwpx(self, hwp: Any, blank_len: int) -> None:
        """
        모든 미주 본문으로 들어가서 앞/뒤 빈줄 제거 (cln 함수, pyhwpx 버전)
        
        Args:
            hwp: pyhwpx Hwp 객체
            blank_len: 빈 줄 길이
        """
        try:
            EN_BODY = 14  # 미주 본문으로 진입
            LST_END = 5   # 현재 리스트 끝(미주 본문 끝)
            LST_BEG = 4   # 현재 리스트 시작(미주 본문 시작)
            
            hwp.Run("MoveTopLevelBegin")
            cnt = 0
            
            c = hwp.HeadCtrl
            while c:
                try:
                    if c.CtrlID == "en":  # 미주
                        # 미주 앵커로 이동
                        anchor_posset = c.GetAnchorPos(0)
                        hwp.SetPosBySet(anchor_posset)
                        
                        # 미주 본문 진입
                        if hasattr(hwp, "move_pos"):
                            hwp.move_pos(EN_BODY)
                        else:
                            hwp.MovePos(EN_BODY, 0, 0)
                        
                        # 미주 본문 끝부분 빈줄 제거 (trb)
                        self._remove_endnote_trailing_blanks_pyhwpx(hwp, blank_len, LST_END)
                        
                        # 미주 본문 시작부분 빈줄 제거 (tlb)
                        self._remove_endnote_leading_blanks_pyhwpx(hwp, blank_len, LST_BEG)
                        
                        cnt += 1
                except Exception as e:
                    print(f"[경고] 미주 본문 빈줄 제거 실패: {e}")
                    pass
                c = c.Next
            
            hwp.Run("MoveTopLevelBegin")
            print(f"[디버그] {cnt}개 미주 본문 빈줄 제거 완료")
        except Exception as e:
            print(f"[경고] 미주 본문 빈줄 제거 실패: {e}")
    
    def _remove_endnote_trailing_blanks_pyhwpx(self, hwp: Any, blank_len: int, lst_end: int) -> None:
        """미주 본문 끝부분 빈줄 제거 (trb 함수, pyhwpx 버전)"""
        try:
            while True:
                if hasattr(hwp, "move_pos"):
                    hwp.move_pos(lst_end)
                else:
                    hwp.MovePos(lst_end, 0, 0)
                
                hwp.MoveLineBegin()
                hwp.MoveSelLineEnd()
                
                s = hwp.GetTextFile("HWP", "saveblock")
                # 불확실할 때(s None/빈문자열) 삭제하지 않음. 빈줄이 아니면 삭제 안 되게.
                if s is None or len(s) == 0:
                    break
                
                if self._is_blank_line_pyhwpx(hwp, blank_len):
                    hwp.Run("DeleteBack")
                    hwp.Run("DeleteBack")
                    continue
                
                hwp.Run("MoveLineEnd")
                ps = hwp.GetPosBySet()
                try:
                    pos_val = int(ps.Item("Pos"))
                    if pos_val == 0:
                        hwp.MoveLineBegin()
                        hwp.MoveSelLineEnd()
                        hwp.Run("DeleteBack")
                        hwp.Run("DeleteBack")
                        continue
                except:
                    pass
                
                break
        except Exception:
            pass
    
    def _remove_endnote_leading_blanks_pyhwpx(self, hwp: Any, blank_len: int, lst_beg: int) -> None:
        """미주 본문 시작부분 빈줄 제거 (tlb 함수, pyhwpx 버전)"""
        try:
            while True:
                if hasattr(hwp, "move_pos"):
                    hwp.move_pos(lst_beg)
                else:
                    hwp.MovePos(lst_beg, 0, 0)
                
                hwp.Run("MoveSelLineEnd")
                
                s = hwp.GetTextFile("HWP", "saveblock")
                n = len(s) if s else 0
                # n==0(GetTextFile 실패/빈문자열)이면 삭제하지 않음. 빈줄이 아니면 삭제 안 되게.
                if n == 0:
                    break
                if (s or "").strip() == "":
                    hwp.Run("Delete")
                    continue
                
                hwp.Run("MoveLineEnd")
                ps = hwp.GetPosBySet()
                try:
                    pos_val = int(ps.Item("Pos"))
                    if pos_val == 0:
                        hwp.MoveLineBegin()
                        hwp.MoveSelLineEnd()
                        hwp.Run("Delete")
                        hwp.Run("Delete")
                        continue
                except:
                    pass
                
                break
        except Exception:
            pass
    
    def _collect_endnote_anchors_pyhwpx(self, hwp: Any) -> List[Tuple[int, int, int]]:
        """
        모든 미주 앵커 위치 수집 (ena 함수, pyhwpx 버전)
        
        Args:
            hwp: pyhwpx Hwp 객체
            
        Returns:
            미주 앵커 위치 리스트 [(sec, para, pos), ...]
        """
        anchors = []
        
        try:
            # 문서 처음으로 이동
            hwp.Run("MoveTopLevelBegin")
            
            # HeadCtrl로 모든 컨트롤 순회
            c = hwp.HeadCtrl
            while c:
                try:
                    if c.CtrlID == "en":  # 미주 컨트롤
                        # 미주 앵커 위치로 이동
                        anchor_posset = c.GetAnchorPos(0)
                        hwp.SetPosBySet(anchor_posset)
                        anchor_pos = self._get_pos_pyhwpx(hwp)
                        if anchor_pos:
                            anchors.append(anchor_pos)
                            print(f"[디버그] 미주 앵커 #{len(anchors)} 발견: {anchor_pos}")
                except Exception as e:
                    print(f"[경고] 미주 앵커 읽기 실패: {e}")
                    pass
                c = c.Next
        except Exception as e:
            print(f"[경고] 미주 앵커 수집 실패: {e}")
        
        return anchors
    
    def _get_pos_pyhwpx(self, hwp: Any) -> Optional[Tuple[int, int, int]]:
        """pyhwpx에서 위치 가져오기 (gps 함수)"""
        try:
            if hasattr(hwp, "get_pos"):
                return tuple(hwp.get_pos())
            return tuple(hwp.GetPos())
        except Exception:
            return None
    
    def _set_pos_pyhwpx(self, hwp: Any, pos: Tuple[int, int, int]) -> None:
        """pyhwpx에서 위치 설정 (sps 함수)"""
        try:
            if hasattr(hwp, "set_pos"):
                hwp.set_pos(*pos)
            else:
                hwp.SetPos(*pos)
        except Exception as e:
            print(f"[경고] 위치 설정 실패: {e}")
    
    def _run_cmd_pyhwpx(self, hwp: Any, cmd: str) -> None:
        """pyhwpx에서 명령 실행 (run 함수)"""
        try:
            fn = getattr(hwp, cmd, None)
            if callable(fn):
                fn()
            else:
                hwp.Run(cmd)
        except Exception as e:
            print(f"[경고] 명령 실행 실패 ({cmd}): {e}")
    
    def _find_noblock_position_pyhwpx(self, hwp: Any, end_txt: str = "노블록") -> Tuple[int, int, int]:
        """
        "노블록" 텍스트 위치 찾기 또는 문서 끝 위치 반환 (nob 함수, pyhwpx 버전)
        
        Args:
            hwp: pyhwpx Hwp 객체
            end_txt: 찾을 텍스트 (기본값: "노블록")
            
        Returns:
            (sec, para, pos) 위치 튜플
        """
        try:
            # 문서 처음으로 이동
            hwp.Run("MoveTopLevelBegin")
            
            # find로 텍스트 찾기
            try:
                ok = hwp.find(end_txt)
                if ok:
                    # 선택 해제 및 이전 단어로 이동
                    for cmd in ("Cancel", "MovePrevWord"):
                        try:
                            self._run_cmd_pyhwpx(hwp, cmd)
                        except:
                            pass
                    return self._get_pos_pyhwpx(hwp) or (0, 0, 0)
            except Exception:
                pass
            
            # 찾지 못했으면 문서 끝 위치 반환
            hwp.Run("MoveTopLevelEnd")
            return self._get_pos_pyhwpx(hwp) or (0, 0, 0)
        except Exception as e:
            print(f"[경고] 노블록 위치 찾기 실패: {e}")
            # 폴백: 문서 끝
            try:
                hwp.Run("MoveTopLevelEnd")
                return self._get_pos_pyhwpx(hwp) or (0, 0, 0)
            except:
                return (0, 0, 0)
    
    def _calculate_endnote_end_positions_pyhwpx(
        self, 
        hwp: Any, 
        anchors: List[Tuple[int, int, int]],
        end_txt: str = "노블록"
    ) -> List[Tuple[int, int, int]]:
        """
        각 미주의 끝점 계산 (ene 함수, pyhwpx 버전)
        e_i = a_{i+1} 직전 (MoveLeft), 마지막 e_last = '노블록' 시작 위치
        마지막 문제의 경우 단나누기/페이지나누기 직전까지로 제한
        
        Args:
            hwp: pyhwpx Hwp 객체
            anchors: 미주 앵커 위치 리스트
            end_txt: 문서 끝 마커 텍스트
            
        Returns:
            각 미주의 끝점 위치 리스트
        """
        if not anchors:
            return []
        
        end_positions = []
        
        # 각 미주의 끝점 계산
        for i in range(len(anchors) - 1):
            # 다음 미주 위치로 이동
            next_anchor = anchors[i + 1]
            self._set_pos_pyhwpx(hwp, next_anchor)
            
            # 왼쪽으로 한 칸 이동 (다음 미주 직전)
            try:
                hwp.Run("MoveLeft")
            except:
                pass
            
            end_pos = self._get_pos_pyhwpx(hwp)
            if end_pos:
                end_positions.append(end_pos)
            else:
                # 폴백: 다음 미주 위치 사용
                end_positions.append(next_anchor)
        
        # 마지막 미주의 끝점 = "노블록" 위치 또는 문서 끝
        # 단나누기/페이지나누기 직전까지로 제한
        noblock_pos = self._find_noblock_position_pyhwpx(hwp, end_txt)
        # 마지막 미주에서 noblock_pos까지 가는 중에 단나누기/페이지나누기 확인
        last_anchor = anchors[-1]
        final_end_pos = self._find_end_position_before_breaks_pyhwpx(hwp, last_anchor, noblock_pos)
        end_positions.append(final_end_pos)
        
        return end_positions
    
    def _find_end_position_before_breaks_pyhwpx(
        self,
        hwp: Any,
        start_pos: Tuple[int, int, int],
        max_end_pos: Tuple[int, int, int]
    ) -> Tuple[int, int, int]:
        """
        시작 위치부터 최대 끝 위치까지 가는 중에 단나누기/페이지나누기 직전 위치 찾기
        
        Args:
            hwp: pyhwpx Hwp 객체
            start_pos: 시작 위치
            max_end_pos: 최대 끝 위치 (노블록 또는 문서 끝)
            
        Returns:
            단나누기/페이지나누기 직전 위치 또는 max_end_pos
        """
        try:
            # 시작 위치로 이동
            self._set_pos_pyhwpx(hwp, start_pos)
            
            sec_start, para_start, pos_start = start_pos
            sec_max, para_max, pos_max = max_end_pos
            
            # 시작 위치부터 최대 위치까지 이동하면서 단나누기/페이지나누기 확인
            last_safe_pos = max_end_pos  # 기본값은 최대 위치
            
            # 페이지나누기 찾기 (반복해서 가장 가까운 것 찾기)
            try:
                # 시작 위치로 다시 이동
                self._set_pos_pyhwpx(hwp, start_pos)
                
                # Goto로 페이지나누기 찾기 (반복)
                found_break = False
                for _ in range(10):  # 최대 10번 시도
                    hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
                    try:
                        hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 32)  # 페이지 나누기
                        hwp.HParameterSet.HGotoE.SetSelectionIndex = 6
                        hwp.HParameterSet.HGotoE.IgnoreMessage = 1
                    except:
                        pass
                    
                    result = hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)
                    if result == 0:
                        # 더 이상 페이지나누기가 없음
                        break
                    
                    page_break_pos = self._get_pos_pyhwpx(hwp)
                    if page_break_pos:
                        sec_pb, para_pb, pos_pb = page_break_pos
                        # 시작 위치와 최대 위치 사이에 있는지 확인
                        if (sec_pb > sec_start or (sec_pb == sec_start and para_pb > para_start)) and \
                           (sec_pb < sec_max or (sec_pb == sec_max and para_pb < para_max)):
                            # 페이지나누기 직전으로 이동
                            self._set_pos_pyhwpx(hwp, page_break_pos)
                            hwp.Run("MoveLeft")
                            candidate_pos = self._get_pos_pyhwpx(hwp)
                            if candidate_pos:
                                # 더 가까운 위치 선택 (시작 위치에서 더 가까운 것)
                                if not found_break or \
                                   (sec_pb < sec_max or (sec_pb == sec_max and para_pb < para_max)):
                                    last_safe_pos = candidate_pos
                                    found_break = True
                                    print(f"[디버그] 페이지나누기 발견, 직전 위치로 설정: {last_safe_pos}")
                    else:
                        # 범위를 벗어났으면 종료
                        break
            except Exception as e:
                print(f"[경고] 페이지나누기 찾기 실패: {e}")
            
            # 단나누기는 이미 _prepare_document_pyhwpx에서 제거했지만,
            # 혹시 모를 경우를 대비해 확인
            try:
                self._set_pos_pyhwpx(hwp, start_pos)
                # HeadCtrl로 단나누기 확인
                c = hwp.HeadCtrl
                while c:
                    try:
                        if c.CtrlID == "col":  # 단나누기 컨트롤
                            ctrl_pos = c.GetAnchorPos(0)
                            hwp.SetPosBySet(ctrl_pos)
                            ctrl_anchor_pos = self._get_pos_pyhwpx(hwp)
                            if ctrl_anchor_pos:
                                sec_ctrl, para_ctrl, pos_ctrl = ctrl_anchor_pos
                                # 시작 위치와 최대 위치 사이에 있는지 확인
                                if (sec_ctrl > sec_start or (sec_ctrl == sec_start and para_ctrl > para_start)) and \
                                   (sec_ctrl < sec_max or (sec_ctrl == sec_max and para_ctrl < para_max)):
                                    # 단나누기 직전으로 이동
                                    self._set_pos_pyhwpx(hwp, ctrl_anchor_pos)
                                    hwp.Run("MoveLeft")
                                    candidate_pos = self._get_pos_pyhwpx(hwp)
                                    if candidate_pos:
                                        # 페이지나누기보다 가까우면 업데이트
                                        sec_cand, para_cand, pos_cand = candidate_pos
                                        sec_last, para_last, pos_last = last_safe_pos
                                        if (sec_cand < sec_last or (sec_cand == sec_last and para_cand < para_last)):
                                            last_safe_pos = candidate_pos
                                            print(f"[디버그] 단나누기 발견, 직전 위치로 설정: {last_safe_pos}")
                    except:
                        pass
                    c = c.Next
            except Exception as e:
                print(f"[경고] 단나누기 찾기 실패: {e}")
            
            return last_safe_pos
        except Exception as e:
            print(f"[경고] 단나누기/페이지나누기 직전 위치 찾기 실패: {e}")
            return max_end_pos
    
    def _select_and_save_block_pyhwpx(
        self,
        hwp: Any,
        start_pos: Tuple[int, int, int],
        end_pos: Tuple[int, int, int],
        output_path: str,
        *,
        apply_style: bool = False
    ) -> bool:
        """
        범위 선택 및 블록 저장 (sel + sav 함수, pyhwpx 버전)
        기본 서식 문서로 저장
        
        Args:
            hwp: pyhwpx Hwp 객체
            start_pos: 시작 위치 (sec, para, pos)
            end_pos: 끝 위치 (sec, para, pos)
            output_path: 저장할 파일 경로
            apply_style: True면 Paste 직후 한 문제짜리 새 문서에 스타일 적용. 파싱 시에는 False로 호출.
            
        Returns:
            성공 여부
        """
        if start_pos == end_pos:
            print(f"[경고] 시작 위치와 끝 위치가 같습니다: {start_pos}")
            return False
        
        try:
            # 시작 위치로 이동
            self._set_pos_pyhwpx(hwp, start_pos)
            
            # 선택 시작
            hwp.Run("Select")
            
            # 끝 위치로 이동 (pyhwpx에서는 Select 후 SetPos로 이동해도 범위 선택이 유지됨)
            self._set_pos_pyhwpx(hwp, end_pos)
            
            # 선택된 내용 복사
            hwp.Run("Copy")
            
            # 선택 해제
            hwp.Run("Cancel")
            
            # 새 문서 생성 (기본 서식)
            hwp.HAction.GetDefault("FileNew", hwp.HParameterSet.HFileOpenSave.HSet)
            hwp.HAction.Execute("FileNew", hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 붙여넣기
            hwp.Run("Paste")
            if apply_style:
                time.sleep(0.12)  # Paste 직후 문서 안정화 (0.1~0.15초)
                self.apply_style_to_current_document_pyhwpx(hwp)
            
            # 기본 서식 문서로 저장
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            hwp.HAction.GetDefault("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
            hwp.HParameterSet.HFileOpenSave.filename = output_path
            hwp.HParameterSet.HFileOpenSave.Format = "HWP"
            hwp.HAction.Execute("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
            
            # 새 문서 닫기
            hwp.HAction.GetDefault("FileClose", hwp.HParameterSet.HFileOpenSave.HSet)
            hwp.HAction.Execute("FileClose", hwp.HParameterSet.HFileOpenSave.HSet)
            
            return os.path.exists(output_path)
        except Exception as e:
            print(f"[경고] 범위 선택 및 저장 실패: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _is_endnote_alone_on_line_pyhwpx(self, hwp: Any, anchor_pos: Tuple[int, int, int]) -> bool:
        """
        미주가 그 줄에 단독으로 있는지 확인
        미주 위치에서 오른쪽 방향키를 2번 눌렀을 때 줄이 바뀌면 미주만 있는 것
        
        Args:
            hwp: pyhwpx Hwp 객체
            anchor_pos: 미주 앵커 위치 (sec, para, pos)
            
        Returns:
            미주가 단독이면 True, 아니면 False
        """
        try:
            # 현재 위치 저장
            current_pos = self._get_pos_pyhwpx(hwp)
            
            # 미주 앵커 위치로 이동
            self._set_pos_pyhwpx(hwp, anchor_pos)
            
            # 현재 줄 번호 확인
            current_para = anchor_pos[1]
            
            # 오른쪽 방향키 2번
            hwp.Run("MoveRight")
            hwp.Run("MoveRight")
            
            # 이동 후 줄 번호 확인
            new_pos = self._get_pos_pyhwpx(hwp)
            new_para = new_pos[1] if new_pos else current_para
            
            # 원래 위치로 복원
            if current_pos:
                self._set_pos_pyhwpx(hwp, current_pos)
            
            # 줄이 바뀌었으면 미주만 있는 것
            return new_para != current_para
            
        except Exception as e:
            print(f"[경고] 미주 단독 확인 실패: {e}")
            return False
    
    def _separate_endnote_to_new_line_pyhwpx(self, hwp: Any, anchor_pos: Tuple[int, int, int]) -> None:
        """
        미주를 새 줄로 분리 (같은 줄에 다른 텍스트가 있을 때)
        
        Args:
            hwp: pyhwpx Hwp 객체
            anchor_pos: 미주 앵커 위치 (sec, para, pos)
        """
        try:
            # 미주 앵커 위치로 이동
            self._set_pos_pyhwpx(hwp, anchor_pos)
            
            # 오른쪽으로 이동
            hwp.Run("MoveRight")
            
            # 문단 나누기
            hwp.Run("BreakPara")
            
            # 새 줄 맨 앞이 공백/탭일 때만 1글자 제거 (GetTextFile 없이 클립보드로 판별)
            hwp.Run("MoveLineBegin")
            hwp.Run("MoveSelRight")
            saved_clipboard = None
            try:
                if CLIPBOARD_AVAILABLE:
                    try:
                        win32clipboard.OpenClipboard()
                        try:
                            saved_clipboard = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                        except (TypeError, OSError):
                            try:
                                saved_clipboard = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
                                if isinstance(saved_clipboard, bytes):
                                    saved_clipboard = saved_clipboard.decode("utf-8", errors="replace")
                            except (TypeError, OSError):
                                pass
                        finally:
                            win32clipboard.CloseClipboard()
                    except Exception:
                        pass
                    hwp.Run("Copy")
                    first_char = None
                    try:
                        win32clipboard.OpenClipboard()
                        try:
                            first_char = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                        except (TypeError, OSError):
                            try:
                                raw = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
                                first_char = (raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else raw) if raw else None
                            except (TypeError, OSError):
                                pass
                        finally:
                            win32clipboard.CloseClipboard()
                    except Exception:
                        pass
                    if first_char is not None and len(first_char) == 1 and first_char in " \t":
                        hwp.Run("Delete")
                    else:
                        hwp.Run("Cancel")
                else:
                    hwp.Run("Cancel")
            except Exception:
                try:
                    hwp.Run("Cancel")
                except Exception:
                    pass
            finally:
                if CLIPBOARD_AVAILABLE and saved_clipboard is not None:
                    try:
                        win32clipboard.OpenClipboard()
                        try:
                            win32clipboard.EmptyClipboard()
                            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, saved_clipboard)
                        finally:
                            win32clipboard.CloseClipboard()
                    except Exception:
                        pass
            
            # 위로 이동 (미주가 있는 줄로)
            hwp.Run("MoveUp")
            
        except Exception as e:
            print(f"[경고] 미주 분리 실패: {e}")
    
    def _handle_dialog_auto_yes_pyhwpx(self, hwp: Any) -> None:
        """
        HWP 대화상자에서 "찾음(Y)"를 자동으로 클릭
        "문서의 끝까지 찾았습니다" 대화상자가 나타날 때 자동으로 Y를 누름
        
        Args:
            hwp: pyhwpx Hwp 객체
        """
        if not WIN32_AVAILABLE:
            return
        
        try:
            # pyhwpx의 내부 win32com 객체 가져오기
            if hasattr(hwp, '_hwp'):
                hwp_com = hwp._hwp
            elif hasattr(hwp, 'Hwnd'):
                hwp_com = hwp
            else:
                return
            
            # 한글 창 핸들 가져오기
            hwnd = hwp_com.Hwnd
            if not hwnd:
                return
            
            # 대화상자가 나타날 때까지 잠시 대기
            time.sleep(0.2)
            
            # 대화상자 창 찾기
            dialog_windows = []
            
            def enum_windows_callback(hwnd_dialog, windows):
                if win32gui.IsWindowVisible(hwnd_dialog):
                    window_text = win32gui.GetWindowText(hwnd_dialog)
                    if "찾았습니다" in window_text or "찾음" in window_text or "계속 찾을까요" in window_text:
                        windows.append(hwnd_dialog)
                return True
            
            # 모든 최상위 창 검색
            def enum_top_windows_callback(hwnd_top, windows):
                if win32gui.IsWindowVisible(hwnd_top):
                    window_text = win32gui.GetWindowText(hwnd_top)
                    if "찾았습니다" in window_text or "찾음" in window_text or "계속 찾을까요" in window_text:
                        windows.append(hwnd_top)
                return True
            
            win32gui.EnumWindows(enum_top_windows_callback, dialog_windows)
            
            if dialog_windows:
                # 대화상자에 Y 키 전송
                for dialog_hwnd in dialog_windows:
                    # Y 키 전송 (한글 입력 고려)
                    win32api.PostMessage(dialog_hwnd, win32con.WM_KEYDOWN, ord('Y'), 0)
                    time.sleep(0.05)
                    win32api.PostMessage(dialog_hwnd, win32con.WM_KEYUP, ord('Y'), 0)
                    time.sleep(0.1)
        except Exception as e:
            # 대화상자 처리 실패해도 계속 진행
            pass
    
    def _goto_endnote_body_with_auto_yes_pyhwpx(self, hwp: Any, anchor_pos: Tuple[int, int, int]) -> bool:
        """
        미주 본문으로 진입 (대화상자 자동 처리)
        
        Args:
            hwp: pyhwpx Hwp 객체
            anchor_pos: 미주 앵커 위치
            
        Returns:
            성공 여부
        """
        try:
            # 미주 앵커 위치로 이동
            self._set_pos_pyhwpx(hwp, anchor_pos)
            
            # 미주 본문으로 진입 (Goto) - 파싱된 문서에는 미주가 하나뿐이므로 1-2회만 시도
            hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
            hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 31)  # 미주로 이동
            hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
            hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)
            
            # 대화상자 자동 처리 (최대 2회 시도)
            for _ in range(2):
                self._handle_dialog_auto_yes_pyhwpx(hwp)
                time.sleep(0.1)
            
            return True
        except Exception as e:
            print(f"[경고] 미주 본문 진입 실패: {e}")
            return False
    
    def _apply_endnote_style_pyhwpx(self, hwp: Any, anchor_pos: Tuple[int, int, int]) -> None:
        """
        미주에 스타일 적용 (정확한 순서)
        4. 미주만 선택 후 세방고딕 Bold, 20pt, 줄간격 120%, 색 (43, 45, 99)
        5. 미주 선택 해제 (오른쪽 한 칸 이동)
        6. 주석 진입
        7. 전체선택 후 나눔고딕, 10pt, 줄간격 160%, 검정색
        8. 주석 나가기
        
        Args:
            hwp: pyhwpx Hwp 객체
            anchor_pos: 미주 앵커 위치 (sec, para, pos)
        """
        try:
            # 4단계: 미주만 선택 후 세방고딕 Bold, 20pt, 줄간격 120%, 색 (43, 45, 99)
            # 미주 앵커 위치로 이동
            self._set_pos_pyhwpx(hwp, anchor_pos)
            
            # 미주만 선택 (오른쪽으로 확장)
            hwp.Run("MoveSelRight")
            
            # 글꼴: 세방고딕 Bold
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FaceNameHangul = "세방고딕 Bold"
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1  # TTF
            hwp.HParameterSet.HCharShape.FaceNameLatin = "세방고딕 Bold"
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FaceNameHanja = "세방고딕 Bold"
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FaceNameJapanese = "세방고딕 Bold"
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FaceNameOther = "세방고딕 Bold"
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FaceNameSymbol = "세방고딕 Bold"
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FaceNameUser = "세방고딕 Bold"
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 글자 크기: 20pt
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            hwp.HParameterSet.HCharShape.Height = 2000  # 20pt
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 글자 색상: RGB(43, 45, 99) - HWP는 BGR 순서 사용
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            # BGR 순서: R43, G45, B99 -> BGR(99, 45, 43)
            hwp.HParameterSet.HCharShape.TextColor = (99 << 16) | (45 << 8) | 43
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 줄간격 120%: CharShape 3회 직후 선택 해제 가능 → 위치복원→SelectPara→0.1초→재선택→0.05초→ParagraphShape
            self._set_pos_pyhwpx(hwp, anchor_pos)
            hwp.Run("SelectPara")
            time.sleep(0.1)
            hwp.Run("SelectPara")
            time.sleep(0.05)
            hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
            hwp.HParameterSet.HParaShape.LineSpacing = 120
            if hasattr(hwp.HParameterSet.HParaShape, "LineSpacingType"):
                hwp.HParameterSet.HParaShape.LineSpacingType = 0  # Percent
            hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
            
            # 5단계: 미주 선택 해제 (왼쪽 한 칸 이동)
            hwp.Run("MoveLeft")
            
            # 6단계: 주석 진입 (미주 앵커 위치로 다시 이동 후 진입)
            self._set_pos_pyhwpx(hwp, anchor_pos)  # 미주 앵커 위치 복원
            
            hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
            hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 32)  # 주석으로 이동
            hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
            hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)
            
            # 대화상자 자동 처리
            for _ in range(2):
                self._handle_dialog_auto_yes_pyhwpx(hwp)
                time.sleep(0.1)
            
            # 7단계: 전체선택 후 나눔고딕, 10pt, 줄간격 160%, 검정색
            hwp.Run("SelectAll")  # 주석 내부 전체 선택
            
            # 글꼴: 나눔고딕
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FaceNameHangul = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1  # TTF
            hwp.HParameterSet.HCharShape.FaceNameLatin = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FaceNameHanja = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FaceNameJapanese = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FaceNameOther = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FaceNameSymbol = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FaceNameUser = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 글자 크기: 10pt
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            hwp.HParameterSet.HCharShape.Height = 1000  # 10pt
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 줄간격 160%: Goto(주석)+CharShape 직후 선택 불안정 → SelectAll→0.1초→재선택→0.05초→ParagraphShape
            hwp.Run("SelectAll")
            time.sleep(0.1)
            hwp.Run("SelectAll")
            time.sleep(0.05)
            hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
            hwp.HParameterSet.HParaShape.LineSpacing = 160
            if hasattr(hwp.HParameterSet.HParaShape, "LineSpacingType"):
                hwp.HParameterSet.HParaShape.LineSpacingType = 0  # Percent
            hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
            
            # 글씨색: 검정색
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            hwp.HParameterSet.HCharShape.TextColor = 0  # 검정색
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 8단계: 주석 나가기
            hwp.HAction.GetDefault("Goto", hwp.HParameterSet.HGotoE.HSet)
            hwp.HParameterSet.HGotoE.HSet.SetItem("DialogResult", 32)  # 주석에서 나오기
            hwp.HParameterSet.HGotoE.SetSelectionIndex = 5
            hwp.HAction.Execute("Goto", hwp.HParameterSet.HGotoE.HSet)
            
            # CloseEx 실행
            hwp.Run("CloseEx")
            
            # 대화상자 자동 처리 (나올 때도)
            for _ in range(2):
                self._handle_dialog_auto_yes_pyhwpx(hwp)
                time.sleep(0.1)
            
        except Exception as e:
            print(f"[경고] 미주 스타일 적용 실패: {e}")
            import traceback
            traceback.print_exc()
    
    def _apply_global_font_pyhwpx(self, hwp: Any) -> None:
        """
        전체 문서의 글꼴을 나눔고딕, 10pt, 줄간격 160%로 변경
        
        Args:
            hwp: pyhwpx Hwp 객체
        """
        try:
            # 전체 선택
            hwp.Run("SelectAll")
            
            # 글머리표 제거 (선택 유지, Cancel 없이 바로 다음 단계로)
            hwp.HAction.GetDefault("BulletDlg", hwp.HParameterSet.HParaShape.HSet)
            if hasattr(hwp.HParameterSet.HParaShape, "HeadingType"):
                hwp.HParameterSet.HParaShape.HeadingType = 0  # None
            hwp.HAction.Execute("BulletDlg", hwp.HParameterSet.HParaShape.HSet)
            
            # 글꼴: 나눔고딕
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FaceNameHangul = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1  # TTF
            hwp.HParameterSet.HCharShape.FaceNameLatin = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FaceNameHanja = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FaceNameJapanese = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FaceNameOther = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FaceNameSymbol = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FaceNameUser = "나눔고딕"
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 글자 크기: 10pt
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            # 10pt를 HWP 단위로 변환 (1pt = 100 HWP 단위)
            hwp.HParameterSet.HCharShape.Height = 1000
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 줄간격 160%: CharShape 직후 선택 해제 가능 → 선택→0.1초→재선택→0.05초→ParagraphShape
            hwp.Run("SelectAll")
            time.sleep(0.1)
            hwp.Run("SelectAll")
            time.sleep(0.05)
            hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
            hwp.HParameterSet.HParaShape.LineSpacing = 160
            if hasattr(hwp.HParameterSet.HParaShape, "LineSpacingType"):
                hwp.HParameterSet.HParaShape.LineSpacingType = 0  # Percent
            hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
            
            # 글씨색: 검정색 (RGB(0, 0, 0))
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.FontTypeHangul = 1
            hwp.HParameterSet.HCharShape.FontTypeLatin = 1
            hwp.HParameterSet.HCharShape.FontTypeHanja = 1
            hwp.HParameterSet.HCharShape.FontTypeJapanese = 1
            hwp.HParameterSet.HCharShape.FontTypeOther = 1
            hwp.HParameterSet.HCharShape.FontTypeSymbol = 1
            hwp.HParameterSet.HCharShape.FontTypeUser = 1
            hwp.HParameterSet.HCharShape.TextColor = 0  # 검정색 (RGB(0, 0, 0))
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
            
            # 선택 해제
            hwp.Run("MoveRight")
            
        except Exception as e:
            print(f"[경고] 전체 글꼴 적용 실패: {e}")
            import traceback
            traceback.print_exc()
    
    def apply_style_to_current_document_pyhwpx(self, hwp: Any) -> None:
        """
        한 문제짜리 새 문서(현재 열린 문서)에만 스타일 적용.
        본문 160%, 미주 숫자 120%, 주석 본문 160% 등. 파싱 경로에서는 호출하지 않음.
        
        Args:
            hwp: pyhwpx Hwp 객체 (한 문제만 붙여넣은 새 문서가 열린 상태)
        """
        self._apply_endnote_styles_pyhwpx(hwp)
    
    def _apply_endnote_styles_pyhwpx(self, hwp: Any) -> None:
        """
        파싱된 문제 문서의 모든 미주에 스타일 적용
        1. 먼저 전체 글꼴을 나눔고딕, 10pt, 줄간격 160%로 변경
        2. 각 미주가 단독인지 확인
        3. 단독이 아니면 분리
        4. 미주 스타일 적용
        
        Args:
            hwp: pyhwpx Hwp 객체 (새로 붙여넣은 문서)
        """
        try:
            # 1. 먼저 전체 글꼴을 나눔고딕, 10pt, 줄간격 160%로 변경
            self._apply_global_font_pyhwpx(hwp)
            
            # 2. 미주 찾기 (파싱된 문서에는 미주가 하나뿐이므로 첫 번째만 찾기)
            anchor_pos = None
            c = hwp.HeadCtrl
            while c:
                try:
                    if c.CtrlID == "en":  # 미주 컨트롤
                        anchor_posset = c.GetAnchorPos(0)
                        hwp.SetPosBySet(anchor_posset)
                        anchor_pos = self._get_pos_pyhwpx(hwp)
                        if anchor_pos:
                            break  # 첫 번째 미주만 찾고 종료
                except Exception:
                    pass
                c = c.Next
            
            if anchor_pos is None:
                print("[경고] 미주를 찾을 수 없습니다.")
                return
            
            # 3단계: 미주 옆에 요소가 있는지 확인하고 있으면 오른쪽방향, 엔터, 위쪽방향, 없으면 OK
            is_alone = self._is_endnote_alone_on_line_pyhwpx(hwp, anchor_pos)
            
            if not is_alone:
                # 단독이 아니면 먼저 분리
                self._separate_endnote_to_new_line_pyhwpx(hwp, anchor_pos)
            
            # 4-8단계: 미주 스타일 적용
            self._apply_endnote_style_pyhwpx(hwp, anchor_pos)
            
            # 9단계: 문서 끝 이동
            hwp.Run("MoveTopLevelEnd")
            
        except Exception as e:
            print(f"[경고] 미주 스타일 적용 실패: {e}")
            import traceback
            traceback.print_exc()