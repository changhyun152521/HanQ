"""
HWP Composer (MVP)

편집된 problem_id 리스트를 받아 "문항 원본 HWP 블록"들을 하나의 HWP 문서로 합칩니다.

현재 구현 목표:
- 문제 HWP 원본(content_raw, GridFS)을 임시 파일로 복원
- 새 HWP 문서를 만들고, 문제 문서를 하나씩 열어서 Ctrl+A → Copy → (출력 문서) Paste
- output_path로 저장

주의:
- HWP COM 자동화 특성상 환경/버전 차이가 있을 수 있어, 예외/팝업 억제에 최대한 안전하게 구현합니다.
"""

from __future__ import annotations

import logging
import os
import sys
import time
from datetime import datetime
from typing import List, Optional, Tuple

from database.sqlite_connection import SQLiteConnection
from database.repositories import ProblemRepository
from utils.hwp_restore import HWPRestore
from processors.hwp.hwp_reader import HWPReader


logger = logging.getLogger(__name__)


class WorksheetComposeError(Exception):
    pass


class WorksheetHwpComposer:
    def __init__(self, db_connection: SQLiteConnection):
        self.db = db_connection
        self.restore = HWPRestore(db_connection)
        self.problem_repo = ProblemRepository(db_connection)
        self._template_missing = False  # compose() 시 템플릿을 못 찾아 빈 문서로 생성했으면 True

    def _resolve_default_template_path(self) -> Optional[str]:
        """
        템플릿 파일 경로를 찾습니다.

        우선순위:
        1) EXE 옆 templates/worksheet_template.hwp (배포 환경)
        2) 프로젝트 templates/worksheet_template.hwp (개발 환경)
        3) 공용 폴더 C:\\Users\\Public\\CH_LMS\\templates\\worksheet_template.hwp
        """
        candidates: List[str] = []
        # 일부 환경에서 파일명이 "worksheet_template.hwp.hwp"처럼 이중 확장자로 저장되는 경우가 있어
        # 두 케이스를 모두 허용합니다(권장은 worksheet_template.hwp).
        template_names = ["worksheet_template.hwp", "worksheet_template.hwp.hwp"]

        # 1) 배포(EXE) 기준
        try:
            base = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])
            if base:
                for name in template_names:
                    candidates.append(os.path.join(base, "templates", name))
        except Exception:
            pass

        # 2) 프로젝트 기준
        try:
            repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
            for name in template_names:
                candidates.append(os.path.join(repo_root, "templates", name))
        except Exception:
            pass

        # 3) 공용 폴더(권장)
        for name in template_names:
            candidates.append(os.path.join(r"C:\Users\Public\CH_LMS\templates", name))

        for p in candidates:
            try:
                if p and os.path.exists(p):
                    return p
            except Exception:
                continue
        return None

    def _compute_scope_text(self, problem_ids: List[str]) -> str:
        """
        범위 텍스트: "첫 소단원 ~ 마지막 소단원"
        - 문제 순서 기준(드래그/재정렬 결과가 반영됨)
        - 과목/대단원 혼합이어도 sub_unit만 사용
        """
        first: Optional[str] = None
        last: Optional[str] = None

        for pid in problem_ids:
            try:
                p = self.problem_repo.find_by_id(str(pid))
            except Exception:
                p = None
            if not p or not getattr(p, "tags", None):
                continue

            # tags 중 sub_unit이 있는 첫 값을 사용(보통 tags[0]이 대표 태그)
            sub = None
            try:
                for t in (p.tags or []):
                    v = (getattr(t, "sub_unit", None) or "").strip()
                    if v:
                        sub = v
                        break
            except Exception:
                sub = None

            if not sub:
                continue
            if first is None:
                first = sub
            last = sub

        if not first and not last:
            return "(범위 미지정)"
        if first and last and first != last:
            return f"{first} ~ {last}"
        return first or last or "(범위 미지정)"

    def _try_put_field_text(self, hwp, field_name: str, value: str) -> bool:
        """
        HWP Field/북마크에 텍스트 주입을 시도합니다.
        - 성공하면 True
        - 실패하면 False(폴백: 검색/치환)
        """
        if not hwp or not field_name:
            return False
        try:
            # HWP COM: PutFieldText(field, text)
            hwp.PutFieldText(str(field_name), str(value))
            return True
        except Exception:
            return False

    def _paste_text(self, hwp, text: str) -> None:
        """클립보드 기반으로 텍스트를 붙여넣습니다(가장 범용)."""
        try:
            import win32clipboard  # type: ignore

            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                # CF_UNICODETEXT = 13
                win32clipboard.SetClipboardData(13, str(text))
            finally:
                win32clipboard.CloseClipboard()
            self._safe_run(hwp, "Paste")
            return
        except Exception:
            pass

        # 폴백: 가능하면 InsertText 계열 시도(환경별 상이)
        try:
            if hasattr(hwp, "InsertText"):
                hwp.InsertText(str(text))
                return
        except Exception:
            pass

        # 최후: 아무 것도 못하면 무시(호출부에서 검증)
        return

    def _move_to_field_or_text(self, reader: HWPReader, name: str) -> bool:
        """
        북마크/필드 이동을 우선 시도하고, 실패하면 텍스트(더미 문자열) 찾기로 폴백합니다.

        성공 시 커서가 해당 위치(또는 선택 상태)에 있도록 만듭니다.
        """
        hwp = reader.hwp
        if hwp is None:
            return False
        nm = (name or "").strip()
        if not nm:
            return False

        # 1) MoveToField 메서드 (환경에 따라 제공)
        try:
            if hasattr(hwp, "MoveToField"):
                r = hwp.MoveToField(nm)
                # 반환값이 없거나 0/1인 경우가 있어, 예외가 아니면 성공 취급
                return True if r is None else bool(r)
        except Exception:
            pass

        # 2) 텍스트 찾기 폴백(템플릿에 더미 문자열로 넣어둔 경우)
        try:
            found = reader.find_text(nm, start_from_beginning=True, move_after=False)
            return bool(found is not None)
        except Exception:
            return False

    def _replace_at_marker(self, reader: HWPReader, marker: str, value: str) -> bool:
        """
        marker(보통 텍스트 더미/HDR_* 등)를 찾아 선택한 뒤,
        marker 텍스트를 지우고(value가 있으면) value를 입력합니다.

        중요:
        - 템플릿에서 marker가 "텍스트"로 들어있는 경우를 가장 우선 지원합니다.
        - MoveToField는 커서만 이동하고 텍스트 선택을 보장하지 않는 환경이 있어,
          삭제/치환은 기본적으로 find_text(선택 상태)를 기준으로 수행합니다.
        """
        hwp = reader.hwp
        if hwp is None:
            return False
        mk = (marker or "").strip()
        if not mk:
            return False

        # 1) 텍스트 찾기 기반(선택 상태 보장) - 가장 안정적
        found = False
        try:
            found = reader.find_text(mk, start_from_beginning=True, move_after=False) is not None
            if not found:
                found = reader.find_text(mk, start_from_beginning=False, move_after=False) is not None
        except Exception:
            found = False

        # 2) (폴백) 필드 이동 시도 후, 그 위치에서 다시 텍스트 찾기
        if not found:
            try:
                if hasattr(hwp, "MoveToField"):
                    hwp.MoveToField(mk)
                found = reader.find_text(mk, start_from_beginning=False, move_after=False) is not None
            except Exception:
                found = False

        if not found:
            return False

        # marker 텍스트 삭제(선택된 상태라고 가정)
        self._safe_run(hwp, "Cut")

        # 값 입력 (빈 문자열이면 삭제만 수행)
        if (value or "") != "":
            self._paste_text(hwp, value)

        # 선택/상태 정리
        self._safe_run(hwp, "Cancel")
        return True

    def _ascii_temp_dir(self) -> str:
        """
        HWP COM이 유니코드/긴 경로에서 Open 실패하는 케이스가 있어,
        가능한 한 ASCII 기반의 짧은 경로를 사용합니다.
        """
        base = os.environ.get("SystemDrive", "C:") + os.sep
        d = os.path.join(base, "CH_LMS_TMP", "worksheet_blocks")
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            # 폴더 생성 실패 시에도 진행은 하되, 이후 Open 실패 가능성이 큼
            pass
        return d

    def _restore_all(self, problem_ids: List[str]) -> List[Tuple[str, str]]:
        """
        각 problem_id의 원본 HWP를 ASCII temp dir로 복원합니다.
        Returns: [(problem_id, restored_path), ...]
        """
        out: List[Tuple[str, str]] = []
        temp_dir = self._ascii_temp_dir()
        for i, pid in enumerate(problem_ids, start=1):
            safe_pid = "".join(ch for ch in str(pid) if ch.isalnum())[:24] or str(i)
            path = os.path.join(temp_dir, f"ws_prob_{i:04d}_{safe_pid}.hwp")
            restored = self.restore.restore_to_file(pid, output_path=path)
            out.append((pid, restored))
        return out

    def _safe_run(self, hwp, action_name: str) -> None:
        try:
            hwp.HAction.Run(action_name)
        except Exception:
            pass

    def _diag_config(self) -> dict:
        """
        강제 실험/원인 고립을 위한 진단 모드 설정.

        환경변수:
        - CH_LMS_HWP_DIAG_MODE:
            - normal (기본): 기존 전체 플로우
            - header_only: 헤더 치환만 수행 후 저장
            - body_only: 본문(PROLEMS_HERE 제거 + 1문항 Paste)만 수행 후 저장
            - combined: header_only + body_only 순서로 수행 후 저장
        - CH_LMS_HWP_DIAG_VISIBLE=1: 한글 창 표시(단계 관찰용)
        - CH_LMS_HWP_DIAG_SLEEP_SEC=0.8: 단계별 대기(초)
        - CH_LMS_HWP_DIAG_HEADER_EXIT=0: 헤더 종료(닫기) 비활성화(헤더-단독 모드에서만 권장)
        - CH_LMS_HWP_DIAG_HEADER_REPLACE=allreplace|cell:
            - allreplace: AllReplace/ExecReplace 사용
            - cell: RepeatFind→선택삭제(Delete)→InsertText 기반(Replace 계열 금지) 실험
        """
        mode = (os.environ.get("CH_LMS_HWP_DIAG_MODE", "normal") or "normal").strip().lower()
        visible = (os.environ.get("CH_LMS_HWP_DIAG_VISIBLE", "0") or "0").strip() == "1"
        try:
            sleep_sec = float(os.environ.get("CH_LMS_HWP_DIAG_SLEEP_SEC", "0") or 0.0)
        except Exception:
            sleep_sec = 0.0
        header_exit = (os.environ.get("CH_LMS_HWP_DIAG_HEADER_EXIT", "1") or "1").strip() != "0"
        # 기본값은 안정성이 높은 "cell" 치환을 사용합니다.
        # (AllReplace/ExecReplace는 헤더 내 표 컨트롤을 깨거나, 저장/닫기 시 '빈 머리말'이 추가 생성되는 케이스가 있음)
        header_replace = (os.environ.get("CH_LMS_HWP_DIAG_HEADER_REPLACE", "cell") or "cell").strip().lower()

        return {
            "mode": mode,
            "visible": bool(visible),
            "sleep_sec": float(max(0.0, sleep_sec)),
            "header_exit": bool(header_exit),
            "header_replace": header_replace if header_replace in ("allreplace", "cell") else "cell",
        }

    def _diag_pause(self, cfg: dict, label: str) -> None:
        sec = float(cfg.get("sleep_sec") or 0.0)
        if sec <= 0:
            return
        try:
            print(f"[HWP_DIAG] {label} (sleep {sec:.2f}s)")
        except Exception:
            pass
        try:
            time.sleep(sec)
        except Exception:
            return

    def _set_hwp_visible(self, reader: HWPReader, visible: bool) -> None:
        """진단용: 한글 창 표시/숨김."""
        hwp = reader.hwp
        if hwp is None:
            return
        try:
            hwp.XHwpWindows.Item(0).Visible = bool(visible)
        except Exception:
            pass

    def _select_token_in_current_context(self, reader: HWPReader, token: str, find_type: int = 1) -> bool:
        """
        현재 편집 컨텍스트에서 token을 RepeatFind로 찾아 '선택된 상태'로 둡니다.
        - _try_repeat_find와 달리 선택을 접지 않습니다(삭제/삽입 실험용).
        """
        hwp = reader.hwp
        if hwp is None:
            return False
        s = (token or "").strip()
        if not s:
            return False
        try:
            with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
                fr = hwp.HParameterSet.HFindReplace
                try:
                    fr.FindString = s
                except Exception:
                    try:
                        fr.HSet.SetItem("FindString", s)
                    except Exception:
                        pass
                try:
                    fr.IgnoreMessage = 1
                except Exception:
                    try:
                        fr.HSet.SetItem("IgnoreMessage", 1)
                    except Exception:
                        pass
                try:
                    fr.HSet.SetItem("FindType", int(find_type))
                except Exception:
                    pass
                r = hwp.HAction.Execute("RepeatFind", fr.HSet)
            return r == 1
        except Exception:
            return False

    def _select_token_default_scope(self, reader: HWPReader, token: str) -> bool:
        """
        RepeatFind로 token을 찾아 '선택된 상태'로 둡니다.
        - FindType을 강제로 설정하지 않습니다(환경별로 헤더/본문 스코프가 달라지는 문제 회피).
        - HeaderFooter(머리말 편집 진입)를 호출하지 않기 위해 사용합니다.
        """
        hwp = reader.hwp
        if hwp is None:
            return False
        s = (token or "").strip()
        if not s:
            return False
        try:
            with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
                fr = hwp.HParameterSet.HFindReplace
                try:
                    fr.FindString = s
                except Exception:
                    try:
                        fr.HSet.SetItem("FindString", s)
                    except Exception:
                        pass
                try:
                    fr.IgnoreMessage = 1
                except Exception:
                    try:
                        fr.HSet.SetItem("IgnoreMessage", 1)
                    except Exception:
                        pass
                r = hwp.HAction.Execute("RepeatFind", fr.HSet)
            return r == 1
        except Exception:
            return False

    def _replace_first_by_repeat_find(self, reader: HWPReader, find: str, replace: str) -> bool:
        """
        문서 전체(기본 스코프)에서 find를 1회 찾아 선택 후 Delete→InsertText로 치환합니다.
        - HeaderFooter 호출 금지(빈 머리말 컨트롤 생성 방지)
        - AllReplace/ExecReplace 호출 금지(헤더 표/컨트롤 깨짐 방지)
        """
        hwp = reader.hwp
        if hwp is None:
            return False

        # 문서 처음부터 찾기
        try:
            hwp.HAction.GetDefault("MoveDocBegin", hwp.HParameterSet.HSelectionOpt.HSet)
            hwp.HAction.Execute("MoveDocBegin", hwp.HParameterSet.HSelectionOpt.HSet)
        except Exception:
            pass

        if not self._select_token_default_scope(reader, find):
            return False

        # 선택된 텍스트만 치환 (표/컨트롤 전체 삭제 방지)
        self._replace_selected_text_with_insert(reader, replace)

        # 선택 상태를 조용히 접기
        try:
            hwp.HAction.GetDefault("MoveSelEnd", hwp.HParameterSet.HSelectionOpt.HSet)
            hwp.HAction.Execute("MoveSelEnd", hwp.HParameterSet.HSelectionOpt.HSet)
        except Exception:
            pass
        try:
            hwp.HAction.GetDefault("MoveRight", hwp.HParameterSet.HSelectionOpt.HSet)
            hwp.HAction.Execute("MoveRight", hwp.HParameterSet.HSelectionOpt.HSet)
        except Exception:
            pass
        return True

    def _replace_selected_text_with_insert(self, reader: HWPReader, value: str) -> None:
        """
        선택된 텍스트를 'Delete'로 제거 후 InsertText(가능하면)로 삽입합니다.
        - Cut/클립보드 의존을 피하는 실험용 루트
        """
        hwp = reader.hwp
        if hwp is None:
            return
        try:
            self._safe_run(hwp, "Delete")
        except Exception:
            pass
        try:
            if hasattr(hwp, "InsertText"):
                hwp.InsertText(str(value or ""))
                return
        except Exception:
            pass
        # 폴백: 텍스트 붙여넣기(클립보드)
        try:
            self._paste_text(hwp, str(value or ""))
        except Exception:
            return

    def _try_enter_header_edit(self, reader: HWPReader) -> bool:
        """
        머리말 편집 모드로 진입합니다.
        (사용자 매크로 기반: HeaderFooterStyle, HeaderFooterCtrlType=0)
        """
        return self._try_enter_header_edit_with_style(reader, style=0, ctrl_type=0)

    def _try_enter_header_edit_with_style(self, reader: HWPReader, *, style: int, ctrl_type: int = 0) -> bool:
        """
        HeaderFooterStyle을 지정해 머리말 편집 모드로 진입합니다.
        - style: HeaderFooterStyle 후보(보통 0~7 내에서 순회)
        - ctrl_type: 0=머리말(헤더)
        """
        hwp = reader.hwp
        if hwp is None:
            return False
        try:
            with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                hwp.HAction.GetDefault("HeaderFooter", hwp.HParameterSet.HHeaderFooter.HSet)
                try:
                    hwp.HParameterSet.HHeaderFooter.HSet.SetItem("HeaderFooterStyle", int(style))
                    hwp.HParameterSet.HHeaderFooter.HSet.SetItem("HeaderFooterCtrlType", int(ctrl_type))
                except Exception:
                    pass
                r = hwp.HAction.Execute("HeaderFooter", hwp.HParameterSet.HHeaderFooter.HSet)
            return True if r is None else bool(r)
        except Exception:
            return False

    def _try_exit_header_edit(self, reader: HWPReader) -> None:
        """
        머리말 편집 종료(본문 복귀) 시도.
        - 사용자 매크로 기록과 동일하게 Close를 우선 호출합니다.
        - 실패 시에만 CloseEx(Shift+Esc)를 폴백으로 시도합니다.
        - 헤더 종료 단계에서는 Cancel(ESC)을 최소화합니다.
          (표 셀 편집 상태에서 ESC가 표 컨트롤 선택으로 승격되는 케이스를 피하기 위함)
        """
        hwp = reader.hwp
        if hwp is None:
            return
        try:
            with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                # ✅ 종료 직전: 선택(특히 표 컨트롤 선택) 상태를 최대한 접습니다.
                # ESC(Cancel)는 표 컨트롤 선택으로 승격될 수 있어 사용하지 않습니다.
                try:
                    hwp.HAction.GetDefault("MoveSelEnd", hwp.HParameterSet.HSelectionOpt.HSet)
                    hwp.HAction.Execute("MoveSelEnd", hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception:
                    pass
                try:
                    hwp.HAction.GetDefault("MoveRight", hwp.HParameterSet.HSelectionOpt.HSet)
                    hwp.HAction.Execute("MoveRight", hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception:
                    pass
                try:
                    # 매크로와 동일: Close 우선
                    hwp.HAction.Run("Close")
                except Exception:
                    # 폴백: Shift+Esc 동작
                    try:
                        hwp.HAction.Run("CloseEx")
                    except Exception:
                        pass
        except Exception:
            return

    def _try_repeat_find(self, reader: HWPReader, *, find: str, find_type: int = 1) -> bool:
        """
        RepeatFind로 문자열 존재 여부를 확인합니다.
        - 성공 시 선택 상태가 되므로 Cancel로 정리합니다.
        - find_type은 매크로에서 사용한 값을 기본(1)으로 둡니다.
        """
        hwp = reader.hwp
        if hwp is None:
            return False
        s = (find or "").strip()
        if not s:
            return False
        try:
            with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
                fr = hwp.HParameterSet.HFindReplace
                try:
                    fr.FindString = s
                except Exception:
                    try:
                        fr.HSet.SetItem("FindString", s)
                    except Exception:
                        pass
                try:
                    fr.IgnoreMessage = 1
                except Exception:
                    try:
                        fr.HSet.SetItem("IgnoreMessage", 1)
                    except Exception:
                        pass
                try:
                    fr.HSet.SetItem("FindType", int(find_type))
                except Exception:
                    pass
                r = hwp.HAction.Execute("RepeatFind", fr.HSet)
            ok = (r == 1)
        except Exception:
            ok = False
        # 선택 해제(ESC)는 표 컨트롤 선택으로 승격될 수 있어,
        # 가능한 경우 선택을 "조용히 접기"로 정리합니다.
        if ok:
            try:
                hwp.HAction.GetDefault("MoveSelEnd", hwp.HParameterSet.HSelectionOpt.HSet)
                hwp.HAction.Execute("MoveSelEnd", hwp.HParameterSet.HSelectionOpt.HSet)
            except Exception:
                pass
            try:
                hwp.HAction.GetDefault("MoveRight", hwp.HParameterSet.HSelectionOpt.HSet)
                hwp.HAction.Execute("MoveRight", hwp.HParameterSet.HSelectionOpt.HSet)
            except Exception:
                pass
        return ok

    def _try_all_replace(self, reader: HWPReader, *, find: str, replace: str, find_type: int = 1) -> bool:
        """
        HWP '찾아바꾸기(모두 바꾸기)'를 실행합니다.
        - 사용자 매크로: AllReplace + HFindReplace
        - 표 셀 내부 텍스트도 '문자열'만 안전 치환 (Cut 금지)
        """
        hwp = reader.hwp
        if hwp is None:
            return False
        f = (find or "").strip()
        if not f:
            return False

        def _set(fr, name: str, value) -> None:
            try:
                setattr(fr, name, value)
                return
            except Exception:
                pass
            try:
                fr.HSet.SetItem(name, value)
            except Exception:
                pass

        try:
            with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                # AllReplace가 없는 환경은 ExecReplace로 폴백
                action = "AllReplace"
                try:
                    hwp.HAction.GetDefault(action, hwp.HParameterSet.HFindReplace.HSet)
                except Exception:
                    action = "ExecReplace"
                    hwp.HAction.GetDefault(action, hwp.HParameterSet.HFindReplace.HSet)

                fr = hwp.HParameterSet.HFindReplace
                _set(fr, "MatchCase", 0)
                _set(fr, "AllWordForms", 0)
                _set(fr, "SeveralWords", 0)
                _set(fr, "UseWildCards", 0)
                _set(fr, "WholeWordOnly", 0)
                _set(fr, "AutoSpell", 1)
                try:
                    _set(fr, "Direction", fr.FindDir("Forward"))
                except Exception:
                    pass
                _set(fr, "IgnoreFindString", 0)
                _set(fr, "IgnoreReplaceString", 0)
                _set(fr, "ReplaceMode", 1)
                _set(fr, "IgnoreMessage", 1)
                _set(fr, "HanjaFromHangul", 0)
                _set(fr, "FindJaso", 0)
                _set(fr, "FindRegExp", 0)
                _set(fr, "FindStyle", "")
                _set(fr, "ReplaceStyle", "")
                _set(fr, "FindType", int(find_type))
                _set(fr, "FindString", str(f))
                _set(fr, "ReplaceString", str(replace or ""))

                r = hwp.HAction.Execute(action, hwp.HParameterSet.HFindReplace.HSet)
            return True if r is None else bool(r)
        except Exception:
            return False

    def _fill_header_markers_via_macro(self, reader: HWPReader, *, date_str: str, title: str, teacher: str, scope_text: str) -> None:
        """
        머리말 표 셀 내부의 HDR_* 텍스트를 AllReplace로 치환합니다.
        (HeaderFooter 진입 필수)
        """
        cfg = self._diag_config()

        # ✅ HeaderFooter(머리말 편집 진입) 자체가 "빈 양쪽 머리말 컨트롤"을 추가 생성하는 케이스가 있어,
        #    머리말 치환은 HeaderFooter 없이 문서 전체 스코프(RepeatFind) 기반 1회 치환으로 수행합니다.
        #    (템플릿 머리말 표가 유지되는 상태에서, 추가 빈 머리말 생성만 차단하는 목적)
        hwp = reader.hwp
        if hwp is None:
            return

        replaced_any = False
        for token, val in [
            ("HDR_DATE", str(date_str)),
            ("HDR_TITLE", (title or "").strip()),
            ("HDR_TEACHER", (teacher or "").strip()),
            ("HDR_SCOPE", str(scope_text)),
        ]:
            if self._replace_first_by_repeat_find(reader, token, val):
                replaced_any = True
                self._diag_pause(cfg, f"HEADER RepeatFind+Insert replaced: {token}")

        # 치환을 한 번이라도 했다면, 본문 포커스 복귀를 시도(저장/붙여넣기 안전)
        if replaced_any:
            try:
                reader.ensure_main_body_focus()
            except Exception:
                pass

    def _remove_body_marker_and_anchor(self, reader: HWPReader, marker: str) -> bool:
        """
        본문에서 marker(예: PROBLEMS_HERE)를 찾아 제거하고,
        제거된 위치를 그대로 "삽입 시작점"으로 사용하도록 커서를 그 위치로 복귀합니다.
        - Cut 금지(표/컨트롤 삭제 방지)
        """
        hwp = reader.hwp
        if hwp is None:
            return False
        mk = (marker or "").strip()
        if not mk:
            return False

        # 본문 포커스 강제 (머리말/주석 편집 상태 방지)
        try:
            self._try_close_subedit(reader)
        except Exception:
            pass
        try:
            reader.ensure_main_body_focus()
        except Exception:
            pass

        # 문서 처음부터 marker 선택
        anchor = None
        try:
            hwp.HAction.GetDefault("MoveDocBegin", hwp.HParameterSet.HSelectionOpt.HSet)
            hwp.HAction.Execute("MoveDocBegin", hwp.HParameterSet.HSelectionOpt.HSet)
        except Exception:
            pass

        # RepeatFind로 첫 번째를 찾고 시작 좌표를 저장
        found = False
        try:
            hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            try:
                hwp.HParameterSet.HFindReplace.FindString = mk
            except Exception:
                try:
                    hwp.HParameterSet.HFindReplace.HSet.SetItem("FindString", mk)
                except Exception:
                    pass
            try:
                hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            except Exception:
                try:
                    hwp.HParameterSet.HFindReplace.HSet.SetItem("IgnoreMessage", 1)
                except Exception:
                    pass
            # 본문 컨텍스트 한정
            try:
                hwp.HParameterSet.HFindReplace.HSet.SetItem("FindType", 1)
            except Exception:
                pass
            with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                r = hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            found = (r == 1)
            if found:
                try:
                    hwp.HAction.GetDefault("MoveSelBegin", hwp.HParameterSet.HSelectionOpt.HSet)
                    hwp.HAction.Execute("MoveSelBegin", hwp.HParameterSet.HSelectionOpt.HSet)
                except Exception:
                    pass
                try:
                    anchor = hwp.GetPos()
                except Exception:
                    anchor = None
        except Exception:
            found = False

        try:
            self._safe_run(hwp, "Cancel")
        except Exception:
            pass

        if not found:
            return False

        # 문서 전체(본문)에서 marker 제거
        try:
            hwp.HAction.GetDefault("MoveDocBegin", hwp.HParameterSet.HSelectionOpt.HSet)
            hwp.HAction.Execute("MoveDocBegin", hwp.HParameterSet.HSelectionOpt.HSet)
        except Exception:
            pass
        self._try_all_replace(reader, find=mk, replace="", find_type=1)

        # 커서 복귀(삽입 시작점)
        if anchor is not None:
            try:
                sec, para, pos = anchor
                hwp.SetPos(sec, para, pos)
            except Exception:
                # SetPos 실패 시에는 현재 위치에 삽입
                pass
        return True

    def _try_close_subedit(self, reader: HWPReader) -> None:
        """
        머리말/각주/미주 등 '서브 편집 영역'에 들어가 있는 경우를 대비해,
        가능한 범위에서 Shift+Esc(=CloseEx) 동작으로 본문으로 복귀를 시도합니다.

        - 문서 자체를 닫는 Close/FileClose는 호출하지 않습니다.
        """
        hwp = reader.hwp
        if hwp is None:
            return
        try:
            with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                try:
                    self._safe_run(hwp, "Cancel")
                except Exception:
                    pass
                try:
                    self._safe_run(hwp, "CloseEx")
                except Exception:
                    pass
        except Exception:
            return

    def _candidate_open_paths(self, path: str) -> List[str]:
        abs_path = os.path.abspath(path)
        cands = [abs_path]
        # short path (8.3) 폴백 (가능한 환경에서만)
        try:
            import win32api  # type: ignore

            sp = win32api.GetShortPathName(abs_path)
            if sp and sp not in cands:
                cands.append(sp)
        except Exception:
            pass
        return cands

    def _open_any(self, reader: HWPReader, file_path: str) -> None:
        """
        XHwpDocuments.Open 실패 시 HAction.FileOpen으로 폴백합니다.
        HWP 팝업 억제 보호막도 함께 적용합니다.
        """
        hwp = reader.hwp
        if hwp is None:
            raise WorksheetComposeError("한글(HWP) 프로그램을 초기화할 수 없습니다.")

        last_err: Optional[Exception] = None
        for p in self._candidate_open_paths(file_path):
            # 1) XHwpDocuments.Open
            try:
                with reader._auto_close_hwp_popups(timeout_sec=6.0), reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                    hwp.XHwpDocuments.Open(p)
                # HWPReader.find_text()는 is_opened에 의존하므로 상태를 동기화합니다.
                reader.is_opened = True
                return
            except Exception as e:
                last_err = e

            # 2) HAction.FileOpen
            try:
                with reader._auto_close_hwp_popups(timeout_sec=6.0), reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                    hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
                    hwp.HParameterSet.HFileOpenSave.filename = p
                    hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
                # HWPReader.find_text()는 is_opened에 의존하므로 상태를 동기화합니다.
                reader.is_opened = True
                return
            except Exception as e:
                last_err = e

        raise WorksheetComposeError(f"문제 HWP 열기 실패: {last_err}")

    def compose(
        self,
        *,
        problem_ids: List[str],
        output_path: str,
        # 템플릿/머리말 채우기 옵션
        template_path: Optional[str] = None,
        title: str = "",
        teacher: str = "",
        date_str: Optional[str] = None,
        # 문항 아래 태그 삽입: {"unit": bool, "source": bool, "difficulty": bool}
        tag_options: Optional[dict] = None,
        # 문항 순서와 동일: [{"problem_id": str, "unit": str, "source": str, "difficulty": str}, ...]
        problem_meta: Optional[List[dict]] = None,
    ) -> str:
        if not self.db.is_connected():
            raise WorksheetComposeError(
                "DB에 연결되지 않았습니다. HWP를 생성할 수 없습니다."
            )
        if not problem_ids:
            raise WorksheetComposeError("문항이 없습니다.")
        if not output_path:
            raise WorksheetComposeError("저장 경로가 비어있습니다.")

        self._template_missing = False
        restored: List[Tuple[str, str]] = []
        try:
            restored = self._restore_all(problem_ids)
            if not restored:
                raise WorksheetComposeError("원본 HWP를 복원할 수 없습니다.")

            with HWPReader() as reader:
                cfg = self._diag_config()
                self._set_hwp_visible(reader, bool(cfg.get("visible")))
                hwp = reader.hwp
                if hwp is None:
                    raise WorksheetComposeError("한글(HWP) 프로그램을 초기화할 수 없습니다.")

                # 템플릿 열기(있으면 템플릿 기반), 없으면 기존 방식(FileNew)
                tpl = template_path
                if not tpl:
                    tpl = self._resolve_default_template_path()
                if tpl and os.path.exists(tpl):
                    try:
                        self._open_any(reader, tpl)
                    except Exception as e:
                        raise WorksheetComposeError(f"템플릿 HWP 열기 실패: {e}\n- path: {tpl}")
                else:
                    self._template_missing = True
                    base_hint = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else "프로젝트 루트"
                    logger.warning(
                        "템플릿을 찾지 못해 빈 문서로 생성합니다. "
                        "학습지/오답노트 서식을 쓰려면 exe와 같은 폴더에 templates 폴더(worksheet_template.hwp 포함)를 두세요. "
                        "참고: exe 기준 폴더=%s",
                        base_hint,
                    )
                    try:
                        hwp.HAction.GetDefault("FileNew", hwp.HParameterSet.HFileOpenSave.HSet)
                        hwp.HAction.Execute("FileNew", hwp.HParameterSet.HFileOpenSave.HSet)
                        reader.is_opened = True
                    except Exception as e:
                        raise WorksheetComposeError(f"HWP 새 문서 생성 실패: {e}")

                try:
                    output_doc = hwp.XHwpDocuments.Active_XHwpDocument
                except Exception:
                    output_doc = None

                # 머리말(필드/북마크) 채우기: 템플릿이 있는 경우만 시도
                if tpl and os.path.exists(tpl):
                    # 날짜/범위 계산
                    if not date_str:
                        date_str = datetime.now().strftime("%Y.%m.%d")
                    scope_text = self._compute_scope_text(problem_ids)

                    # 1) PutFieldText 시도(누름틀/필드 기반 템플릿인 경우)
                    #    일부 환경에서는 예외 없이 "성공처럼" 반환되지만 실제 반영이 안 되는 케이스가 있어,
                    #    2) 텍스트 마커(HDR_*) 치환도 항상 추가로 시도합니다.
                    self._try_put_field_text(hwp, "HDR_DATE", date_str)
                    self._try_put_field_text(hwp, "HDR_TITLE", (title or "").strip())
                    self._try_put_field_text(hwp, "HDR_TEACHER", (teacher or "").strip())
                    self._try_put_field_text(hwp, "HDR_SCOPE", scope_text)
                    # 2) 머리말(표 셀) 내부 HDR_*는 HeaderFooter 진입 후 AllReplace로 치환
                    self._fill_header_markers_via_macro(
                        reader,
                        date_str=str(date_str),
                        title=(title or "").strip(),
                        teacher=(teacher or "").strip(),
                        scope_text=str(scope_text),
                    )
                    self._diag_pause(cfg, "AFTER header fill")

                    # 진단 모드: header_only면 여기서 저장하고 종료
                    if (cfg.get("mode") or "normal") == "header_only":
                        self._diag_pause(cfg, "HEADER_ONLY checkpoint before save")
                        # 저장 후 바로 반환
                        hwp.HAction.GetDefault("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
                        hwp.HParameterSet.HFileOpenSave.filename = output_path
                        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
                        with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                            hwp.HAction.Execute("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
                        return output_path

                # 문제 삽입 위치로 이동(템플릿) 또는 문서 끝(기존)
                if tpl and os.path.exists(tpl):
                    # 머리말/주석 등 편집 영역에 들어간 상태면 본문으로 복귀 시도
                    self._try_close_subedit(reader)
                    # PROBLEMS_HERE(본문) 제거 + 그 위치를 삽입 시작점으로 사용
                    ok = self._remove_body_marker_and_anchor(reader, "PROBLEMS_HERE")
                    if not ok:
                        self._safe_run(hwp, "MoveDocEnd")
                else:
                    try:
                        self._safe_run(hwp, "MoveDocEnd")
                    except Exception:
                        pass

                # 진단 모드: body_only면 1문항만 붙이고 종료
                if (cfg.get("mode") or "normal") in ("body_only", "combined"):
                    restored_to_use = restored[:1]
                else:
                    restored_to_use = restored

                # 문제 문서들을 순서대로 삽입
                for pid, pth in restored_to_use:
                    if not os.path.exists(pth):
                        continue
                    try:
                        if os.path.getsize(pth) <= 0:
                            raise WorksheetComposeError(f"원본 HWP가 비어있습니다. (problem_id={pid})")
                    except OSError:
                        pass

                    # 문제 문서 열기
                    try:
                        self._open_any(reader, pth)
                        prob_doc = hwp.XHwpDocuments.Active_XHwpDocument
                    except Exception as e:
                        raise WorksheetComposeError(f"문제 HWP 열기 실패: {e} (problem_id={pid}, path={pth})")

                    # 문제 문서에서 본문 포커스로 이동 → Select All이 본문 전체를 선택하도록 (미주/머리말에 포커스 있으면 깨짐 방지)
                    try:
                        self._try_close_subedit(reader)
                    except Exception:
                        pass
                    try:
                        reader.ensure_main_body_focus()
                    except Exception:
                        pass

                    # 전체 선택 + 복사
                    self._safe_run(hwp, "SelectAll")
                    self._safe_run(hwp, "Copy")

                    # 문제 문서 닫기(저장 질문 없이)
                    try:
                        prob_doc.Close(False)
                    except Exception:
                        try:
                            prob_doc.Close(isDirty=False)
                        except Exception:
                            pass

                    # 출력 문서로 돌아와 붙여넣기
                    try:
                        if output_doc is not None:
                            output_doc.Activate()
                    except Exception:
                        pass

                    # Paste 직전: 본문 포커스 강제(머리말/주석 편집 상태 방지)
                    try:
                        self._try_close_subedit(reader)
                    except Exception:
                        pass
                    try:
                        reader.ensure_main_body_focus()
                    except Exception:
                        pass
                    try:
                        self._safe_run(hwp, "Cancel")
                    except Exception:
                        pass

                    self._safe_run(hwp, "Paste")

                    # 문항 아래 태그 삽입(옵션에 따라 단원/출처/난이도)
                    opts = tag_options or {}
                    meta = next((m for m in (problem_meta or []) if str(m.get("problem_id")) == str(pid)), None)
                    if meta and any((opts.get("unit"), opts.get("source"), opts.get("difficulty"))):
                        parts = []
                        if opts.get("unit") and (meta.get("unit") or "").strip():
                            parts.append(f"[단원] {(meta.get('unit') or '').strip()}")
                        if opts.get("source") and (meta.get("source") or "").strip():
                            parts.append(f"[출처] {(meta.get('source') or '').strip()}")
                        if opts.get("difficulty") and (meta.get("difficulty") or "").strip():
                            parts.append(f"[난이도] {(meta.get('difficulty') or '').strip()}")
                        if parts:
                            for line in parts:
                                self._paste_text(hwp, line)
                                self._safe_run(hwp, "BreakPara")

                    # 문제 사이 줄바꿈(가능하면)
                    self._safe_run(hwp, "BreakPara")
                    self._safe_run(hwp, "BreakPara")

                # 진단 모드: body_only는 여기서 저장하고 종료
                if (cfg.get("mode") or "normal") == "body_only":
                    self._diag_pause(cfg, "BODY_ONLY checkpoint before save")
                    hwp.HAction.GetDefault("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
                    hwp.HParameterSet.HFileOpenSave.filename = output_path
                    hwp.HParameterSet.HFileOpenSave.Format = "HWP"
                    with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                        hwp.HAction.Execute("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
                    return output_path

                # 저장
                try:
                    # 혹시 남아있는 PROBLEMS_HERE가 있으면 제거(본문 전체)
                    try:
                        self._try_close_subedit(reader)
                        self._safe_run(hwp, "MoveDocBegin")
                    except Exception:
                        pass
                    self._try_all_replace(reader, find="PROBLEMS_HERE", replace="", find_type=1)

                    hwp.HAction.GetDefault("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
                    hwp.HParameterSet.HFileOpenSave.filename = output_path
                    hwp.HParameterSet.HFileOpenSave.Format = "HWP"
                    # 저장/확인 팝업 방지(가능한 범위)
                    try:
                        with reader._temp_message_box_mode(0x20021):  # type: ignore[attr-defined]
                            hwp.HAction.Execute("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
                    except Exception:
                        hwp.HAction.Execute("FileSaveAs", hwp.HParameterSet.HFileOpenSave.HSet)
                except Exception as e:
                    raise WorksheetComposeError(f"HWP 저장 실패: {e}")

            if not os.path.exists(output_path):
                raise WorksheetComposeError("HWP 저장에 실패했습니다. (파일이 생성되지 않음)")
            return output_path
        finally:
            # 임시 문제 파일 정리
            for _pid, p in restored:
                try:
                    if p and os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass

