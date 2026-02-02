# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any, List, Tuple

from pyhwpx import Hwp

Pos = Tuple[int, int, int]  # (List, Para, Pos)

EN_BODY = 14      # 미주 본문으로 진입
LST_END = 5       # 현재 리스트 끝(미주 본문 끝)
LST_BEG = 4       # 현재 리스트 시작(미주 본문 시작)

# ─────────────────────────────────────────────────────────────
def run(hwp: Any, cmd: str) -> None:
    fn = getattr(hwp, cmd, None)
    if callable(fn):
        fn()
    else:
        hwp.Run(cmd)

def mps(hwp: Any, kind: int) -> None:
    if hasattr(hwp, "move_pos"):
        hwp.move_pos(kind)
    elif hasattr(hwp, "MovePos"):
        hwp.MovePos(kind, 0, 0)
    else:
        raise RuntimeError("MovePos/move_pos not available")

def sps(hwp: Any, p: Pos) -> None:
    if hasattr(hwp, "set_pos"):
        hwp.set_pos(*p)
    else:
        hwp.SetPos(*p)

def gps(hwp: Any) -> Pos:
    if hasattr(hwp, "get_pos"):
        return tuple(hwp.get_pos())
    return tuple(hwp.GetPos())

def spb(hwp: Any, posset: Any) -> None:
    hwp.SetPosBySet(posset)

# ─────────────────────────────────────────────────────────────
# (A) 빈줄 판정용 길이 측정: emp   by koshon
# ─────────────────────────────────────────────────────────────
def sln(hwp: Any) -> None:
    hwp.MoveLineBegin()
    hwp.MoveSelLineEnd()


def gln(hwp: Any) -> int:
    s = hwp.GetTextFile("HWP", "saveblock")
    if s is None:
        return 0
    return len(s)


def emp(hwp: Any) -> int:
    """
    '빈 줄' 길이 측정: 문서 끝에 공백 넣고 선택 길이 측정 후 원복
    """
    act = hwp.HAction
    p = hwp.HParameterSet.HInsertText
    hs = p.HSet
    act.GetDefault("InsertText", hs)

    run(hwp, "MoveTopLevelEnd")
    run(hwp, "BreakPara")
    run(hwp, "BreakPara")

    p.Text = "  "            # 공백 2개
    act.Execute("InsertText", hs)

    sln(hwp)
    n = gln(hwp)

    # 원복(원본처럼 3번 backspace)
    run(hwp, "DeleteBack")
    run(hwp, "DeleteBack")
    run(hwp, "DeleteBack")
    return n

def isb(hwp: Any, blank_len: int) -> bool:
    """
    현재 문단이 '빈 줄'인지 판정.
    1) hwp.is_empty_para() 가 있으면 먼저 사용 (모듈)
    2) 아니면 / 또는 그 외에, GetTextFile로 공백/개행만 있는 줄인지 확인
    """
    # 1) pyhwpx 쪽 is_empty_para 메서드가 있으면 최우선 사용
    if hasattr(hwp, "is_empty_para"):
        try:
            if hwp.is_empty_para():
                return True
        except Exception:
            # 혹시 예외 나도 밑으로 내려가서 텍스트 기반 판정
            pass

    # 2) 텍스트 기반 판정
    sln(hwp)  # 현재 줄 전체 선택
    s = hwp.GetTextFile("HWP", "saveblock") or ""

    # 눈에 보이는 문자가 하나도 없으면(공백/탭/개행만) 빈 줄로 처리
    if s.strip() == "":
        return True

    # 혹시 모를 예외 케이스에 대비해 기존 길이 기준도 남겨둠
    n = len(s)
    return (n == 0) or (n == blank_len)

def gpo(hwp: Any) -> int:
    """GetPosBySet().Item('Pos') 값."""
    ps = hwp.GetPosBySet()
    try:
        return int(ps.Item("Pos"))
    except Exception:
        return -1

# ─────────────────────────────────────────────────────────────
# (B) 미주 본문 앞/뒤 빈줄 제거: trb / tlb
# ─────────────────────────────────────────────────────────────
def trb(hwp: Any, blank_len: int) -> None:
    """endnote 끝부분 빈줄 제거"""
    while True:
        mps(hwp, LST_END)
        sln(hwp)

        # GetTextFile None/0이면 backspace
        if gln(hwp) == 0:
            run(hwp, "DeleteBack")
            continue

        if isb(hwp, blank_len):
            # blank면 backspace 2회
            run(hwp, "DeleteBack")
            run(hwp, "DeleteBack")
            continue

        run(hwp, "MoveLineEnd")
        if gpo(hwp) == 0:
            sln(hwp)
            run(hwp, "DeleteBack")
            run(hwp, "DeleteBack")
            continue

        break

def tlb(hwp: Any, blank_len: int) -> None:
    """endnote 시작부분 빈줄 제거"""
    while True:
        mps(hwp, LST_BEG)
        run(hwp, "MoveSelLineEnd")

        n = gln(hwp)
        if n == 0 or n == blank_len:
            run(hwp, "Delete")
            continue

        run(hwp, "MoveLineEnd")
        if gpo(hwp) == 0:
            sln(hwp)
            run(hwp, "Delete")
            run(hwp, "Delete")
            continue

        break

def cln(hwp: Any) -> int:
    """
    모든 endnote(en) 본문으로 들어가서 앞/뒤 빈줄 제거.
    """
    blank_len = emp(hwp)

    run(hwp, "MoveTopLevelBegin")
    cnt = 0

    c = hwp.HeadCtrl
    while c:
        try:
            if c.CtrlID == "en":              # 미주
                spb(hwp, c.GetAnchorPos(0))   # 미주 앵커로
                mps(hwp, EN_BODY)             # 본문 진입
                trb(hwp, blank_len)
                tlb(hwp, blank_len)
                cnt += 1
        except Exception:
            pass
        c = c.Next

    run(hwp, "MoveTopLevelBegin")
    return cnt

end_txt = "노블록"
# ─────────────────────────────────────────────────────────────
# (C) 본문 전체 텍스트 스캔: 연속 빈줄 삭제
# ─────────────────────────────────────────────────────────────
def 본문스캔(hwp: Any) -> None:
    """
    미주(Endnote) 기준으로 본문 빈줄 정리:
    - 미주 앵커 위치 a_i 를 이용해서,
    - 마지막 미주부터 역순으로 올라가며
      각 미주 바로 위쪽(이전 문단)의 *연속된 빈 문단*을 삭제하되
      최소 1줄은 남기기 위해 마지막에 BreakPara()로 한 줄 생성.
    - 문제 사이에 일부러 비워둔 빈 줄(미주와 떨어진 곳)은 건드리지 않는다.
    """
    # 1) '빈 줄 1줄'의 길이 측정
    blank_len = emp(hwp)

    # 2) 미주 앵커 리스트 수집
    anchors: List[Pos] = ena(hwp)
    if not anchors:
        return
    # '노블록'을 가상의 마지막 앵커로 추가
    nob_pos = nob(hwp, end_txt)
    anchors.append(nob_pos)          
    # 3) 마지막 미주부터 역순 처리
    for a in reversed(anchors):
        # 미주 앵커 위치로 이동
        sps(hwp, a)

        if hasattr(hwp, "move_pos"):
            moved = hwp.move_pos(11)   # movePrevPara
        else:
            moved = hwp.MovePos(11, 0, 0)

        if not moved:
            continue

        deleted_any = False

        while True:
            # 표 안이면 종료
            if hasattr(hwp, "is_cell") and hwp.is_cell():
                break

            if isb(hwp, blank_len):
                run(hwp, "DeleteBack")
                deleted_any = True
                continue
            else:
                break

        if deleted_any:
            run(hwp, "MoveParaEnd")
            hwp.BreakPara()

    run(hwp, "MoveTopLevelBegin")

# ─────────────────────────────────────────────────────────────
# (D) 블럭 추출/저장/삭제: save_block_as + find(end_txt)
# ─────────────────────────────────────────────────────────────
def ena(hwp: Any) -> List[Pos]:
    out: List[Pos] = []
    run(hwp, "MoveTopLevelBegin")
    c = hwp.HeadCtrl
    while c:
        try:
            if c.CtrlID == "en":
                spb(hwp, c.GetAnchorPos(0))
                out.append(gps(hwp))
        except Exception:
            pass
        c = c.Next
    return out

def nob(hwp: Any, word: str = end_txt) -> Pos:
    run(hwp, "MoveTopLevelBegin")
    try:
        ok = hwp.find(word)
    except Exception:
        ok = False

    if ok:
        for cmd in ("Cancel", "MovePrevWord"):
            try:
                run(hwp, cmd)
            except Exception:
                pass
        return gps(hwp)

    run(hwp, "MoveTopLevelEnd")
    return gps(hwp)

def ene(hwp: Any, a: List[Pos]) -> List[Pos]:
    """
    e_i = a_{i+1} 직전 (MoveLeft), 마지막 e_last = '노블록' 시작 위치.
    """
    if not a:
        return []
    e: List[Pos] = []
    for i in range(len(a) - 1):
        sps(hwp, a[i + 1])
        run(hwp, "MoveLeft")
        e.append(gps(hwp))
    e.append(nob(hwp, end_txt))
    return e

def sel(hwp: Any, a: Pos, e: Pos) -> None:
    sps(hwp, a)
    run(hwp, "Select")
    sps(hwp, e)

def sav(hwp: Any, out_path: str) -> None:
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    hwp.save_block_as(out_path, format="HWP")

def cut(hwp: Any, a: Pos, e: Pos, out_path: str) -> bool:
    if a == e:
        return False
    sel(hwp, a, e)
    sav(hwp, out_path)
    run(hwp, "Delete")
    return True


@dataclass
class Res:
    a: List[Pos]
    e: List[Pos]
    files: List[str]

def ext(hwp: Any, out_dir: str, fmt: str = "endnote_{i:03d}.hwp") -> Res:
    """
    미주 앵커 a_i 기준으로 [a_i ~ e_i] 블럭을 뒤에서부터 저장 + 삭제.
    """
    a = ena(hwp)
    e = ene(hwp, a)

    n = min(len(a), len(e))
    files: List[str] = []

    for i in range(n - 1, -1, -1):
        path = os.path.join(out_dir, fmt.format(i=i + 1))
        if cut(hwp, a[i], e[i], path):
            files.append(path)

    files.reverse()
    return Res(a=a, e=e, files=files)

def main():
    SRC = r"x:\out_blocks\endnote.hwp"   # 환경에 맞게 수정
    OUT = r"x:\out_blocks"

    hwp = Hwp()    
    hwp.open(SRC)

    본문스캔(hwp)
    cln(hwp)
    ext(hwp, out_dir=OUT, fmt="endnote_{i:03d}.hwp")


if __name__ == "__main__":
    main()
