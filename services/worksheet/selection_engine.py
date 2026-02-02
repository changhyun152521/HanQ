"""
학습지 문제 선택 엔진(순수 로직)

요구사항(확정):
- 단원(과목/대단원/소단원) 다중 선택
- 총 문항수 N을 단원 수 U로 균등 배분 (N//U, 나머지는 앞 단원부터 +1)
- 각 단원에서 난이도 비율(킬/상/중/하)을 동일하게 적용
- 특정 단원/난이도에서 부족하면 "보충 없이" 그대로 감산(총 문항 감소 허용)
- 문제는 중복(problem_id 중복) 허용하지 않음
- 정렬 옵션: 랜덤(단원/난이도 체크 불가) 또는 단원순/난이도순(둘 다 체크 시 2단 정렬)
- 랜덤은 "출력 순서"만 랜덤이 아니라, 최종 출력 목록 자체를 완전 셔플
  (단원도 섞이며, 선택된 단원/출처 범위 내에서만 섞임)

엔진 입력은 DB/서비스에서 미리 필터링된 Problem 리스트(후보)로 받는 것을 권장하지만,
안전하게 필수 조건(content_raw + 대표 태그 완전성) 검사 유틸도 제공합니다.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
import math
import random

from core.models import Problem, Tag
from core.unit_catalog import UNIT_CATALOG


# 난이도 정렬/표시 기준(요구사항): 하 -> 중 -> 상 -> 킬
DIFFICULTY_ORDER: Tuple[str, ...] = ("하", "중", "상", "킬")


@dataclass(frozen=True)
class UnitKey:
    """단원 식별 키(필수: subject/major/sub)."""

    subject: str
    major_unit: str
    sub_unit: str

    def normalized(self) -> "UnitKey":
        return UnitKey(
            subject=(self.subject or "").strip(),
            major_unit=(self.major_unit or "").strip(),
            sub_unit=(self.sub_unit or "").strip(),
        )

    def is_valid(self) -> bool:
        u = self.normalized()
        return bool(u.subject and u.major_unit and u.sub_unit)


DifficultyRatios = Dict[str, int]  # {"킬": 30, "상": 20, "중": 20, "하": 30}


@dataclass(frozen=True)
class OrderOptions:
    """
    출력 정렬 옵션
    - randomize: True면 전체 결과를 완전 셔플(단원/난이도 체크 불가)
    - order_by_unit: 단원 순서(카탈로그/선택 순 기반)는 상위 키
    - order_by_difficulty: 난이도 순서(킬/상/중/하)
    """

    randomize: bool = False
    order_by_unit: bool = True
    order_by_difficulty: bool = True


@dataclass(frozen=True)
class WorksheetSelectionSpec:
    """선택 엔진 입력 스펙(순수 로직)."""

    units: List[UnitKey]  # 선택 단원(정렬 기준이 되므로 순서 유지)
    total_count: int
    difficulty_ratios: DifficultyRatios
    order: OrderOptions
    seed: Optional[int] = None  # randomize/추출 랜덤에 사용 (None이면 완전 랜덤)


@dataclass
class WorksheetSelectionResult:
    """
    선택 결과
    - selected_problem_ids: 최종 선택된 문제 id 리스트(출력 순서 반영)
    - per_unit_selected: 단원별 선택 수
    - requested_total: 사용자가 요청한 총 문항수
    - actual_total: 실제 생성된 문항수(부족 시 감소)
    - warnings: 사용자 안내용 경고 메시지
    """

    selected_problem_ids: List[str]
    per_unit_selected: Dict[UnitKey, int]
    requested_total: int
    actual_total: int
    warnings: List[str]


class WorksheetSelectionEngine:
    """단원/난이도 비율 기반 문제 선택 엔진."""

    def _catalog_unit_sort_key(self, u: UnitKey) -> Tuple[int, int, int]:
        """
        unit_catalog.py의 삽입 순서를 "단원순서"로 사용합니다.

        반환 키: (subject_index, major_index, sub_index)
        - 카탈로그에 없는 값은 큰 값으로 밀어 뒤로 보냄
        """
        u = u.normalized()
        big = 10**9

        try:
            subjects = list(UNIT_CATALOG.keys())
            s_idx = subjects.index(u.subject) if u.subject in UNIT_CATALOG else big
        except Exception:
            s_idx = big

        try:
            majors_map = UNIT_CATALOG.get(u.subject, {})
            majors = list(majors_map.keys())
            m_idx = majors.index(u.major_unit) if u.major_unit in majors_map else big
        except Exception:
            m_idx = big

        try:
            subs = majors_map.get(u.major_unit, [])
            sub_idx = subs.index(u.sub_unit) if u.sub_unit in subs else big
        except Exception:
            sub_idx = big

        return (s_idx, m_idx, sub_idx)

    def validate_spec(self, spec: WorksheetSelectionSpec) -> None:
        if spec.total_count <= 0:
            raise ValueError("총 문항수는 1 이상이어야 합니다.")
        if not spec.units:
            raise ValueError("단원을 1개 이상 선택해야 합니다.")

        # 단원 유효성(소단원까지 필수)
        bad_units = [u for u in spec.units if not u.is_valid()]
        if bad_units:
            raise ValueError("단원 선택이 완전하지 않습니다(과목/대단원/소단원 모두 필요).")

        # 난이도 비율: 합계 100 강제 + 허용 키만
        ratios = spec.difficulty_ratios or {}
        if any(k not in DIFFICULTY_ORDER for k in ratios.keys()):
            raise ValueError("난이도 비율 키는 '킬/상/중/하'만 허용됩니다.")
        total_ratio = sum(int(v) for v in ratios.values())
        if total_ratio != 100:
            raise ValueError("난이도 비율 합계는 100%여야 합니다.")
        # 모든 난이도 키 존재 강제(입력 UI 단순화)
        for k in DIFFICULTY_ORDER:
            if k not in ratios:
                raise ValueError("난이도 비율은 킬/상/중/하 모두 입력해야 합니다.")

        # 랜덤이면 다른 체크는 의미 없으니 강제(추후 UI에서 비활성 처리하지만 로직도 보호)
        if spec.order.randomize and (spec.order.order_by_unit or spec.order.order_by_difficulty):
            # order_by_* 값은 무시되지만, 명확히 하려고 예외 대신 강제 무시 경고는 결과에 남김.
            pass

    def _normalize_units(self, units: Sequence[UnitKey]) -> List[UnitKey]:
        """
        단원 키 정규화 + 중복 제거(순서 유지).
        UI/입력에서 공백이나 중복 선택이 들어와도 안정적으로 동작하게 합니다.
        """
        kept: List[Tuple[int, UnitKey]] = []
        seen: set[UnitKey] = set()
        for idx, u in enumerate(units):
            nu = u.normalized()
            if not nu.is_valid():
                continue
            if nu in seen:
                continue
            seen.add(nu)
            kept.append((idx, nu))

        # "단원순서"는 카탈로그 기준으로 정렬(동일 키는 입력 순서 유지)
        kept.sort(key=lambda t: (self._catalog_unit_sort_key(t[1]), t[0]))
        return [u for _idx, u in kept]

    # ----------------------------
    # 후보 필터링(안전 장치)
    # ----------------------------
    def _primary_tag(self, p: Problem) -> Optional[Tag]:
        if not p.tags:
            return None
        return p.tags[0]

    def is_problem_usable(self, p: Problem) -> bool:
        """
        사용 가능 문제 조건(확정):
        - content_raw_file_id 존재
        - 대표 태그(0번)에 subject/major_unit/sub_unit/difficulty 모두 존재
        """
        if not p or not p.id:
            return False
        if not p.content_raw_file_id:
            return False
        t = self._primary_tag(p)
        if not t:
            return False
        subject = (t.subject or "").strip()
        major = (getattr(t, "major_unit", None) or "").strip()
        sub = (getattr(t, "sub_unit", None) or "").strip()
        diff = (getattr(t, "difficulty", None) or "").strip()
        return bool(subject and major and sub and diff)

    def problem_unit_key(self, p: Problem) -> Optional[UnitKey]:
        t = self._primary_tag(p)
        if not t:
            return None
        return UnitKey(
            subject=(t.subject or "").strip(),
            major_unit=(getattr(t, "major_unit", None) or "").strip(),
            sub_unit=(getattr(t, "sub_unit", None) or "").strip(),
        )

    def problem_difficulty(self, p: Problem) -> Optional[str]:
        t = self._primary_tag(p)
        if not t:
            return None
        d = (getattr(t, "difficulty", None) or "").strip()
        return d or None

    # ----------------------------
    # 배분 로직
    # ----------------------------
    def _split_even(self, total: int, unit_count: int) -> List[int]:
        """총 문항수 total을 unit_count로 균등 분배: 앞에서부터 +1."""
        q, r = divmod(int(total), int(unit_count))
        out = [q] * unit_count
        for i in range(r):
            out[i] += 1
        return out

    def _difficulty_allocation(self, total: int, ratios: DifficultyRatios) -> Dict[str, int]:
        """
        비율 → 정수 문항수 변환:
        - 기본 내림(floor)
        - 남는 문항은 소수점(잔여)이 큰 순서로 배분
        """
        total = int(total)
        if total <= 0:
            return {k: 0 for k in DIFFICULTY_ORDER}

        raw = []
        for k in DIFFICULTY_ORDER:
            r = int(ratios.get(k, 0))
            x = total * (r / 100.0)
            base = int(math.floor(x))
            frac = x - base
            raw.append((k, base, frac))

        allocated = {k: base for (k, base, _frac) in raw}
        used = sum(allocated.values())
        remain = total - used
        if remain > 0:
            # frac 큰 순, 동률이면 난이도 order 유지
            order_index = {k: i for i, k in enumerate(DIFFICULTY_ORDER)}
            raw_sorted = sorted(raw, key=lambda t: (-t[2], order_index[t[0]]))
            for i in range(remain):
                allocated[raw_sorted[i % len(raw_sorted)][0]] += 1

        return allocated

    # ----------------------------
    # 메인 선택
    # ----------------------------
    def select(self, spec: WorksheetSelectionSpec, candidates: Sequence[Problem]) -> WorksheetSelectionResult:
        """
        candidates: 선택된 출처(Textbook/Exam)들에서 모은 후보 문제들
        """
        self.validate_spec(spec)

        rng = random.Random(spec.seed) if spec.seed is not None else random.Random()

        # 단원 키 정규화/중복 제거(순서 유지)
        units = self._normalize_units(spec.units)
        if not units:
            raise ValueError("단원을 1개 이상 선택해야 합니다.")

        # 1) 후보 문제 유효성 필터(안전)
        usable: List[Problem] = [p for p in candidates if self.is_problem_usable(p)]

        # 2) 단원별 풀 구성
        #    - dedupe는 최종 선택 과정에서 적용
        unit_to_pool: Dict[UnitKey, Dict[str, List[Problem]]] = {}
        for u in units:
            unit_to_pool[u] = {k: [] for k in DIFFICULTY_ORDER}

        for p in usable:
            uk = self.problem_unit_key(p)
            if not uk:
                continue
            # 선택 단원과 정확히 일치해야 함(소단원까지 필수)
            # spec.units의 UnitKey는 이미 정규화/유효성 검증됨
            # uk도 정규화
            uk = uk.normalized()
            if uk not in unit_to_pool:
                continue
            d = self.problem_difficulty(p)
            if d not in DIFFICULTY_ORDER:
                continue
            unit_to_pool[uk][d].append(p)

        # 3) 단원별 균등 배분
        allocations = self._split_even(spec.total_count, len(units))

        selected_ids: List[str] = []
        per_unit_selected: Dict[UnitKey, int] = {}
        warnings: List[str] = []

        used_id_set = set()

        for idx, unit in enumerate(units):
            target_for_unit = allocations[idx]
            if target_for_unit <= 0:
                per_unit_selected[unit] = 0
                continue

            # 단원 내 난이도 배분
            per_diff_target = self._difficulty_allocation(target_for_unit, spec.difficulty_ratios)

            picked_for_unit: List[Problem] = []
            for diff in DIFFICULTY_ORDER:
                need = int(per_diff_target.get(diff, 0))
                if need <= 0:
                    continue

                pool = [p for p in unit_to_pool[unit][diff] if p.id not in used_id_set]
                if not pool:
                    continue

                if len(pool) <= need:
                    chosen = pool
                else:
                    chosen = rng.sample(pool, need)

                for p in chosen:
                    if not p.id or p.id in used_id_set:
                        continue
                    used_id_set.add(p.id)
                    picked_for_unit.append(p)

            # 부족하면 그대로 감산(보충 없음)
            per_unit_selected[unit] = len(picked_for_unit)
            selected_ids.extend([p.id for p in picked_for_unit if p.id])

            if len(picked_for_unit) < target_for_unit:
                warnings.append(
                    f"단원 '{unit.subject} / {unit.major_unit} > {unit.sub_unit}'에서 "
                    f"{target_for_unit}문항 중 {len(picked_for_unit)}문항만 선택되었습니다(부족분 감산)."
                )

        # 4) 출력 순서 결정
        # 랜덤이면 완전 셔플(단원도 섞임)
        if spec.order.randomize:
            rng.shuffle(selected_ids)
        else:
            # 정렬은 선택된 id에 대응하는 Problem을 찾아 key 기반으로 정렬
            # (후보 리스트에 존재하는 문제만 정렬 가능)
            id_to_problem: Dict[str, Problem] = {}
            for p in usable:
                if p.id:
                    id_to_problem[p.id] = p

            unit_index = {u: i for i, u in enumerate(units)}

            def unit_sort_key(unit: UnitKey) -> int:
                # 정규화+카탈로그 정렬된 units 순서를 기준으로 정렬
                return unit_index.get(unit, 10**9)

            diff_rank = {d: i for i, d in enumerate(DIFFICULTY_ORDER)}

            def key(pid: str) -> Tuple:
                p = id_to_problem.get(pid)
                if not p:
                    return (10**9, 10**9, 10**9)
                uk = self.problem_unit_key(p)
                d = self.problem_difficulty(p) or ""
                # 기본 tie-breaker: source_type, source_id, problem_index
                base = getattr(p, "problem_index", 0) or 0
                if spec.order.order_by_unit and uk:
                    ukey = unit_sort_key(uk.normalized())
                else:
                    ukey = 0
                if spec.order.order_by_difficulty:
                    dkey = diff_rank.get(d, 10**9)
                else:
                    dkey = 0
                return (ukey, dkey, base)

            selected_ids.sort(key=key)

        # 중복 제거 안전(이미 used_id_set으로 막지만, 혹시라도)
        final_ids: List[str] = []
        for pid in selected_ids:
            if pid in used_id_set and pid not in final_ids:
                final_ids.append(pid)

        return WorksheetSelectionResult(
            selected_problem_ids=final_ids,
            per_unit_selected=per_unit_selected,
            requested_total=int(spec.total_count),
            actual_total=len(final_ids),
            warnings=warnings,
        )

