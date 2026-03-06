"""
부서 배정 엔진 (Excel/구글폼 결과 정리용)

- 입력: pandas.DataFrame (지원자 1행 = 1명)
- 출력: 배정 결과 DF + 요약 DF + (옵션) 추첨 로그 DF

배정 규칙 (사용자 요구사항):
1) 모집 TO 대비 1순위 지원자가 적으면 1순위 전원 배정
2) 모집 TO 대비 1순위 지원자가 많으면 TO만큼 랜덤 배정
3) 미배정자 중 2순위로 지망하는 부서가 미달이면 남은 TO만큼 랜덤 배정
4) 미배정자 중 3순위로 지망하는 부서가 미달이면 남은 TO만큼 랜덤 배정
5) 끝까지 미배정자는 TO 남은 부서에 랜덤 배정

TO 규칙:
- 전체 인원 N을 6으로 나눈 몫을 각 부서 기본 TO로 배정
- 나머지 r명은 [기획부, 문화부, 연대사업부, 선전부, 교육부, 미디어부] 순서로 1명씩 추가 TO
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
import re
import pandas as pd
import random
from datetime import datetime, timezone

DEPARTMENTS: List[str] = ["기획부", "문화부", "연대사업부", "선전부", "교육부", "미디어부"]

# 유연한 입력(설문에서 줄여쓴 경우 등)을 정규화
SYNONYM_MAP = {
    "기획": "기획부",
    "기획부": "기획부",
    "문화": "문화부",
    "문화부": "문화부",
    "연대": "연대사업부",
    "연대사업": "연대사업부",
    "연대사업부": "연대사업부",
    "선전": "선전부",
    "선전부": "선전부",
    "교육": "교육부",
    "교육부": "교육부",
    "미디어": "미디어부",
    "미디어부": "미디어부",
}

def normalize_department(value) -> Optional[str]:
    """입력값을 부서명으로 정규화. 매칭 실패/공백은 None."""
    if value is None:
        return None
    s = str(value).strip()
    if s == "" or s.lower() in {"nan", "none", "null"}:
        return None
    # 공백/특수문자 제거(가벼운 정규화)
    s2 = re.sub(r"\s+", "", s)
    # 괄호 등 제거
    s2 = re.sub(r"[()\[\]{}]", "", s2)

    # 정확 매칭
    if s2 in SYNONYM_MAP:
        return SYNONYM_MAP[s2]

    # 포함 매칭(예: "연대사업부(대외)" 같이 들어오는 경우)
    for key, dept in SYNONYM_MAP.items():
        if key and key in s2:
            return dept

    return None

def calc_capacity(n_people: int, departments: List[str] = DEPARTMENTS) -> Dict[str, int]:
    """TO 규칙에 따라 부서별 정원을 계산."""
    base, r = divmod(int(n_people), len(departments))
    cap = {d: base for d in departments}
    for i in range(r):
        cap[departments[i]] += 1
    return cap

@dataclass
class AllocationResult:
    result_df: pd.DataFrame
    summary_df: pd.DataFrame
    log_df: Optional[pd.DataFrame] = None
    capacity: Optional[Dict[str, int]] = None
    seed: Optional[int] = None

def _pick_with_audit(
    rng: random.Random,
    candidates: List[int],
    k: int,
    stage: str,
    dept: str,
    id_values: Dict[int, str],
    log_records: List[dict],
) -> List[int]:
    """
    candidates 중 k명을 랜덤으로 뽑되,
    '랜덤값'을 부여해 정렬 후 앞에서 k명을 선택하는 방식으로
    추첨 기록을 남길 수 있게 한다.
    """
    if k <= 0 or len(candidates) == 0:
        return []

    # 랜덤값 부여
    scored = [(idx, rng.random()) for idx in candidates]
    scored.sort(key=lambda x: x[1])  # 낮을수록 당첨

    chosen_set = set(idx for idx, _ in scored[:k])

    # 로그
    for idx, score in scored:
        log_records.append({
            "단계": stage,
            "부서": dept,
            "row_index": idx,
            "식별자": id_values.get(idx, str(idx)),
            "랜덤값": score,
            "선정여부": (idx in chosen_set),
        })

    return [idx for idx, _ in scored[:k]]

def allocate_departments(
    df: pd.DataFrame,
    *,
    first_col: str,
    second_col: str,
    third_col: str,
    id_col: Optional[str] = None,
    seed: Optional[int] = None,
    include_log: bool = True,
) -> AllocationResult:
    """
    배정 실행.

    Parameters
    ----------
    df : DataFrame
        지원자 데이터
    first_col, second_col, third_col : str
        1/2/3순위 컬럼명
    id_col : str | None
        로그/출력에서 사람을 식별할 컬럼(예: 이름/학번). 없으면 행번호 사용.
    seed : int | None
        랜덤 시드(재현성). None이면 시스템 랜덤.
    include_log : bool
        추첨 로그를 결과에 포함할지 여부.

    Returns
    -------
    AllocationResult
    """
    work = df.copy()

    # 선호 컬럼 정규화
    for col in [first_col, second_col, third_col]:
        work[col] = work[col].apply(normalize_department)

    n = len(work)
    capacity = calc_capacity(n, DEPARTMENTS)
    remaining = capacity.copy()

    rng = random.Random(seed)

    # 식별자 값 준비
    if id_col is not None and id_col in work.columns:
        id_values = {idx: str(work.loc[idx, id_col]) for idx in work.index}
    else:
        # 사람이 보기 좋게 1부터 시작하는 번호
        id_values = {idx: f"#{i+1}" for i, idx in enumerate(work.index)}

    assigned_dept = {idx: None for idx in work.index}
    assigned_stage = {idx: None for idx in work.index}

    log_records: List[dict] = []

    def unassigned_indices() -> List[int]:
        return [idx for idx in work.index if assigned_dept[idx] is None]

    # 1/2/3순위 배정
    stages = [("1순위", first_col), ("2순위", second_col), ("3순위", third_col)]
    for stage_name, pref_col in stages:
        for dept in DEPARTMENTS:
            slots = remaining[dept]
            if slots <= 0:
                continue

            candidates = [idx for idx in unassigned_indices() if work.loc[idx, pref_col] == dept]
            if not candidates:
                continue

            if len(candidates) <= slots:
                chosen = candidates
                # 투명성을 위해, 미달이어도 로그를 남기고 싶다면 랜덤값 없이 기록
                if include_log:
                    for idx in candidates:
                        log_records.append({
                            "단계": stage_name,
                            "부서": dept,
                            "row_index": idx,
                            "식별자": id_values.get(idx, str(idx)),
                            "랜덤값": None,
                            "선정여부": True,
                            "비고": "미달(전원배정)",
                        })
            else:
                chosen = _pick_with_audit(
                    rng, candidates, slots, stage_name, dept, id_values, log_records
                )

            for idx in chosen:
                assigned_dept[idx] = dept
                assigned_stage[idx] = stage_name
            remaining[dept] -= len(chosen)

    # 5) 남은 TO에 랜덤 배정
    unassigned = unassigned_indices()
    remaining_slots: List[str] = []
    for dept in DEPARTMENTS:
        if remaining[dept] > 0:
            remaining_slots.extend([dept] * remaining[dept])

    # 전체 TO 합은 항상 N이므로 일반적으로 길이가 같아야 함.
    # 다만, 입력 데이터가 중복행/필터링 등으로 꼬였을 때를 대비해 방어적으로 처리.
    if len(remaining_slots) != len(unassigned):
        # 가능한 만큼만 배정하고, 나머지는 None 유지 (요약에서 확인 가능)
        # 이 상황은 사용자가 입력 파일을 확인해야 한다.
        # 그래도 프로그램이 죽지 않도록 한다.
        min_len = min(len(remaining_slots), len(unassigned))
        unassigned_to_assign = unassigned[:min_len]
        slots_to_use = remaining_slots[:min_len]
    else:
        unassigned_to_assign = unassigned
        slots_to_use = remaining_slots

    rng.shuffle(unassigned_to_assign)
    rng.shuffle(slots_to_use)

    for idx, dept in zip(unassigned_to_assign, slots_to_use):
        assigned_dept[idx] = dept
        assigned_stage[idx] = "랜덤"
        if include_log:
            log_records.append({
                "단계": "랜덤",
                "부서": dept,
                "row_index": idx,
                "식별자": id_values.get(idx, str(idx)),
                "랜덤값": rng.random(),
                "선정여부": True,
                "비고": "잔여TO 배정",
            })

    # 결과 DF 생성
    work["배정부서"] = pd.Series(assigned_dept)
    work["배정단계"] = pd.Series(assigned_stage)
    work["랜덤시드"] = seed
    # ISO 시간(UTC)로 기록 — 필요 시 사용자가 엑셀에서 현지시간으로 변환 가능
    work["배정시각_UTC"] = datetime.now(timezone.utc).isoformat(timespec="seconds")

    # 요약
    counts = work["배정부서"].value_counts(dropna=False).to_dict()
    summary_rows = []
    for dept in DEPARTMENTS:
        assigned = int(counts.get(dept, 0))
        to_ = int(capacity[dept])
        summary_rows.append({
            "부서": dept,
            "TO": to_,
            "최종배정": assigned,
            "잔여TO": to_ - assigned,
        })
    # 미배정(이상상황)도 표시
    unassigned_count = int(counts.get(None, 0) + counts.get(float("nan"), 0) if counts else 0)
    if work["배정부서"].isna().any():
        summary_rows.append({
            "부서": "(미배정)",
            "TO": 0,
            "최종배정": int(work["배정부서"].isna().sum()),
            "잔여TO": 0,
        })

    summary_df = pd.DataFrame(summary_rows)

    log_df = pd.DataFrame(log_records) if include_log else None

    return AllocationResult(
        result_df=work,
        summary_df=summary_df,
        log_df=log_df,
        capacity=capacity,
        seed=seed,
    )
