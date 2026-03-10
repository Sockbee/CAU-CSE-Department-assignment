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

성별 TO 규칙 (gender_col 지정 시):
- 남/여 인원을 각각 6등분하여 성별별 부서 TO를 독립적으로 계산
- 남성 집단과 여성 집단을 완전히 분리하여 1/2/3순위 배정 진행
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, cast
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

# 성별 정규화 매핑
GENDER_MAP = {
    "남": "남",
    "남성": "남",
    "m": "남",
    "male": "남",
    "여": "여",
    "여성": "여",
    "f": "여",
    "female": "여",
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


def normalize_gender(value) -> Optional[str]:
    """입력값을 '남' 또는 '여'로 정규화. 매칭 실패/공백은 None."""
    if value is None:
        return None
    s = str(value).strip().lower()
    if s in {"nan", "none", "null", ""}:
        return None
    return GENDER_MAP.get(s, None)


def calc_capacity(n_people: int, departments: Optional[List[str]] = None) -> Dict[str, int]:
    """TO 규칙에 따라 부서별 정원을 계산."""
    if departments is None:
        departments = DEPARTMENTS
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
    # 성별 배정 시 성별별 TO 정보 (gender_col 미사용 시 None)
    capacity_by_gender: Optional[Dict[str, Dict[str, int]]] = None

def _pick_with_audit(
    rng: random.Random,
    candidates: List[int],
    k: int,
    stage: str,
    dept: str,
    id_values: Dict[int, str],
    log_records: List[dict],
    gender_label: Optional[str] = None,
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
        record = {
            "단계": stage,
            "부서": dept,
            "row_index": idx,
            "식별자": id_values.get(idx, str(idx)),
            "랜덤값": score,
            "선정여부": (idx in chosen_set),
        }
        if gender_label is not None:
            record["성별"] = gender_label
        log_records.append(record)

    return [idx for idx, _ in scored[:k]]


def _run_allocation_for_group(
    work: pd.DataFrame,
    group_indices: List[int],
    capacity: Dict[str, int],
    first_col: str,
    second_col: str,
    third_col: str,
    id_values: Dict[int, str],
    rng: random.Random,
    include_log: bool,
    log_records: List[dict],
    assigned_dept: Dict[int, Optional[str]],
    assigned_stage: Dict[int, Optional[str]],
    gender_label: Optional[str] = None,
) -> None:
    """
    주어진 group_indices(행 인덱스 목록)에 대해
    capacity 범위 내에서 1/2/3순위 → 랜덤 배정을 수행하고
    assigned_dept / assigned_stage 를 인플레이스로 갱신한다.
    """
    remaining = capacity.copy()

    def unassigned_in_group() -> List[int]:
        return [idx for idx in group_indices if assigned_dept[idx] is None]

    # 1/2/3순위 배정
    stages = [("1순위", first_col), ("2순위", second_col), ("3순위", third_col)]
    for stage_name, pref_col in stages:
        for dept in DEPARTMENTS:
            slots = remaining[dept]
            if slots <= 0:
                continue

            candidates = [
                idx for idx in unassigned_in_group()
                if work.loc[idx, pref_col] == dept
            ]
            if not candidates:
                continue

            if len(candidates) <= slots:
                chosen = candidates
                if include_log:
                    for idx in candidates:
                        record = {
                            "단계": stage_name,
                            "부서": dept,
                            "row_index": idx,
                            "식별자": id_values.get(idx, str(idx)),
                            "랜덤값": None,
                            "선정여부": True,
                            "비고": "미달(전원배정)",
                        }
                        if gender_label is not None:
                            record["성별"] = gender_label
                        log_records.append(record)
            else:
                chosen = _pick_with_audit(
                    rng, candidates, slots, stage_name, dept,
                    id_values, log_records, gender_label=gender_label,
                )

            for idx in chosen:
                assigned_dept[idx] = dept
                assigned_stage[idx] = stage_name
            remaining[dept] -= len(chosen)

    # 잔여 TO 랜덤 배정
    unassigned = unassigned_in_group()
    remaining_slots: List[str] = []
    for dept in DEPARTMENTS:
        if remaining[dept] > 0:
            remaining_slots.extend([dept] * remaining[dept])

    if len(remaining_slots) != len(unassigned):
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
            record = {
                "단계": "랜덤",
                "부서": dept,
                "row_index": idx,
                "식별자": id_values.get(idx, str(idx)),
                "랜덤값": rng.random(),
                "선정여부": True,
                "비고": "잔여TO 배정",
            }
            if gender_label is not None:
                record["성별"] = gender_label
            log_records.append(record)


def allocate_departments(
    df: pd.DataFrame,
    *,
    first_col: str,
    second_col: str,
    third_col: str,
    id_col: Optional[str] = None,
    gender_col: Optional[str] = None,
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
    gender_col : str | None
        성별 컬럼명. 지정하면 '남'/'여'를 분리하여 각 성별 인원 수 기준으로
        부서별 TO를 독립적으로 계산하고, 성별 집단 내에서만 배정한다.
        성별 인식 불가(공백·기타 값) 행은 별도 그룹으로 처리된다.
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

    rng = random.Random(seed)

    # 식별자 값 준비
    if id_col is not None and id_col in work.columns:
        id_values = {idx: str(work.loc[idx, id_col]) for idx in work.index}
    else:
        id_values = {idx: f"#{i+1}" for i, idx in enumerate(work.index)}

    assigned_dept: Dict[int, Optional[str]] = {idx: None for idx in work.index}
    assigned_stage: Dict[int, Optional[str]] = {idx: None for idx in work.index}

    log_records: List[dict] = []

    # ── 성별 분리 배정 ──────────────────────────────────────────────
    capacity_by_gender: Optional[Dict[str, Dict[str, int]]] = None
    overall_capacity: Dict[str, int]

    if gender_col is not None and gender_col in work.columns:
        # 성별 정규화
        work["_성별"] = work[gender_col].apply(normalize_gender)

        capacity_by_gender = {}
        groups = {}  # label -> List[index]

        for label in ["남", "여"]:
            mask = work["_성별"] == label
            indices = list(work.index[mask])
            groups[label] = indices
            cap = calc_capacity(len(indices), DEPARTMENTS)
            capacity_by_gender[label] = cap

        # 성별 인식 불가 행 (기타 그룹)
        other_indices = list(work.index[work["_성별"].isna()])
        if other_indices:
            cap_other = calc_capacity(len(other_indices), DEPARTMENTS)
            capacity_by_gender["기타"] = cap_other
            groups["기타"] = other_indices

        # 각 그룹별 배정
        for label, indices in groups.items():
            if not indices:
                continue
            _run_allocation_for_group(
                work=work,
                group_indices=indices,
                capacity=capacity_by_gender[label],
                first_col=first_col,
                second_col=second_col,
                third_col=third_col,
                id_values=id_values,
                rng=rng,
                include_log=include_log,
                log_records=log_records,
                assigned_dept=assigned_dept,
                assigned_stage=assigned_stage,
                gender_label=label,
            )

        # 전체 요약용 통합 capacity (성별 TO 합산)
        overall_capacity = {dept: 0 for dept in DEPARTMENTS}
        for cap in capacity_by_gender.values():
            for dept, v in cap.items():
                overall_capacity[dept] += v

        work.pop("_성별")  # 임시 컬럼 제거

    else:
        # ── 기존 동작 (성별 구분 없음) ────────────────────────────────
        overall_capacity = calc_capacity(len(work), DEPARTMENTS)
        all_indices = list(work.index)
        _run_allocation_for_group(
            work=work,
            group_indices=all_indices,
            capacity=overall_capacity,
            first_col=first_col,
            second_col=second_col,
            third_col=third_col,
            id_values=id_values,
            rng=rng,
            include_log=include_log,
            log_records=log_records,
            assigned_dept=assigned_dept,
            assigned_stage=assigned_stage,
            gender_label=None,
        )

    # ── 결과 DF 생성 ─────────────────────────────────────────────────
    work["배정부서"] = pd.Series(assigned_dept)
    work["배정단계"] = pd.Series(assigned_stage)
    work["랜덤시드"] = seed
    work["배정시각_UTC"] = datetime.now(timezone.utc).isoformat(timespec="seconds")

    # ── 요약 ─────────────────────────────────────────────────────────
    counts = work["배정부서"].value_counts(dropna=False).to_dict()
    summary_rows = []

    if capacity_by_gender is not None:
        assert isinstance(capacity_by_gender, dict)
        # 성별별 TO와 배정 결과를 함께 표시
        for dept in DEPARTMENTS:
            row: dict = {"부서": dept}
            total_to = 0
            total_assigned = 0
            for label in ["남", "여", "기타"]:
                if label not in capacity_by_gender:
                    continue
                label_cap = cast(Dict[str, int], cast(object, capacity_by_gender[label]))
                cap_val = int(label_cap.get(dept, 0))
                assigned_val = int(
                    work[(work["배정부서"] == dept) & (work.get("_성별", pd.Series(dtype=str)) == label)].shape[0]
                    if "_성별" in work.columns
                    else 0
                )
                # _성별 컬럼은 이미 제거됐으므로 gender_col로 직접 계산
                if gender_col is not None and gender_col in work.columns:
                    gender_normalized = work[gender_col].apply(normalize_gender)
                    assigned_val = int(
                        ((work["배정부서"] == dept) & (gender_normalized == label)).sum()
                    )
                row[f"TO({label})"] = cap_val
                row[f"배정({label})"] = assigned_val
                total_to += cap_val
                total_assigned += assigned_val
            row["TO(합계)"] = total_to
            row["최종배정"] = int(counts.get(dept, 0))
            row["잔여TO"] = total_to - row["최종배정"]
            summary_rows.append(row)
    else:
        for dept in DEPARTMENTS:
            assigned = int(counts.get(dept, 0))
            to_ = int(overall_capacity[dept])
            summary_rows.append({
                "부서": dept,
                "TO": to_,
                "최종배정": assigned,
                "잔여TO": to_ - assigned,
            })

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
        capacity=overall_capacity,
        seed=seed,
        capacity_by_gender=capacity_by_gender,
    )
