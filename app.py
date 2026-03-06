"""
Streamlit UI: 엑셀 업로드 -> 컬럼 선택 -> 시드 입력 -> 배정 결과 다운로드
실행:
  streamlit run app.py
"""

import io
import pandas as pd
import streamlit as st

from allocator import allocate_departments, DEPARTMENTS

st.set_page_config(page_title="부서 배정기", layout="wide")

st.title("부서 배정기 (1/2/3순위 + TO 자동 계산)")

with st.expander("사용 방식", expanded=True):
    st.markdown(
        """
- 구글폼 결과를 엑셀(.xlsx)로 저장한 뒤 업로드하세요.
- 1/2/3순위 컬럼을 선택하고, (선택) 이름/학번 같은 **식별자 컬럼**을 고르세요.
- 랜덤 배정의 재현성을 원하면 **시드(seed)** 값을 숫자로 입력하세요. (예: 20260303)
- 결과는 `배정결과` / `요약` / (선택) `추첨로그` 시트로 내려받을 수 있습니다.
        """
    )

uploaded = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.stop()

# 시트 선택
xl = pd.ExcelFile(uploaded)
sheet_name = st.selectbox("시트 선택", xl.sheet_names, index=0)

df = xl.parse(sheet_name=sheet_name)

st.subheader("미리보기")
st.dataframe(df.head(20), use_container_width=True)

cols = list(df.columns)

col1, col2, col3 = st.columns(3)

with col1:
    first_col = st.selectbox("1순위 컬럼", cols, index=0)
with col2:
    second_col = st.selectbox("2순위 컬럼", cols, index=min(1, len(cols)-1))
with col3:
    third_col = st.selectbox("3순위 컬럼", cols, index=min(2, len(cols)-1))

id_col = st.selectbox("식별자(이름/학번) 컬럼 (선택)", ["(없음)"] + cols, index=0)

seed_str = st.text_input("랜덤 시드(seed) (선택, 숫자)", value="20260303")
include_log = st.checkbox("추첨 로그 포함(권장)", value=True)

run = st.button("배정 실행")

if not run:
    st.stop()

seed = None
if seed_str.strip() != "":
    try:
        seed = int(seed_str.strip())
    except ValueError:
        st.error("시드는 숫자만 입력할 수 있어요. 예: 20260303")
        st.stop()

result = allocate_departments(
    df,
    first_col=first_col,
    second_col=second_col,
    third_col=third_col,
    id_col=None if id_col == "(없음)" else id_col,
    seed=seed,
    include_log=include_log,
)

st.success("배정 완료!")

st.subheader("요약")
st.dataframe(result.summary_df, use_container_width=True)

st.subheader("배정 결과 미리보기")
st.dataframe(result.result_df.head(50), use_container_width=True)

# 엑셀로 내보내기
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    result.result_df.to_excel(writer, index=False, sheet_name="배정결과")
    result.summary_df.to_excel(writer, index=False, sheet_name="요약")
    if include_log and result.log_df is not None:
        result.log_df.to_excel(writer, index=False, sheet_name="추첨로그")

output.seek(0)

st.download_button(
    label="📥 배정 결과 엑셀 다운로드",
    data=output.getvalue(),
    file_name="부서배정_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

with st.expander("부서 목록 / TO 규칙", expanded=False):
    st.write("부서 순서(나머지 배분 순서):", " / ".join(DEPARTMENTS))
    st.markdown(
        """
TO는 다음과 같이 계산됩니다.
- 전체 인원 N을 6으로 나눈 몫을 각 부서 기본 TO로 배정
- 나머지 r명은 `기획부 → 문화부 → 연대사업부 → 선전부 → 교육부 → 미디어부` 순서로 1명씩 추가
        """
    )
