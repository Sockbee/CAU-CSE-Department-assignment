"""
CLI (커맨드라인) 버전
예시:
  python allocate_cli.py --input input.xlsx --output output.xlsx --sheet Sheet1 \
    --first "1순위" --second "2순위" --third "3순위" --id "이름" --seed 20260303

Streamlit UI가 더 편하면 app.py를 추천합니다.
"""

import argparse
import pandas as pd

from allocator import allocate_departments

def main():
    p = argparse.ArgumentParser(description="부서 배정기 (Excel)")
    p.add_argument("--input", required=True, help="입력 엑셀 파일(.xlsx)")
    p.add_argument("--output", required=True, help="출력 엑셀 파일(.xlsx)")
    p.add_argument("--sheet", default=None, help="시트명(없으면 첫 시트)")
    p.add_argument("--first", required=True, help="1순위 컬럼명")
    p.add_argument("--second", required=True, help="2순위 컬럼명")
    p.add_argument("--third", required=True, help="3순위 컬럼명")
    p.add_argument("--id", default=None, help="식별자 컬럼명(이름/학번 등)")
    p.add_argument("--seed", type=int, default=None, help="랜덤 시드(재현성)")
    p.add_argument("--no-log", action="store_true", help="추첨로그 생성 안 함")
    args = p.parse_args()

    xl = pd.ExcelFile(args.input)
    sheet = args.sheet if args.sheet is not None else xl.sheet_names[0]
    df = xl.parse(sheet_name=sheet)

    res = allocate_departments(
        df,
        first_col=args.first,
        second_col=args.second,
        third_col=args.third,
        id_col=args.id,
        seed=args.seed,
        include_log=not args.no_log,
    )

    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        res.result_df.to_excel(writer, index=False, sheet_name="배정결과")
        res.summary_df.to_excel(writer, index=False, sheet_name="요약")
        if (not args.no_log) and res.log_df is not None:
            res.log_df.to_excel(writer, index=False, sheet_name="추첨로그")

    print("완료:", args.output)

if __name__ == "__main__":
    main()
