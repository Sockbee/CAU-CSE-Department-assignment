# 부서 배정기 (구글폼 → 엑셀 기반)

## 1) 설치
Python 3.10+ 권장

```bash
pip install -r requirements.txt
```

## 2) 실행(추천: 웹 UI)
```bash
streamlit run app.py
```

브라우저가 열리면 엑셀(.xlsx) 업로드 → 1/2/3순위 컬럼 선택 → (선택) 시드 입력 → 결과 다운로드

## 3) 실행(대안: CLI)
```bash
python allocate_cli.py --input input.xlsx --output output.xlsx --sheet Sheet1 \
  --first "1순위" --second "2순위" --third "3순위" --id "이름" --seed 20260303
```

## 입력 엑셀 형식
- 1행 = 1명
- 최소 컬럼: 1순위, 2순위, 3순위
- (선택) 이름/학번 같은 식별자 컬럼

부서명은 다음 중 하나로 들어오면 됩니다.
- 기획부 / 문화부 / 연대사업부 / 선전부 / 교육부 / 미디어부
- 또는 줄여쓴 형태(기획, 문화, 연대, 선전, 교육, 미디어)도 자동 인식합니다.
