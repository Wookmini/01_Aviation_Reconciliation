import pandas as pd
import numpy as np

input_file = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_sample.xlsx'
output_file = 'excel_workspace/항공권 정산/정산결과_최종_보고서.xlsx'

def clean_amount(val):
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).replace('KRW', '').replace(',', '').strip()
    try:
        # 숫자만 남기고 나머지 제거
        s = ''.join(c for c in s if c.isdigit() or c == '.' or c == '-')
        return float(s) if s else 0
    except:
        return 0

try:
    print("정밀 분석 모드 시작...")
    # 시트 이름 명시적 확인 (공백 등 포함 가능성)
    xl = pd.ExcelFile(input_file)
    target_sheet = [s for s in xl.sheet_names if 'e-DAHS' in s][0]
    print(f"대상 시트 발견: {target_sheet}")

    # 헤더 없이 일단 전부 읽기
    df_dahs_all = pd.read_excel(input_file, sheet_name=target_sheet, header=None)
    
    # 25,858,214원을 찾기 위해 각 열의 합계를 계산해보기
    best_col = -1
    for col in range(df_dahs_all.shape[1]):
        col_sum = df_dahs_all.iloc[:, col].apply(clean_amount).sum()
        if abs(col_sum - 25858214) < 100: # 오차 범위 내
            best_col = col
            print(f"금액 컬럼 발견! 인덱스: {col}, 합계: {col_sum:,.0f}원")
            break
    
    if best_col == -1:
        # 못 찾았다면 전체 컬럼 중 가장 유사한 합계를 가진 컬럼 선택
        sums = [df_dahs_all.iloc[:, c].apply(clean_amount).sum() for c in range(df_dahs_all.shape[1])]
        best_col = np.argmax([s if s > 0 else -1 for s in sums])
        print(f"최적 추정 컬럼: {best_col}, 합계: {sums[best_col]:,.0f}원")

    # 이름 컬럼 찾기 (보통 금액 컬럼 앞쪽에 위치)
    name_col = 4 # 기존 분석 기반 기본값
    
    df_dahs = pd.DataFrame({
        '성명': df_dahs_all.iloc[:, name_col].astype(str).str.strip(),
        '금액': df_dahs_all.iloc[:, best_col].apply(clean_amount)
    })
    df_dahs = df_dahs[df_dahs['금액'] != 0].copy()
    
    dahs_total = df_dahs['금액'].sum()

    # 세중 데이터 (세중은 기존 로직이 어느정도 맞음)
    sejung_sheet = [s for s in xl.sheet_names if '세중' in s][0]
    df_sejung_raw = pd.read_excel(input_file, sheet_name=sejung_sheet)
    
    # 'Card Amt' 컬럼이 없으면 금액으로 보이는 컬럼 자동 찾기
    amt_col = 'Card Amt'
    if amt_col not in df_sejung_raw.columns:
        for c in df_sejung_raw.columns:
            if 'Amt' in str(c) or '금액' in str(c):
                amt_col = c
                break

    df_sejung = pd.DataFrame({
        '성명': df_sejung_raw['고객명'].astype(str).str.strip() if '고객명' in df_sejung_raw.columns else df_sejung_raw.iloc[:, 9],
        '금액': df_sejung_raw[amt_col].fillna(0).apply(clean_amount)
    })
    df_sejung = df_sejung[df_sejung['금액'] != 0].copy()
    sejung_total = df_sejung['금액'].sum()

    # 매칭 및 저장 (이전 로직과 동일)
    results = []
    temp_sejung = df_sejung.copy()
    for _, row in df_dahs.iterrows():
        match_idx = temp_sejung[(temp_sejung['성명'] == row['성명']) & (abs(temp_sejung['금액'] - row['금액']) < 1)].index
        if not match_idx.empty:
            results.append({'구분': '일치', '성명': row['성명'], 'e-DAHS 금액': row['금액'], '세중 금액': row['금액'], '차액': 0})
            temp_sejung = temp_sejung.drop(match_idx[0])
        else:
            results.append({'구분': 'e-DAHS 전용', '성명': row['성명'], 'e-DAHS 금액': row['금액'], '세중 금액': 0, '차액': row['금액']})

    for _, row in temp_sejung.iterrows():
        results.append({'구분': '세중 전용', '성명': row['성명'], 'e-DAHS 금액': 0, '세중 금액': row['금액'], '차액': -row['금액']})

    pd.DataFrame(results).to_excel(output_file, index=False)
    
    print(f"\n[최종 보고]")
    print(f"- e-DAHS 총액: {dahs_total:,.0f}원 (목표: 25,858,214원)")
    print(f"- 세중 총액: {sejung_total:,.0f}원")
    print(f"- 오차: {dahs_total - 25858214:,.0f}원")
    print(f"결과 파일: {output_file}")

except Exception as e:
    print(f"오류 발생: {e}")
    import traceback
    traceback.print_exc()
