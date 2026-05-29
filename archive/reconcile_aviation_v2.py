import pandas as pd
import numpy as np

file_path = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_sample.xlsx'

def clean_amount(val):
    if pd.isna(val): return 0
    if isinstance(val, str):
        # 'KRW' 제거 및 숫자만 추출
        s = ''.join(filter(lambda x: x.isdigit() or x == '-' or x == '.', val))
        return float(s) if s else 0
    return float(val)

try:
    # 1. e-DAHS 데이터 로드 및 검증
    df_dahs = pd.read_excel(file_path, sheet_name='e-DAHS 전표내역')
    # 합계 컬럼 (보통 17번째 칸 근처) - 실제 컬럼명 확인 필요
    # 깨진 글자 대신 인덱스로 접근 시도
    dahs_amounts = df_dahs.iloc[:, 16].apply(clean_amount)
    dahs_total = dahs_amounts.sum()
    print(f"검증된 e-DAHS 총액: {dahs_total:,.0f}원")

    # 2. 세중 데이터 로드
    df_sejung = pd.read_excel(file_path, sheet_name='세중 발권리스트')
    # 'Card Amt' 컬럼 찾기
    sejung_amounts = df_sejung['Card Amt'].fillna(0)
    sejung_total = sejung_amounts.sum()
    print(f"검증된 세중 발권 총액: {sejung_total:,.0f}원")

    # 3. 1:1 매칭 로직 (성명 + 금액 기준)
    # 성명 컬럼 인덱스 확인 필요 (e-DAHS는 4번, 세중은 9번 근처)
    df_dahs['clean_name'] = df_dahs.iloc[:, 4].astype(str).str.strip()
    df_dahs['clean_amt'] = dahs_amounts
    
    df_sejung['clean_name'] = df_sejung['고객명'].astype(str).str.strip()
    df_sejung['clean_amt'] = sejung_amounts

    # 매칭 시작
    matched_results = []
    # e-DAHS 기준으로 세중 내역 찾기
    for idx, row in df_dahs.iterrows():
        match = df_sejung[(df_sejung['clean_name'] == row['clean_name']) & 
                          (df_sejung['clean_amt'] == row['clean_amt'])]
        if not match.empty:
            matched_results.append({
                "성명": row['clean_name'],
                "e-DAHS금액": row['clean_amt'],
                "세중금액": match.iloc[0]['clean_amt'],
                "결과": "일치"
            })
            # 매칭된 내역은 세중 리스트에서 제거 (중복 방지)
            df_sejung = df_sejung.drop(match.index[0])
        else:
            matched_results.append({
                "성명": row['clean_name'],
                "e-DAHS금액": row['clean_amt'],
                "세중금액": 0,
                "결과": "e-DAHS에만 존재"
            })
            
    # 남은 세중 내역 추가
    for idx, row in df_sejung.iterrows():
        matched_results.append({
            "성명": row['clean_name'],
            "e-DAHS금액": 0,
            "세중금액": row['clean_amt'],
            "결과": "세중에만 존재"
        })

    # 4. 결과 저장
    df_final = pd.DataFrame(matched_results)
    df_final.to_excel('excel_workspace/항공권 정산/최종_매칭_보고서.xlsx', index=False)
    print("매칭 보고서 생성이 완료되었습니다: 최종_매칭_보고서.xlsx")

except Exception as e:
    print(f"오류 발생: {e}")
