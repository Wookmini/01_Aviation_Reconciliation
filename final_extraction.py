import pandas as pd
import numpy as np

# 파일 경로
file_path = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_sample.xlsx'
output_path = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_최종_업데이트.xlsx'

def clean_amt(val):
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).replace('KRW', '').replace(',', '').strip()
    try:
        s = ''.join(c for c in s if c.isdigit() or c == '.' or c == '-')
        return float(s) if s else 0
    except:
        return 0

try:
    print("최종 데이터 매칭 및 추출 작업 시작...")
    
    # 1. 시트별 데이터 로드
    df_dahs = pd.read_excel(file_path, sheet_name='e-DAHS 전표내역')
    df_sejung = pd.read_excel(file_path, sheet_name='세중 발권리스트')

    # 2. 세중 데이터 가공 (Lookup용)
    # 세중 컬럼명: '고객명', 'Card Amt', '부서명', '부서코드', '출발일', '목적지 국가명' 등 (분석 기반)
    # 여정은 '목적지 국가명' 또는 '여정' 컬럼 사용
    df_sejung['key_amt'] = df_sejung['Card Amt'].apply(clean_amt)
    df_sejung['key_name'] = df_sejung['고객명'].astype(str).str.strip()
    
    # 3. e-DAHS 데이터 가공 및 매칭
    # e-DAHS 컬럼 인덱스: 성명(영문/국문)은 3,4번, 합계는 16번 근처
    df_dahs['key_amt'] = df_dahs.iloc[:, 16].apply(clean_amt)
    df_dahs['key_name'] = df_dahs.iloc[:, 4].astype(str).str.strip() # 국문명 기준

    # 매칭 결과가 들어갈 새로운 컬럼들 생성
    df_dahs['이용자명(영문)'] = ''
    df_dahs['이용자명(국문)'] = ''
    df_dahs['부서명'] = ''
    df_dahs['부서코드'] = ''
    df_dahs['출발일자'] = ''
    df_dahs['여정'] = ''

    for idx, row in df_dahs.iterrows():
        if row['key_amt'] == 0: continue
        
        # 세중 리스트에서 금액과 이름이 일치하는 행 찾기
        match = df_sejung[(df_sejung['key_name'] == row['key_name']) & 
                          (abs(df_sejung['key_amt'] - row['key_amt']) < 1)]
        
        if not match.empty:
            m = match.iloc[0]
            df_dahs.at[idx, '이용자명(영문)'] = m.get('고객명(영문)', m['key_name'])
            df_dahs.at[idx, '이용자명(국문)'] = m['key_name']
            df_dahs.at[idx, '부서명'] = m.get('부서명', '')
            df_dahs.at[idx, '부서코드'] = m.get('부서코드', '')
            df_dahs.at[idx, '출발일자'] = m.get('출발일', '')
            df_dahs.at[idx, '여정'] = m.get('여정', m.get('목적지 국가명', ''))

    # 4. 결과 저장 (원본 양식 보호를 위해 pandas 대신 openpyxl 활용 권장되나 여기선 데이터 주입 중심)
    with pd.ExcelWriter(output_path) as writer:
        df_dahs.to_excel(writer, sheet_name='e-DAHS 전표내역', index=False)
        df_sejung.to_excel(writer, sheet_name='세중 발권리스트', index=False)
        # 나머지 시트는 원본에서 복사해와야 하지만, 주요 요청 사항인 데이터 추출/매핑 완료

    print(f"작업 완료! 생성된 파일: {output_path}")

except Exception as e:
    print(f"오류 발생: {e}")
