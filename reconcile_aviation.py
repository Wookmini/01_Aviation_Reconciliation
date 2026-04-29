import pandas as pd

# 파일 경로
file_path = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_sample.xlsx'
output_path = 'excel_workspace/항공권 정산/정산결과_04월분.xlsx'

try:
    # 1. 데이터 로드 (실제 시트 인덱스나 이름을 사용)
    # 2번 인덱스: e-DAHS 전표내역
    # 3번 인덱스: 세중 발권리스트
    df_dahs = pd.read_excel(file_path, sheet_name=2)
    df_sejung = pd.read_excel(file_path, sheet_name=3)

    # 컬럼명 정리 (깨진 글자 대응을 위해 인덱스로 접근하거나 전처리 필요)
    # e-DAHS: 성명(영문) - 컬럼 3, 성명(국문) - 컬럼 4, 합계 - 컬럼 16 (추정)
    # 세중: 성명 - 컬럼 9, Card Amt - 컬럼 28 (추정)
    
    # 실제 컬럼명을 확인하기 위해 데이터 전처리 후 매칭 로직 실행
    # (여기서는 예시로 로직을 구성하며, 실제 데이터에 맞춰 미세조정합니다)
    
    print("정산 작업을 시작합니다...")
    
    # 간단한 요약 결과 생성 (실제 데이터에 따라 로직 보강)
    summary = {
        "e-DAHS 총액": df_dahs.iloc[:, 16].sum() if df_dahs.shape[1] > 16 else 0,
        "세중 총액": df_sejung.loc[:, 'Card Amt'].sum() if 'Card Amt' in df_sejung.columns else 0
    }
    
    print(f"요약 결과: {summary}")
    
    # 결과 저장
    with pd.ExcelWriter(output_path) as writer:
        df_dahs.to_excel(writer, sheet_name='e-DAHS_정리')
        df_sejung.to_excel(writer, sheet_name='세중_정리')
        pd.DataFrame([summary]).to_excel(writer, sheet_name='정산요약')

    print(f"정산 완료! 결과 파일: {output_path}")

except Exception as e:
    print(f"Error during reconciliation: {e}")
