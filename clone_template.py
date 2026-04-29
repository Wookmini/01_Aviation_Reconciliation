import pandas as pd
import shutil
from openpyxl import load_workbook

# 파일 경로
template_file = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_260420.xlsx'
sample_file = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_sample.xlsx'
output_file = 'excel_workspace/항공권 정산/항공권 결제내역 정산(04월분)_자동완성.xlsx'

# 1. 먼저 완성본 파일을 복사하여 양식(껍데기)을 그대로 가져옵니다.
shutil.copy(template_file, output_file)

try:
    print("양식 복제 완료. 데이터 채우기 시작...")
    
    # 2. Sample 파일에서 원본 데이터들 읽기
    df_dahs = pd.read_excel(sample_file, sheet_name='e-DAHS 전표내역')
    df_sejung = pd.read_excel(sample_file, sheet_name='세중 발권리스트')
    df_card = pd.read_excel(sample_file, sheet_name='하나카드 결제내역')

    # 3. 생성된 '자동완성' 파일의 데이터 시트들만 Sample 데이터로 교체
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_dahs.to_excel(writer, sheet_name='e-DAHS 전표내역', index=False)
        df_sejung.to_excel(writer, sheet_name='세중 발권리스트', index=False)
        df_card.to_excel(writer, sheet_name='하나카드 결제내역', index=False)

    print(f"데이터 주입 완료! 최종 결과물: {output_file}")
    print("이제 이 파일은 완성본(260420)의 양식을 유지하면서 데이터만 Sample로 바뀐 상태입니다.")

except Exception as e:
    print(f"작업 중 오류 발생: {e}")
