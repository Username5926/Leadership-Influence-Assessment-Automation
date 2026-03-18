import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
import io
import os

# [매핑] 이미지 및 원본 엑셀 수식 기준 (Q1 ~ Q30) 
PHASE_MAPPING = {
    "Position": [6, 7, 14, 15, 19, 28], "Personality": [2, 10, 11, 18, 21, 25],
    "Relationship": [3, 4, 20, 22, 27, 29], "Results": [5, 13, 26, 30],
    "Development": [1, 9, 17, 24], "Principles": [8, 12, 16, 23]
}
STRATEGY_MAPPING = {
    "우호성": [1, 9, 17, 24], "동기유발": [2, 10, 18, 25], "자문": [3, 11],
    "협력제휴": [4, 12, 19, 26], "협상거래": [5, 13, 20, 27], "합리적설득": [6, 14, 21, 28],
    "합법화": [7, 15, 22, 29], "강요": [8, 16, 23, 30]
}

def process_all_data(df, template_path):
    prs = Presentation(template_path)
    slide_layout = prs.slides[0].slide_layout
    
    # 엑셀 파일 객체 생성 (메모리 상에서 작업)
    output_xlsx = io.BytesIO()
    
    with pd.ExcelWriter(output_xlsx, engine='xlsxwriter') as writer:
        for idx, row in df.iterrows():
            name = str(row['성함을 작성해주세요.'])
            scores = row.iloc[2:32].values # 문항 데이터 

            # 1. 엑셀: 개인별 시트 생성 및 양식 구성
            # 원본 양식과 유사하게 문항/점수 배치 
            individual_df = pd.DataFrame({
                '구분': range(1, 31),
                '문항 점수': scores
            })
            individual_df.to_excel(writer, sheet_name=name[:30], index=False, startrow=1)
            
            # 계산 로직 적용 (소계/평균) 
            p_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in PHASE_MAPPING.items()}
            s_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in STRATEGY_MAPPING.items()}

            # 2. PPT: 슬라이드 추가 및 개체 업데이트
            slide = prs.slides[0] if idx == 0 else prs.slides.add_slide(slide_layout)
            
            for shape in slide.shapes:
                if shape.has_text_frame and "{{NAME}}" in shape.text:
                    shape.text = shape.text.replace("{{NAME}}", name)
                
                if shape.name == 'table_phase' and shape.has_table:
                    for i, val in enumerate(p_scores.values()):
                        shape.table.cell(i, 1).text = str(val)
                
                if shape.name == 'table_strategy' and shape.has_table:
                    for i, val in enumerate(s_scores.values()):
                        shape.table.cell(i+1, 1).text = str(val)

                if shape.name == 'chart_phase' and shape.has_chart:
                    chart_data = CategoryChartData()
                    chart_data.categories = list(p_scores.keys())
                    chart_data.add_series('Score', list(p_scores.values()))
                    shape.chart.replace_data(chart_data)

    return output_xlsx.getvalue(), prs

# Streamlit 설정
st.title("📊 HRD 진단 결과 통합 생성기")
uploaded_file = st.file_uploader("구글 폼 결과 엑셀 업로드", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    template_name = "template.pptx.pptx" # 깃허브 파일명 확인!
    
    xlsx_data, final_ppt = process_all_data(df, template_name)

    if xlsx_data:
        st.success(f"✅ {len(df)}명의 데이터 처리가 완료되었습니다.")
        
        # 다운로드 버튼
        st.download_button("📂 1. 개인별 시트가 포함된 엑셀 다운로드", xlsx_data, "진단결과_원본양식통합.xlsx")
        
        ppt_buffer = io.BytesIO()
        final_ppt.save(ppt_buffer)
        st.download_button("📊 2. 모든 인원이 포함된 PPT 다운로드", ppt_buffer.getvalue(), "최종진단보고서.pptx")
