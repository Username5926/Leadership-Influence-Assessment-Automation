import pandas as pd
import streamlit as st
from pptx import Presentation
from pptx.chart.data import CategoryChartData
import io
import copy

# [매핑] 이미지 및 요청 사항 기반 최종 확정본
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

def process_automation(input_df, template_path):
    prs = Presentation(template_path)
    # 첫 번째 슬라이드는 수동으로 처리하고, 두 번째 응답자부터 슬라이드를 추가하거나 
    # 혹은 모든 응답자에 대해 새 슬라이드를 구성합니다.
    
    refined_list = []

    for idx, row in input_df.iterrows():
        name = str(row['성함을 작성해주세요.'])
        # 실제 데이터 시작 컬럼 확인 (타임스탬프, 이름 이후 2번 인덱스부터 30개)
        scores = row.iloc[2:32].values 
        
        # 1. 개인별 점수 계산
        p_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in PHASE_MAPPING.items()}
        s_scores = {k: round(sum([float(scores[q-1]) for q in v])/len(v), 2) for k, v in STRATEGY_MAPPING.items()}
        
        refined_list.append({"성함": name, **p_scores, **s_scores})

        # 2. PPT 데이터 주입 (각 응답자별 슬라이드 처리)
        # 예시를 위해 첫 슬라이드를 계속 재사용하거나 복제 로직을 태웁니다.
        if idx == 0:
            slide = prs.slides[0]
        else:
            # 템플릿의 첫 슬라이드 레이아웃을 사용하여 새 슬라이드 추가
            slide_layout = prs.slides[0].slide_layout
            slide = prs.slides.add_slide(slide_layout)

        for shape in slide.shapes:
            # 이름 변경
            if shape.has_text_frame and "{{NAME}}" in shape.text:
                shape.text = shape.text.replace("{{NAME}}", name)
            
            # 표 업데이트
            if shape.name == 'table_phase' and shape.has_table:
                for i, val in enumerate(p_scores.values()):
                    shape.table.cell(i, 1).text = str(val)
            
            if shape.name == 'table_strategy' and shape.has_table:
                for i, val in enumerate(s_scores.values()):
                    shape.table.cell(i, 1).text = str(val)

            # 차트 업데이트
            if shape.name == 'chart_phase' and shape.has_chart:
                chart_data = CategoryChartData()
                chart_data.categories = list(p_scores.keys())
                chart_data.add_series('점수', list(p_scores.values()))
                shape.chart.replace_data(chart_data)

    return pd.DataFrame(refined_list), prs

# 스트림릿 UI
st.title("🚗 넥센타이어 진단 자동화 시스템")
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # 파일명이 pptx.pptx 인 경우 아래 이름을 수정하세요
    refined_df, final_ppt = process_automation(df, "pptx.pptx") 

    st.success(f"{len(df)}명의 데이터 분석 완료!")

    # 엑셀 다운로드
    output_xlsx = io.BytesIO()
    with pd.ExcelWriter(output_xlsx, engine='openpyxl') as writer:
        refined_df.to_excel(writer, index=False)
    st.download_button("1. 정제된 엑셀 다운로드", output_xlsx.getvalue(), "진단결과_정제본.xlsx")

    # PPT 다운로드
    output_ppt = io.BytesIO()
    final_ppt.save(output_ppt)
    st.download_button("2. 통합 PPT 보고서 다운로드", output_ppt.getvalue(), "최종결과보고서.pptx")
