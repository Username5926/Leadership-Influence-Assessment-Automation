"""
리더십 진단 보고서 자동화 툴 v4
- 실제 첨부 템플릿(excel_template.xlsx / template_pptx.pptx) 구조 기반
- 엑셀: A1:C1 병합 셀에 성함, C4:C33 점수, G/H열 수식 보존 (평균 직접 주입)
- PPT:  spTree 복사 + _Relationship 올바른 시그니처로 rels 복사 → 다중 슬라이드
- chart_phase / chart_strategy: COLUMN_CLUSTERED 차트 데이터 교체
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import copy, io, zipfile

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.opc.package import _Relationship
from pptx.chart.data import ChartData

# ══════════════════════════════════════════════════════════════════
# 1. 매핑 정의 — 엑셀 C열 행번호 기준 (row = Q번호 + 3)
# ══════════════════════════════════════════════════════════════════

COMPETENCY_MAP = {   # 역량명: [C열 행번호, ...]
    "Position":     [9, 10, 17, 18, 22, 31],
    "Personality":  [5, 13, 14, 21, 24, 28],
    "Relationship": [6,  7, 23, 25, 30, 32],
    "Results":      [8, 16, 29, 33],
    "Development":  [4, 12, 20, 27],
    "Principles":   [11, 15, 19, 26],
}

SKILL_MAP = {        # 기술명: [C열 행번호, ...]
    "우호성":     [4, 12, 20, 27],
    "동기유발":   [5, 13, 21, 28],
    "자문":       [6, 14],
    "협력제휴":   [7, 15, 22, 29],
    "협상거래":   [8, 16, 23, 30],
    "합리적설득": [9, 17, 24, 31],
    "합법화":     [10, 18, 25, 32],
    "강요":       [11, 19, 26, 33],
}

SOFT_SKILLS = ["우호성", "동기유발", "자문"]
HARD_SKILLS = ["협력제휴", "협상거래", "합리적설득", "합법화", "강요"]

# G/H열 행 위치 (엑셀 템플릿 실측 기준)
COMP_ROW   = {"Position":4,"Personality":5,"Relationship":6,"Results":7,"Development":8,"Principles":9}
SKILL_ROW  = {"우호성":12,"동기유발":13,"자문":14,"협력제휴":15,"협상거래":16,"합리적설득":17,"합법화":18,"강요":19}

# ══════════════════════════════════════════════════════════════════
# 2. 점수 계산
# ══════════════════════════════════════════════════════════════════

def avg_by_rows(scores: dict, row_list: list) -> float:
    """row_list의 행번호를 Q번호로 변환(Q = row - 3) 후 평균."""
    vals = [scores[r - 3] for r in row_list if (r - 3) in scores]
    return round(sum(vals) / len(vals), 2) if vals else 0.0


def compute_person(scores: dict) -> dict:
    competency = {k: avg_by_rows(scores, v) for k, v in COMPETENCY_MAP.items()}
    skill_raw  = {k: avg_by_rows(scores, v) for k, v in SKILL_MAP.items()}
    soft_avg   = round(sum(skill_raw[k] for k in SOFT_SKILLS) / len(SOFT_SKILLS), 2)
    hard_avg   = round(sum(skill_raw[k] for k in HARD_SKILLS) / len(HARD_SKILLS), 2)
    return {"competency": competency, "skill_raw": skill_raw,
            "soft_avg": soft_avg, "hard_avg": hard_avg}

# ══════════════════════════════════════════════════════════════════
# 3. 입력 파싱
# ══════════════════════════════════════════════════════════════════

def parse_response_excel(file) -> list:
    """A열:타임스탬프, B열:성함, C~AF열:Q1~Q30 (0-based col = Q+1)"""
    df = pd.read_excel(file, header=0)
    people = []
    for _, row in df.iterrows():
        name = str(row.iloc[1]).strip()
        if not name or name.lower() in ("nan", ""):
            continue
        scores = {}
        for q in range(1, 31):
            try:
                scores[q] = float(row.iloc[q + 1])
            except Exception:
                scores[q] = 0.0
        people.append({"name": name, "scores": scores})
    return people

# ══════════════════════════════════════════════════════════════════
# 4. 엑셀 출력 생성
# ══════════════════════════════════════════════════════════════════

def _copy_sheet(wb_dest, src_ws, new_title: str):
    ws = wb_dest.create_sheet(title=new_title)
    for col_letter, cd in src_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = cd.width
    for row_num, rd in src_ws.row_dimensions.items():
        ws.row_dimensions[row_num].height = rd.height
    for row in src_ws.iter_rows():
        for cell in row:
            nc = ws.cell(row=cell.row, column=cell.column)
            nc.value = cell.value
            if cell.has_style:
                nc.font          = copy.copy(cell.font)
                nc.border        = copy.copy(cell.border)
                nc.fill          = copy.copy(cell.fill)
                nc.number_format = cell.number_format
                nc.protection    = copy.copy(cell.protection)
                nc.alignment     = copy.copy(cell.alignment)
    for merge in src_ws.merged_cells.ranges:
        ws.merge_cells(str(merge))
    return ws   # ← 반드시 return


def build_excel(people: list, template_src) -> bytes:
    """
    실제 템플릿 기준:
      A1 (병합 A1:C1) → 성함
      C4:C33          → Q1~Q30 점수
      G4:G9 / H4:H9   → 6대역량 소계/평균 직접 주입
      G12:G19 / H12:H19 → 8대기술 소계/평균 직접 주입
      시트명          → 응답자 성함
    """
    if hasattr(template_src, "read"):
        raw = template_src.read()
        template_src = io.BytesIO(raw)

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    for person in people:
        template_src.seek(0)
        wb_tpl = load_workbook(template_src)
        src_ws = wb_tpl.worksheets[0]
        ws = _copy_sheet(wb_out, src_ws, person["name"][:31])
        result = compute_person(person["scores"])

        # ① 성함 (A1 — 병합 셀 A1:C1)
        ws.cell(row=1, column=1).value = person["name"]

        # ② Q1~Q30 점수 (C4:C33)
        for q in range(1, 31):
            ws.cell(row=q + 3, column=3).value = person["scores"].get(q, 0)

        # ③ 6대역량 소계(G) / 평균(H)
        for key, row_list in COMPETENCY_MAP.items():
            er  = COMP_ROW[key]
            avg = result["competency"][key]
            ws.cell(row=er, column=7).value = round(avg * len(row_list), 2)
            ws.cell(row=er, column=7).number_format = "0.00"
            ws.cell(row=er, column=8).value = avg
            ws.cell(row=er, column=8).number_format = "0.00"

        # ④ 8대기술 소계(G) / 평균(H)
        for key, row_list in SKILL_MAP.items():
            er  = SKILL_ROW[key]
            avg = result["skill_raw"][key]
            ws.cell(row=er, column=7).value = round(avg * len(row_list), 2)
            ws.cell(row=er, column=7).number_format = "0.00"
            ws.cell(row=er, column=8).value = avg
            ws.cell(row=er, column=8).number_format = "0.00"

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════════
# 5. PPT 출력 생성
# ══════════════════════════════════════════════════════════════════

def _clone_slide_safe(prs: Presentation, src_slide_index: int = 0):
    """
    슬라이드 완전 복제 (chart/image 관계 포함):
      1. spTree 하위 elements deep-copy
      2. _Relationship(base_uri, rId, reltype, target_mode, target) 로 rels 정확히 복사
    """
    src_slide = prs.slides[src_slide_index]
    src_part  = src_slide.part

    layout    = src_slide.slide_layout
    new_slide = prs.slides.add_slide(layout)
    new_part  = new_slide.part

    # 자동 생성된 placeholder 제거
    for ph in list(new_slide.placeholders):
        ph._element.getparent().remove(ph._element)

    # spTree 복사
    for child in list(src_slide.shapes._spTree):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        new_slide.shapes._spTree.append(copy.deepcopy(child))

    # rels 복사 (_Relationship 정확한 시그니처 사용)
    for rId, rel in src_part.rels.items():
        if rId in new_part.rels:
            continue
        new_part.rels._rels[rId] = _Relationship(
            new_part.partname.baseURI,  # base_uri
            rel._rId,                   # rId
            rel._reltype,               # reltype
            rel._target_mode,           # target_mode ("Internal" / "External")
            rel._target                 # target (Part 또는 str)
        )

    return new_slide


def _replace_text(shape, old: str, new_val: str):
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if old in run.text:
                run.text = run.text.replace(old, new_val)


def _set_table_data(shape, rows_data: list):
    """1행부터 데이터 주입. 0행 헤더 유지."""
    tbl = shape.table
    for r_idx, (label, val) in enumerate(rows_data, start=1):
        if r_idx >= len(tbl.rows):
            break
        tbl.cell(r_idx, 0).text = str(label)
        tbl.cell(r_idx, 1).text = f"{val:.2f}" if isinstance(val, float) else str(val)


def _update_chart(shape, labels: list, values: list):
    try:
        cd = ChartData()
        cd.categories = labels
        cd.add_series("계열 1", values)
        shape.chart.replace_data(cd)
    except Exception:
        pass


def _fill_slide(slide, name: str, result: dict):
    competency = result["competency"]
    skill_raw  = result["skill_raw"]
    soft_avg   = result["soft_avg"]
    hard_avg   = result["hard_avg"]

    for shape in slide.shapes:
        # {{NAME}} 치환 (text_NAME, TextBox 27 등 모두)
        _replace_text(shape, "{{NAME}}", name)

        # 6대역량 표
        if shape.name == "table_phase" and shape.has_table:
            _set_table_data(shape, list(competency.items()))

        # 8대기술 표
        elif shape.name == "table_strategy" and shape.has_table:
            rows = [(k, skill_raw[k]) for k in SOFT_SKILLS]
            rows.append(("소프트스킬 평균", soft_avg))
            rows += [(k, skill_raw[k]) for k in HARD_SKILLS]
            rows.append(("하드스킬 평균", hard_avg))
            _set_table_data(shape, rows)

        # 6대역량 차트
        elif shape.name == "chart_phase" and shape.has_chart:
            _update_chart(shape, list(competency.keys()), list(competency.values()))

        # 8대기술 차트
        elif shape.name == "chart_strategy" and shape.has_chart:
            _update_chart(shape, list(skill_raw.keys()), list(skill_raw.values()))


def build_ppt(people: list, template_src) -> bytes:
    """
    핵심 전략:
      1. 먼저 슬라이드[0] 기준으로 복제를 모두 수행 (데이터 주입 전)
      2. 복제 완료 후 각 슬라이드에 데이터 주입
    """
    if hasattr(template_src, "read"):
        raw = template_src.read()
        template_src = io.BytesIO(raw)

    template_src.seek(0)
    prs = Presentation(template_src)

    # 1단계: 슬라이드 복제 (항상 원본 슬라이드[0] 기준)
    slides = [prs.slides[0]]
    for _ in range(len(people) - 1):
        slides.append(_clone_slide_safe(prs, src_slide_index=0))

    # 2단계: 데이터 주입
    for slide, person in zip(slides, people):
        result = compute_person(person["scores"])
        _fill_slide(slide, person["name"], result)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════════
# 6. Streamlit UI
# ══════════════════════════════════════════════════════════════════

st.set_page_config(page_title="리더십 진단 보고서 자동화", layout="wide")
st.markdown("""
<style>
.main-title{font-size:2rem;font-weight:800;color:#1F3864;}
.sub{font-size:1rem;color:#666;margin-bottom:.5rem;}
.sec{font-size:1.05rem;font-weight:700;color:#2E75B6;margin-top:1.2rem;margin-bottom:.3rem;}
</style>""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 리더십 진단 보고서 자동화 툴</div>', unsafe_allow_html=True)
st.markdown('<div class="sub">구글 폼 응답 엑셀 → 개인별 엑셀 (평균값 포함) + 응답자별 슬라이드 PPT</div>',
            unsafe_allow_html=True)
st.markdown("---")

st.markdown('<div class="sec">① 구글 폼 응답 엑셀 (필수)</div>', unsafe_allow_html=True)
response_file = st.file_uploader(
    "A열: 타임스탬프 / B열: 성함 / C~AF열: Q1~Q30 점수",
    type=["xlsx", "xls"], key="response"
)

st.markdown('<div class="sec">② 엑셀 템플릿 (선택 – 미업로드 시 GitHub의 excel_template.xlsx 자동 사용)</div>',
            unsafe_allow_html=True)
excel_tpl_file = st.file_uploader(
    "A1:C1 병합=성함, C4:C33=점수, G/H열=소계/평균 양식",
    type=["xlsx"], key="excel_tpl"
)

st.markdown('<div class="sec">③ PPT 템플릿 (선택 – 미업로드 시 GitHub의 template_pptx.pptx 자동 사용)</div>',
            unsafe_allow_html=True)
ppt_tpl_file = st.file_uploader(
    "{{NAME}}, table_phase, table_strategy, chart_phase, chart_strategy 개체 포함",
    type=["pptx"], key="ppt_tpl"
)

st.markdown("---")

if st.button("🚀 보고서 생성", type="primary", use_container_width=True):
    if not response_file:
        st.error("❌ 구글 폼 응답 엑셀을 업로드해주세요.")
        st.stop()

    def load_template(uploaded, path):
        if uploaded:
            return uploaded
        try:
            with open(path, "rb") as f:
                return io.BytesIO(f.read())
        except FileNotFoundError:
            st.error(f"❌ '{path}' 파일을 찾을 수 없습니다. 직접 업로드해주세요.")
            st.stop()

    excel_src = load_template(excel_tpl_file, "excel_template.xlsx")
    ppt_src   = load_template(ppt_tpl_file,   "template_pptx.pptx")

    if not excel_tpl_file: st.info("ℹ️ 엑셀 템플릿: GitHub 루트의 excel_template.xlsx 사용")
    if not ppt_tpl_file:   st.info("ℹ️ PPT 템플릿: GitHub 루트의 template_pptx.pptx 사용")

    with st.spinner("📂 응답 데이터 파싱 중..."):
        try:
            people = parse_response_excel(response_file)
        except Exception as e:
            st.error(f"❌ 파싱 실패: {e}"); st.stop()

    if not people:
        st.error("❌ 응답자 데이터가 비어있습니다."); st.stop()

    st.success(f"✅ {len(people)}명 응답 데이터 파싱 완료")

    with st.expander(f"📋 응답자 미리보기 ({len(people)}명)"):
        preview = []
        for p in people:
            r = compute_person(p["scores"])
            row = {"성함": p["name"]}
            row.update({k: f"{v:.2f}" for k, v in r["competency"].items()})
            for k in SOFT_SKILLS + HARD_SKILLS:
                row[k] = f"{r['skill_raw'][k]:.2f}"
            row["소프트스킬 평균"] = f"{r['soft_avg']:.2f}"
            row["하드스킬 평균"]   = f"{r['hard_avg']:.2f}"
            preview.append(row)
        st.dataframe(pd.DataFrame(preview), use_container_width=True)

    with st.spinner("📊 개인별 엑셀 생성 중..."):
        try:
            excel_bytes = build_excel(people, excel_src)
        except Exception as e:
            st.error(f"❌ 엑셀 생성 실패: {e}"); st.exception(e); st.stop()

    with st.spinner(f"📑 PPT 생성 중 ({len(people)}슬라이드)..."):
        try:
            ppt_bytes = build_ppt(people, ppt_src)
        except Exception as e:
            st.error(f"❌ PPT 생성 실패: {e}"); st.exception(e); st.stop()

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("리더십진단_개인별.xlsx", excel_bytes)
        zf.writestr("리더십진단_통합.pptx",  ppt_bytes)
    zip_buf.seek(0)

    st.balloons()
    st.success(f"🎉 완료! 엑셀 {len(people)}시트 + PPT {len(people)}슬라이드")

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button("⬇️ ZIP (전체)", data=zip_buf,
            file_name="리더십진단_결과.zip", mime="application/zip",
            use_container_width=True)
    with d2:
        st.download_button("⬇️ 엑셀 (개인별)", data=excel_bytes,
            file_name="리더십진단_개인별.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with d3:
        st.download_button("⬇️ PPT (통합)", data=ppt_bytes,
            file_name="리더십진단_통합.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True)
