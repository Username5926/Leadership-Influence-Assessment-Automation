import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import copy, io, zipfile, os, json, traceback
from pathlib import Path

from pptx import Presentation
from pptx.opc.package import _Relationship
from pptx.opc.packuri import PackURI
from pptx.parts.chart import ChartPart
from pptx.chart.data import ChartData

# ══════════════════════════════════════════════════════════════════
# 매핑
# ══════════════════════════════════════════════════════════════════
COMPETENCY_MAP = {
    "Position":     [9, 10, 17, 18, 22, 31],
    "Personality":  [5, 13, 14, 21, 24, 28],
    "Relationship": [6,  7, 23, 25, 30, 32],
    "Results":      [8, 16, 29, 33],
    "Development":  [4, 12, 20, 27],
    "Principles":   [11, 15, 19, 26],
}
SKILL_MAP = {
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
COMP_ROW  = {"Position":4,"Personality":5,"Relationship":6,
              "Results":7,"Development":8,"Principles":9}
SKILL_ROW = {"우호성":12,"동기유발":13,"자문":14,"협력제휴":15,
              "협상거래":16,"합리적설득":17,"합법화":18,"강요":19}

# ══════════════════════════════════════════════════════════════════
# 계산
# ══════════════════════════════════════════════════════════════════
def avg_rows(scores, rows):
    vals = [scores[r-3] for r in rows if (r-3) in scores]
    return round(sum(vals)/len(vals), 2) if vals else 0.0

def compute(scores):
    c = {k: avg_rows(scores, v) for k, v in COMPETENCY_MAP.items()}
    s = {k: avg_rows(scores, v) for k, v in SKILL_MAP.items()}
    return {"competency": c, "skill_raw": s,
            "soft_avg": round(sum(s[k] for k in SOFT_SKILLS)/3, 2),
            "hard_avg": round(sum(s[k] for k in HARD_SKILLS)/5, 2)}

# ══════════════════════════════════════════════════════════════════
# 파싱
# ══════════════════════════════════════════════════════════════════
def parse_people(raw: bytes) -> list:
    df = pd.read_excel(io.BytesIO(raw), header=0)
    out = []
    for _, row in df.iterrows():
        name = str(row.iloc[1]).strip()
        if not name or name.lower() == "nan":
            continue
        scores = {}
        for q in range(1, 31):
            try:
                scores[q] = float(row.iloc[q + 1])
            except:
                scores[q] = 0.0
        out.append({"name": name, "scores": scores})
    return out

# ══════════════════════════════════════════════════════════════════
# 엑셀 생성 (캐시 적용)
# ══════════════════════════════════════════════════════════════════
def _copy_ws(wb, src, title):
    ws = wb.create_sheet(title=title)
    for cl, cd in src.column_dimensions.items():
        ws.column_dimensions[cl].width = cd.width
    for rn, rd in src.row_dimensions.items():
        ws.row_dimensions[rn].height = rd.height
    for row in src.iter_rows():
        for cell in row:
            nc = ws.cell(row=cell.row, column=cell.column)
            nc.value = cell.value
            if cell.has_style:
                nc.font       = copy.copy(cell.font)
                nc.border     = copy.copy(cell.border)
                nc.fill       = copy.copy(cell.fill)
                nc.number_format = cell.number_format
                nc.protection = copy.copy(cell.protection)
                nc.alignment  = copy.copy(cell.alignment)
    for m in src.merged_cells.ranges:
        ws.merge_cells(str(m))
    return ws

@st.cache_data(show_spinner=False)
def build_excel(people_json: str, tpl: bytes) -> bytes:
    """people_json: JSON string (cache_data는 hashable 인수만 받음)"""
    people = json.loads(people_json)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for p in people:
        src = load_workbook(io.BytesIO(tpl)).worksheets[0]
        ws  = _copy_ws(wb, src, p["name"][:31])
        r   = compute(p["scores"])
        ws.cell(1, 1).value = p["name"]
        for q in range(1, 31):
            ws.cell(q+3, 3).value = p["scores"].get(str(q), 0)
        for k, rl in COMPETENCY_MAP.items():
            avg = r["competency"][k]
            ws.cell(COMP_ROW[k], 7).value = round(avg*len(rl), 2)
            ws.cell(COMP_ROW[k], 7).number_format = "0.00"
            ws.cell(COMP_ROW[k], 8).value = avg
            ws.cell(COMP_ROW[k], 8).number_format = "0.00"
        for k, rl in SKILL_MAP.items():
            avg = r["skill_raw"][k]
            ws.cell(SKILL_ROW[k], 7).value = round(avg*len(rl), 2)
            ws.cell(SKILL_ROW[k], 7).number_format = "0.00"
            ws.cell(SKILL_ROW[k], 8).value = avg
            ws.cell(SKILL_ROW[k], 8).number_format = "0.00"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════
# PPT 생성 (캐시 적용)
# ══════════════════════════════════════════════════════════════════
def _clone_chart(pkg, orig, new_name):
    nc = ChartPart(new_name, orig.content_type, pkg,
                   copy.deepcopy(orig._element))
    for rid, rel in orig.rels.items():
        nc.rels._rels[rid] = _Relationship(
            nc.partname.baseURI, rel._rId, rel._reltype,
            rel._target_mode, rel._target)
    return nc

def _clone_slide(prs, src_idx=0):
    src  = prs.slides[src_idx]
    sp   = src.part
    pkg  = prs.part.package
    ns   = prs.slides.add_slide(src.slide_layout)
    np_  = ns.part
    for ph in list(ns.placeholders):
        ph._element.getparent().remove(ph._element)
    for child in list(src.shapes._spTree):
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        ns.shapes._spTree.append(copy.deepcopy(child))
    cc = sum(1 for p in pkg.iter_parts()
             if str(p.partname).startswith("/ppt/charts/chart"))
    for rid, rel in sp.rels.items():
        if rid in np_.rels:
            continue
        if "chart" in rel._reltype:
            cc += 1
            nc = _clone_chart(pkg, rel._target,
                              PackURI(f"/ppt/charts/chart{cc}.xml"))
            np_.rels._rels[rid] = _Relationship(
                np_.partname.baseURI, rel._rId, rel._reltype,
                rel._target_mode, nc)
        else:
            np_.rels._rels[rid] = _Relationship(
                np_.partname.baseURI, rel._rId, rel._reltype,
                rel._target_mode, rel._target)

def _fill(slide, name, result):
    c  = result["competency"]
    s  = result["skill_raw"]
    sa = result["soft_avg"]
    ha = result["hard_avg"]
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if "{{NAME}}" in run.text:
                        run.text = run.text.replace("{{NAME}}", name)
        if shape.name == "table_phase" and shape.has_table:
            tbl = shape.table
            for i, (k, v) in enumerate(c.items(), 1):
                if i >= len(tbl.rows): break
                tbl.cell(i, 0).text = k
                tbl.cell(i, 1).text = f"{v:.2f}"
        elif shape.name == "table_strategy" and shape.has_table:
            rows = ([(k, s[k]) for k in SOFT_SKILLS] +
                    [("소프트스킬 평균", sa)] +
                    [(k, s[k]) for k in HARD_SKILLS] +
                    [("하드스킬 평균", ha)])
            tbl = shape.table
            for i, (k, v) in enumerate(rows, 1):
                if i >= len(tbl.rows): break
                tbl.cell(i, 0).text = str(k)
                tbl.cell(i, 1).text = f"{v:.2f}"
        elif shape.name == "chart_phase" and shape.has_chart:
            try:
                cd = ChartData()
                cd.categories = list(c.keys())
                cd.add_series("역량 점수", list(c.values()))
                shape.chart.replace_data(cd)
            except Exception as e:
                st.warning(f"chart_phase 오류: {e}")
        elif shape.name == "chart_strategy" and shape.has_chart:
            try:
                cats = SOFT_SKILLS + ["소프트 평균"] + HARD_SKILLS + ["하드 평균"]
                vals = ([s[k] for k in SOFT_SKILLS] + [sa] +
                        [s[k] for k in HARD_SKILLS] + [ha])
                cd = ChartData()
                cd.categories = cats
                cd.add_series("계열 1", vals)
                shape.chart.replace_data(cd)
            except Exception as e:
                st.warning(f"chart_strategy 오류: {e}")

@st.cache_data(show_spinner=False)
def build_ppt(people_json: str, tpl: bytes) -> bytes:
    """people_json: JSON string (cache_data는 hashable 인수만 받음)"""
    people = json.loads(people_json)
    prs = Presentation(io.BytesIO(tpl))
    # 1단계: 복제 (항상 slides[0] 기준)
    for _ in range(len(people) - 1):
        _clone_slide(prs, src_idx=0)
    # 2단계: 데이터 주입
    for i, person in enumerate(people):
        _fill(prs.slides[i], person["name"], compute(person["scores"]))
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════
# 템플릿 자동 탐색 — 확장자 기준
# ══════════════════════════════════════════════════════════════════
def find_template(ext: str):
    """GitHub 루트에서 확장자로 탐색. 파일명 무관."""
    base = Path(__file__).parent
    found = sorted(base.glob(f"*{ext}"))
    if found:
        return found[0].read_bytes(), str(found[0])
    found = sorted(Path(os.getcwd()).glob(f"*{ext}"))
    if found:
        return found[0].read_bytes(), str(found[0])
    return None, None

# ══════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════
st.set_page_config(page_title="리더십 진단 보고서 자동화", layout="wide")
st.title("📊 리더십 진단 보고서 자동화 툴")

# 사이드바
with st.sidebar:
    st.header("🔍 환경")
    try:
        base_dir = Path(__file__).parent
        files = [f.name for f in sorted(base_dir.iterdir()) if f.is_file()]
        st.write("루트 파일:", files)
    except Exception as e:
        st.write(f"오류: {e}")
    _, ep = find_template(".xlsx")
    _, pp = find_template(".pptx")
    st.write("📁 엑셀:", ep or "❌")
    st.write("📁 PPT:", pp or "❌")

st.markdown("---")

col1, col2, col3 = st.columns(3)
with col1:
    response_file  = st.file_uploader("① 구글 폼 응답 엑셀 (필수)", type=["xlsx","xls"])
with col2:
    excel_tpl_file = st.file_uploader("② 엑셀 템플릿 (선택)", type=["xlsx"])
with col3:
    ppt_tpl_file   = st.file_uploader("③ PPT 템플릿 (선택)", type=["pptx"])

st.markdown("---")

if st.button("🚀 보고서 생성", type="primary", use_container_width=True):

    # 응답 엑셀 읽기
    if response_file is None:
        st.error("❌ 응답 엑셀을 업로드해주세요."); st.stop()
    resp_bytes = response_file.read()

    # 엑셀 템플릿
    if excel_tpl_file is not None:
        excel_tpl = excel_tpl_file.read()
        st.success("✅ 엑셀 템플릿: 업로드 파일")
    else:
        excel_tpl, ep = find_template(".xlsx")
        if excel_tpl:
            st.success(f"✅ 엑셀 템플릿 자동: `{ep}`")
        else:
            st.error("❌ 엑셀 템플릿 없음"); st.stop()

    # PPT 템플릿
    if ppt_tpl_file is not None:
        ppt_tpl = ppt_tpl_file.read()
        st.success("✅ PPT 템플릿: 업로드 파일")
    else:
        ppt_tpl, pp = find_template(".pptx")
        if ppt_tpl:
            st.success(f"✅ PPT 템플릿 자동: `{pp}`")
        else:
            st.error("❌ PPT 템플릿 없음"); st.stop()

    # 파싱
    try:
        people = parse_people(resp_bytes)
    except Exception as e:
        st.error(f"❌ 파싱 실패: {e}")
        st.code(traceback.format_exc()); st.stop()

    if not people:
        st.error("❌ 응답자 없음"); st.stop()

    st.info(f"👥 {len(people)}명: {[p['name'] for p in people]}")

    with st.expander("📋 점수 미리보기"):
        rows = []
        for p in people:
            r = compute(p["scores"])
            row = {"성함": p["name"]}
            row.update({k: f"{v:.2f}" for k, v in r["competency"].items()})
            row["소프트평균"] = r["soft_avg"]
            row["하드평균"]   = r["hard_avg"]
            rows.append(row)
        st.dataframe(pd.DataFrame(rows), use_container_width=True)

    # JSON으로 변환 (cache_data용 — scores 키를 str로)
    people_for_cache = [
        {"name": p["name"], "scores": {str(k): v for k, v in p["scores"].items()}}
        for p in people
    ]
    people_json = json.dumps(people_for_cache, ensure_ascii=False)

    # 엑셀 생성
    with st.spinner(f"📊 엑셀 {len(people)}시트 생성 중..."):
        try:
            excel_out = build_excel(people_json, excel_tpl)
            st.success(f"✅ 엑셀 {len(people)}시트 완료")
        except Exception as e:
            st.error(f"❌ 엑셀 실패: {e}")
            st.code(traceback.format_exc()); st.stop()

    # PPT 생성
    with st.spinner(f"📑 PPT {len(people)}슬라이드 생성 중..."):
        try:
            ppt_out = build_ppt(people_json, ppt_tpl)
            n_slides = len(Presentation(io.BytesIO(ppt_out)).slides)
            st.success(f"✅ PPT {n_slides}슬라이드 완료")
        except Exception as e:
            st.error(f"❌ PPT 실패: {e}")
            st.code(traceback.format_exc()); st.stop()

    # 다운로드 — 버튼 클릭 블록 안에서 바로 표시
    # (cache_data 덕분에 rerun 후에도 동일 결과 반환)
    st.markdown("---")
    st.success(f"🎉 완료: 엑셀 {len(people)}시트 + PPT {n_slides}슬라이드")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("리더십진단_개인별.xlsx", excel_out)
        zf.writestr("리더십진단_통합.pptx",   ppt_out)

    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button("⬇️ ZIP (전체)", data=zip_buf.getvalue(),
            file_name="리더십진단_결과.zip", mime="application/zip",
            use_container_width=True)
    with d2:
        st.download_button("⬇️ 엑셀", data=excel_out,
            file_name="리더십진단_개인별.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with d3:
        st.download_button("⬇️ PPT", data=ppt_out,
            file_name="리더십진단_통합.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True)
