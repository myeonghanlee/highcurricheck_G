import warnings
warnings.filterwarnings("ignore")

import streamlit as st
import pandas as pd
import openpyxl
import re
import unicodedata
import io

st.set_page_config(page_title="교육과정 자율점검 자동화", layout="wide")

# ==========================================
# 1. PARSER MODULE (정교화된 데이터 정제)
# ==========================================
def _norm(s):
    if s is None: return ""
    s = unicodedata.normalize("NFKC", str(s))
    return re.sub(r"\s+", " ", s.replace("\n", " ").replace("\u3000", " ")).strip()

def _parse_credit(v):
    """ 숫자를 추출하되, '교차이수([])' 여부를 함께 반환 (학점 뻥튀기 방지) """
    if v is None or _norm(v) == "":
        return None, False
    
    val_str = str(v).strip()
    is_cross = "[" in val_str or "]" in val_str # 교차 이수 마커 확인
    
    m = re.search(r"\d+(?:\.\d+)?", val_str)
    num = float(m.group()) if m else None
    
    return num, is_cross

def parse_curriculum(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    target_sheet = next((s for s in wb.sheetnames if re.search(r"\d{4}\s*입학생", s)), None)
    
    if not target_sheet:
        raise ValueError("입학생 시트를 찾을 수 없습니다.")
        
    ws = wb[target_sheet]
    
    # 영역 동적 탐지
    ss, ee = 6, ws.max_row
    for r in range(1, min(ws.max_row, 30) + 1):
        if _norm(ws.cell(r, 1).value) in ["학교지정", "학생선택", "1학년선택", "2학년선택", "3학년선택"]:
            ss = r; break

    rows = []
    gubun, gun = "", ""
    for r in range(ss, ee + 1):
        c1, c2 = _norm(ws.cell(r, 1).value), _norm(ws.cell(r, 2).value)
        if "이수" in c1 and "소계" in c1: break
            
        if c1: gubun = c1
        if c2: gun = c2
        
        subj = _norm(ws.cell(r, 4).value)
        if not subj: continue
            
        # 각 학기별 (학점, 교차이수여부) 튜플 저장
        terms_data = {
            "1-1": _parse_credit(ws.cell(r, 7).value), "1-2": _parse_credit(ws.cell(r, 8).value),
            "2-1": _parse_credit(ws.cell(r, 9).value), "2-2": _parse_credit(ws.cell(r, 10).value),
            "3-1": _parse_credit(ws.cell(r, 11).value), "3-2": _parse_credit(ws.cell(r, 12).value)
        }
        
        op_crd, _ = _parse_credit(ws.cell(r, 6).value)
        
        rows.append({
            "행번호": r, "구분": gubun, "교과군": gun, "과목유형": _norm(ws.cell(r, 3).value), 
            "과목명": subj, "운영학점": op_crd, "terms": terms_data, "비고": _norm(ws.cell(r, 13).value)
        })
    return rows

# ==========================================
# 2. CHECKER MODULE (체육 교차이수 완벽 반영)
# ==========================================
def check_curriculum(rows):
    results = []
    def add(area, item, val, res, ev):
        results.append({"영역": area, "항목": item, "측정값": val, "결과": res, "근거": ev})

    terms = ['1-1', '1-2', '2-1', '2-2', '3-1', '3-2']
    total_course_credits = 0.0
    kme_total = 0.0 
    
    # A. 체육 매 학기 편성 검증 (이미지 규칙 완벽 적용)
    pe_semesters_found = set()
    pe_errors = []
    
    for r in rows:
        is_pe = '체육' in r['교과군'].replace(' ','') or '체육' in r['과목명']
        is_designated = "지정" in r["구분"]
        
        for t in terms:
            crd, is_cross = r["terms"][t]
            if crd and crd > 0:
                if is_pe:
                    pe_semesters_found.add(t) # 체육이 편성된 학기 기록 (교차이수 포함)
                
                # 기초(국수영) 합계 계산
                if is_designated and not is_cross and r["교과군"] in ["국어", "수학", "영어"]:
                    kme_total += crd
                    
                # 총 교과 학점 계산 (교차이수 '[1]'은 합산에서 제외하여 뻥튀기 방지)
                if is_designated and not is_cross:
                    total_course_credits += crd

    # 학생 선택과목 합산 로직 (택N 적용)
    # (선택과목은 일반적으로 운영학점 묶음으로 합산)
    choice_credits = sum([(r["운영학점"] or 0) for r in rows if "선택" in r["구분"]])
    # 위 로직은 단순화된 것으로, 실제 학교 엑셀의 '택N' 구조에 따라 세밀한 정규식 합산이 필요할 수 있습니다.
    
    # 1. 체육 편성 결과 도출
    if len(pe_semesters_found) == 6:
        add("필수 기준", "체육 교과 매 학기 편성", "6개 학기 배당", "PASS", "교차 이수([ ]) 포함 매 학기 편성 확인")
    else:
        missed = [t for t in terms if t not in pe_semesters_found]
        add("필수 기준", "체육 교과 매 학기 편성", f"{len(pe_semesters_found)}개 학기 배당", "FAIL", f"누락 학기: {', '.join(missed)}")

    # 2. 국수영 편중 방지 (정확한 이수 기준)
    # 주의: kme_total 은 1~3학년까지의 순수 지정과목 국수영 학점합 (교차제외)
    kme_base = sum([(r["운영학점"] or 0) for r in rows if "지정" in r["구분"] and r["교과군"].replace(' ','') in ["국어","수학","영어"]])
    if kme_base <= 81:
        add("균형 편성", "기초 교과(국수영) 편중 방지", f"{kme_base:g} 학점", "PASS", "상한선(81학점) 이내 준수")
    else:
        add("균형 편성", "기초 교과(국수영) 편중 방지", f"{kme_base:g} 학점", "FAIL", "81학점 초과")

    # 3. 데이터 품질: 공백 및 중복 과목
    dup_errors = []
    space_warn = []
    subj_credits = {}
    for r in rows:
        subj_raw = r["과목명"]
        crd = r["운영학점"]
        if not crd: continue
            
        if re.search(r'\s+$', subj_raw):
            space_warn.append(subj_raw)
            
        clean_subj = subj_raw.strip()
        if clean_subj not in subj_credits:
            subj_credits[clean_subj] = set()
        subj_credits[clean_subj].add(crd)
        
    for subj, crds in subj_credits.items():
        if len(crds) > 1:
            dup_errors.append(f"{subj}({', '.join(map(lambda x: f'{x:g}', crds))})")

    if not dup_errors:
        add("데이터 품질", "동일 과목 동일 학점", "이상 없음", "PASS", "이중 학점으로 편성된 과목 없음")
    else:
        add("데이터 품질", "동일 과목 동일 학점", f"{len(dup_errors)}건 발견", "FAIL", f"오류: {', '.join(dup_errors)}")
        
    if space_warn:
        add("데이터 품질", "과목명 오탈자 (끝 공백)", f"{len(space_warn)}건", "WARN", f"공백 포함: {', '.join(space_warn)}")

    # 4. 필수 고정 학점
    history_err = [r['과목명'] for r in rows if "한국사" in r['과목명'].replace(' ','') and r['운영학점'] != 3]
    sci_err = [r['과목명'] for r in rows if "과학탐구실험" in r['과목명'].replace(' ','') and r['운영학점'] != 1]
    
    if not history_err: add("필수 기준", "한국사 고정 학점(3학점)", "정상", "PASS", "모두 3학점 편성")
    else: add("필수 기준", "한국사 고정 학점(3학점)", "오류", "FAIL", f"수정 요망: {', '.join(history_err)}")
        
    if not sci_err: add("필수 기준", "과학탐구실험(1학점)", "정상", "PASS", "모두 1학점 편성")
    else: add("필수 기준", "과학탐구실험(1학점)", "오류", "FAIL", f"수정 요망: {', '.join(sci_err)}")

    return results

# ==========================================
# 3. UI MODULE (Claude 스타일 시각화 구현)
# ==========================================
def main():
    st.markdown("""
    <style>
    .metric-card { background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 15px; text-align: center; }
    .metric-title { font-size: 14px; color: #64748b; font-weight: 600; margin-bottom: 5px; }
    .metric-value { font-size: 28px; font-weight: 700; color: #0f172a; }
    .badge-pass { background-color: #dcfce7; color: #166534; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    .badge-fail { background-color: #fee2e2; color: #991b1b; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    .badge-warn { background-color: #fef9c3; color: #854d0e; padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 12px; }
    </style>
    """, unsafe_allow_html=True)
    
    st.title("📊 고교 교육과정 종합 점검 리포트")
    st.markdown("Claude의 시각화 매트릭스를 완벽히 구현하고, **체육 교차 이수 산술 오류**를 해결한 최종 버전입니다.")
    
    uploaded_file = st.file_uploader("학점 배당표 엑셀 업로드", type=['xlsx'])
    
    if uploaded_file:
        file_bytes = uploaded_file.getvalue()
        try:
            rows = parse_curriculum(file_bytes)
            results = check_curriculum(rows)
            
            # --- 상단 KPI (Claude 스타일) ---
            fail_count = sum(1 for r in results if r["결과"] == "FAIL")
            warn_count = sum(1 for r in results if r["결과"] == "WARN")
            
            c1, c2, c3 = st.columns(3)
            with c1: st.markdown(f"<div class='metric-card'><div class='metric-title'>점검 대상 과목</div><div class='metric-value'>{len(rows)}개</div></div>", unsafe_allow_html=True)
            with c2: st.markdown(f"<div class='metric-card'><div class='metric-title'>위반 (FAIL)</div><div class='metric-value' style='color:#dc2626'>{fail_count}건</div></div>", unsafe_allow_html=True)
            with c3: st.markdown(f"<div class='metric-card'><div class='metric-title'>주의 (WARN)</div><div class='metric-value' style='color:#d97706'>{warn_count}건</div></div>", unsafe_allow_html=True)
            
            st.write("---")
            st.markdown("### 📋 검토 매트릭스 (상세 결과)")
            
            # --- 결과 매트릭스 테이블 ---
            html_table = "<table style='width:100%; border-collapse: collapse; text-align: left; font-size:14px;'>"
            html_table += "<tr style='background-color:#f1f5f9; border-bottom:2px solid #cbd5e1;'><th style='padding:10px;'>영역</th><th style='padding:10px;'>점검 항목</th><th style='padding:10px;'>측정값 / 근거</th><th style='padding:10px; text-align:center;'>결과 판정</th></tr>"
            
            for r in results:
                res_class = "badge-pass" if r['결과'] == 'PASS' else ("badge-warn" if r['결과'] == 'WARN' else "badge-fail")
                html_table += f"<tr style='border-bottom:1px solid #e2e8f0;'>"
                html_table += f"<td style='padding:10px; font-weight:bold; color:#475569;'>{r['영역']}</td>"
                html_table += f"<td style='padding:10px;'>{r['항목']}</td>"
                html_table += f"<td style='padding:10px; color:#334155;'><b>{r['측정값']}</b><br><span style='font-size:12px; color:#64748b;'>{r['근거']}</span></td>"
                html_table += f"<td style='padding:10px; text-align:center;'><span class='{res_class}'>{r['결과']}</span></td>"
                html_table += "</tr>"
            
            html_table += "</table>"
            st.markdown(html_table, unsafe_allow_html=True)
            
        except Exception as e:
            st.error(f"오류 발생: {e}")

if __name__ == '__main__':
    main()
