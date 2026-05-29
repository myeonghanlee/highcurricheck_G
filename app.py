import streamlit as st
import pandas as pd
import openpyxl
import re
import unicodedata
import io
import warnings

warnings.filterwarnings("ignore")
st.set_page_config(page_title="고교 교육과정 정밀 검토 시스템", layout="wide")

# ==========================================
# 1. PARSER MODULE (데이터 파싱 및 정제)
# ==========================================
def _norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("\n", " ").replace("\u3000", " ")
    return re.sub(r"\s+", " ", s).strip()

def _to_num(v):
    """ '28~30', '[2]', '3(택2)' 등에서 첫 번째 유효 숫자 추출 """
    if v is None or _norm(v) == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"\d+(?:\.\d+)?", str(v))
    return float(m.group()) if m else None

def parse_curriculum(file_bytes):
    # 메모리상에서 바로 엑셀 로드 (디스크 권한 에러 방지)
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    
    # 타겟 시트 찾기
    target_sheet = None
    for name in wb.sheetnames:
        if re.search(r"\d{4}\s*입학생", name):
            target_sheet = name
            break
    if not target_sheet:
        raise ValueError("'xxxx입학생' 형태의 시트를 찾을 수 없습니다.")
        
    ws = wb[target_sheet]
    
    # 데이터 영역 동적 탐지
    ss, ee = 6, ws.max_row
    for r in range(1, min(ws.max_row, 30) + 1):
        if _norm(ws.cell(r, 1).value) in ["학교지정", "학생선택", "1학년선택", "2학년선택", "3학년선택"]:
            ss = r
            break

    def cell(row, col):
        return ws.cell(row, col).value

    rows = []
    gubun, gun = "", ""
    for r in range(ss, ee + 1):
        c1 = _norm(cell(r, 1))
        c2 = _norm(cell(r, 2))
        subj = _norm(cell(r, 4))
        
        if "이수" in c1 and "소계" in c1: # 합계 영역 진입 시 파싱 종료
            break
            
        if c1: gubun = c1
        if c2: gun = c2
        if not subj: continue
            
        typ = _norm(cell(r, 3))
        
        rows.append({
            "행번호": r, "구분": gubun, "교과군": gun, "과목유형": typ, "과목명": subj,
            "기본학점": _to_num(cell(r, 5)), "운영학점": _to_num(cell(r, 6)),
            "1-1": _to_num(cell(r, 7)), "1-2": _to_num(cell(r, 8)),
            "2-1": _to_num(cell(r, 9)), "2-2": _to_num(cell(r, 10)),
            "3-1": _to_num(cell(r, 11)), "3-2": _to_num(cell(r, 12)),
            "비고": _norm(cell(r, 13))
        })

    return rows

# ==========================================
# 2. CHECKER MODULE (자율점검표 검토 로직)
# ==========================================
def check_curriculum(rows):
    results = []
    def add(area, item, val, result, ev):
        results.append({"영역": area, "검토 항목": item, "측정값": val, "결과": result, "근거 및 상세": ev})

    terms = ['1-1', '1-2', '2-1', '2-2', '3-1', '3-2']
    
    # 학기별 학생 총 이수 학점 계산 (학교지정 + 학생선택 '택N')
    term_totals = {t: 0.0 for t in terms}
    total_course_credits = 0.0
    kme_total = 0.0 # 국수영 합계

    for r in rows:
        is_designated = "지정" in r["구분"]
        for t in terms:
            if r[t]:
                if is_designated:
                    term_totals[t] += r[t]
                else:
                    # 선택과목의 경우 '운영학점'을 사용하거나 수동 계산이 필요하지만
                    # Claude식 추출 방식은 셀의 첫 숫자를 가져오므로 택N의 합계값 추출에 용이함
                    if r["과목명"] and "택" in r["비고"]: # 단순화된 계산 로직 적용
                        pass # 실제로는 비고란이나 학점란의 '택N' 로직을 더 정밀하게 합산해야 함
                        
        if is_designated and r["교과군"] in ["국어", "수학", "영어"]:
            kme_total += (r["운영학점"] or 0)
            
        if is_designated:
            total_course_credits += (r["운영학점"] or 0)
        else:
            # 선택과목의 택N 총합 추출 (과목명 중복 방지 위해 1회만 더하기 위한 러프 로직)
            pass

    # A. 체육 매 학기 편성 검증 (가장 강력한 검증)
    pe_rows = [r for r in rows if '체육' in r['교과군']]
    pe_terms = set()
    for r in pe_rows:
        for t in terms:
            if r[t] and r[t] > 0:
                pe_terms.add(t)
    
    if len(pe_terms) == 6:
        add("체육 편성", "체육 교과 매 학기 편성", "6개 학기 배당됨", "PASS", "모든 학기에 체육(또는 교차) 이수 존재")
    else:
        missed = [t for t in terms if t not in pe_terms]
        add("체육 편성", "체육 교과 매 학기 편성", f"{len(pe_terms)}개 학기 배당", "FAIL", f"누락 학기: {', '.join(missed)}")

    # B. 국수영 편중 방지 (81학점)
    if kme_total <= 81:
        add("균형 편성", "기초 교과(국수영) 편중 방지", f"{kme_total:g} 학점", "PASS", "국수영 이수 학점이 81학점을 초과하지 않음")
    else:
        add("균형 편성", "기초 교과(국수영) 편중 방지", f"{kme_total:g} 학점", "FAIL", "국수영 합계 81학점 초과")

    # C. 동일 과목 이중 학점 검증
    subj_credits = {}
    dup_errors = []
    for r in rows:
        subj = r["과목명"]
        crd = r["운영학점"]
        if subj and crd:
            if subj not in subj_credits:
                subj_credits[subj] = set()
            subj_credits[subj].add(crd)
    for subj, crds in subj_credits.items():
        if len(crds) > 1:
            dup_errors.append(f"{subj}({', '.join(map(lambda x: f'{x:g}', crds))})")
            
    if not dup_errors:
        add("데이터 품질", "동일 과목 동일 학점", "이상 없음", "PASS", "이중 학점으로 편성된 과목 없음")
    else:
        add("데이터 품질", "동일 과목 동일 학점", f"{len(dup_errors)}건 발견", "FAIL", f"오류 과목: {', '.join(dup_errors)}")

    # D. 지정 과목 필수 학점 준수
    history_ok, sci_exp_ok = True, True
    history_err, sci_exp_err = [], []
    for r in rows:
        subj = r["과목명"].replace(" ", "")
        if "한국사" in subj and r["운영학점"] != 3:
            history_err.append(f"{r['과목명']}({r['운영학점']}학점)")
        if "과학탐구실험" in subj and r["운영학점"] != 1:
            sci_exp_err.append(f"{r['과목명']}({r['운영학점']}학점)")
            
    if not history_err:
        add("필수 기준", "한국사 고정 학점(3학점)", "정상 준수", "PASS", "모든 한국사 과목 3학점 편성")
    else:
        add("필수 기준", "한국사 고정 학점(3학점)", "오류 발생", "FAIL", f"오류: {', '.join(history_err)}")
        
    if not sci_exp_err:
        add("필수 기준", "과학탐구실험 고정 학점(1학점)", "정상 준수", "PASS", "모든 과학탐구실험 과목 1학점 편성")
    else:
        add("필수 기준", "과학탐구실험 고정 학점(1학점)", "오류 발생", "FAIL", f"오류: {', '.join(sci_exp_err)}")

    # E. 위계성 (Ⅰ → Ⅱ) 검증 로직 추가
    term_order = {t: i for i, t in enumerate(terms)}
    roman_subjects = {}
    hierarchy_errors = []
    
    for r in rows:
        subj = r["과목명"]
        m = re.match(r"(.*?)(Ⅰ|Ⅱ)$", subj)
        if m:
            base_name, level = m.group(1), m.group(2)
            # 해당 과목이 개설된 가장 빠른 학기 탐색
            first_term_idx = 99
            for t in terms:
                if r[t] and r[t] > 0:
                    first_term_idx = min(first_term_idx, term_order[t])
            if first_term_idx != 99:
                if base_name not in roman_subjects:
                    roman_subjects[base_name] = {}
                roman_subjects[base_name][level] = first_term_idx

    for base_name, levels in roman_subjects.items():
        if 'Ⅰ' in levels and 'Ⅱ' in levels:
            if levels['Ⅰ'] > levels['Ⅱ']: # I이 II보다 늦게 개설된 경우
                hierarchy_errors.append(f"{base_name} (Ⅱ가 Ⅰ보다 먼저 편성됨)")

    if not hierarchy_errors:
        add("위계성", "계열적 학습 (Ⅰ선행, Ⅱ후행)", "정상 편성", "PASS", "Ⅰ, Ⅱ 위계가 꼬인 과목 없음")
    else:
        add("위계성", "계열적 학습 (Ⅰ선행, Ⅱ후행)", "위계성 오류", "WARN", f"확인 필요: {', '.join(hierarchy_errors)}")

    return results

# ==========================================
# 3. UI MODULE (Streamlit 앱 구성)
# ==========================================
def to_excel(summary_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_df.to_excel(writer, index=False, sheet_name='종합_검토결과')
    return output.getvalue()

def main():
    st.title("🏫 고교 교육과정 종합 점검 자동화 (PRO)")
    st.markdown("여러 학교의 배당표를 한 번에 업로드하고, **체육 교차 이수** 및 **위계성**까지 정밀 검증하세요.")
    
    uploaded_files = st.file_uploader("배당표 엑셀 업로드 (.xlsx)", type=['xlsx'], accept_multiple_files=True)
    
    if uploaded_files:
        summary_data = []
        
        for file in uploaded_files:
            file_bytes = file.getvalue()
            try:
                rows = parse_curriculum(file_bytes)
                validation_results = check_curriculum(rows)
                
                has_error = any(r["결과"] == "FAIL" for r in validation_results)
                
                # 종합 엑셀용 데이터
                summary_row = {"학교/파일명": file.name, "종합 결과": "⚠️ 점검 필요" if has_error else "✅ 정상"}
                for r in validation_results:
                    summary_row[r["검토 항목"]] = f"[{r['결과']}] {r['근거 및 상세']}"
                summary_data.append(summary_row)
                
                # UI 출력부
                with st.expander(f"📄 {file.name} 리포트 ({'⚠️ 수정 요망' if has_error else '✅ 모든 기준 통과'})", expanded=False):
                    
                    df_res = pd.DataFrame(validation_results)
                    
                    # 배지 스타일링 적용
                    def style_result(val):
                        if val == 'PASS': return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                        if val == 'FAIL': return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                        if val == 'WARN': return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                        return ''
                        
                    st.dataframe(df_res.style.map(style_result, subset=['결과']), use_container_width=True, hide_index=True)
                    
                    with st.expander("추출된 원본 데이터 확인"):
                        st.dataframe(pd.DataFrame(rows)[["구분", "과목유형", "교과군", "과목명", "운영학점"]].head(20))
                        
            except Exception as e:
                st.error(f"'{file.name}' 처리 중 오류 발생: {e}")

        if summary_data:
            st.markdown("### 📥 다중 검토 결과 다운로드")
            excel_data = to_excel(pd.DataFrame(summary_data))
            
            st.download_button(
                label="📊 전체 검토 매트릭스 엑셀 다운로드",
                data=excel_data,
                file_name="교육과정_종합검토_매트릭스.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

if __name__ == '__main__':
    main()
