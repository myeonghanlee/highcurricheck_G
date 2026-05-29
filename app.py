import streamlit as st
import pandas as pd
import numpy as np
import re
import io

st.set_page_config(page_title="고교 교육과정 자율점검 자동화", layout="wide")

# --- 1. 정밀 데이터 파싱 및 오류 탐지 함수 ---
def parse_and_audit(file):
    try:
        df_raw = pd.read_excel(file, sheet_name='2026입학생', header=None)
        
        row2 = df_raw.iloc[2].replace(r'^\s*$', np.nan, regex=True).ffill().fillna('')
        row3 = df_raw.iloc[3].fillna('')
        
        raw_headers = []
        for a, b in zip(row2, row3):
            header = f"{str(a).strip()}_{str(b).strip()}".strip('_').replace('\n', '')
            raw_headers.append(header if header else "빈칸")
            
        headers = []
        seen = {}
        for h in raw_headers:
            if h in seen:
                seen[h] += 1
                headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 0
                headers.append(h)
                
        df = pd.DataFrame(df_raw.values[6:], columns=headers)
        
        df['구분'] = df['구분'].replace('', np.nan).ffill().str.replace('\n', '').str.strip()
        subject_col = [col for col in df.columns if '교과' in col][0]
        df[subject_col] = df[subject_col].replace('', np.nan).ffill().str.replace('\n', ' ').str.strip()
        
        course_col = [col for col in df.columns if '과목' in col and '유형' not in col and '개설' not in col][0]
        type_col = [col for col in df.columns if '과목' in col and '유형' in col][0]
        
        df_valid = df.dropna(subset=[course_col]).copy()
        
        df_valid[course_col] = df_valid[course_col].astype(str).str.strip()
        df_valid[type_col] = df_valid[type_col].astype(str).str.strip()
        
        credit_col = [col for col in df.columns if '운영' in col and '학점' in col][0]
        df_valid[credit_col] = pd.to_numeric(df_valid[credit_col], errors='coerce').fillna(0)
        
        try:
            comp_col_name = [c for c in df_valid.columns if '이수' in c.replace('\n', '') and '필수' not in c.replace('\n', '')][0]
            req_col_name = [c for c in df_valid.columns if '필수' in c.replace('\n', '') and '이수' in c.replace('\n', '')][0]
            df_valid[comp_col_name] = pd.to_numeric(df_valid[comp_col_name], errors='coerce').fillna(0)
            df_valid[req_col_name] = pd.to_numeric(df_valid[req_col_name], errors='coerce').fillna(0)
            df_valid = df_valid.rename(columns={comp_col_name: '이수학점_합계', req_col_name: '필수이수학점_기준'})
        except IndexError:
            df_valid['이수학점_합계'] = 0
            df_valid['필수이수학점_기준'] = 0

        df_valid = df_valid.rename(columns={subject_col: '교과(군)', course_col: '과목', credit_col: '운영학점', type_col: '과목유형'})
        
        duplicate_issues = {}
        groups = df_valid.groupby('과목')['운영학점'].nunique()
        dup_subjects = groups[groups > 1].index.tolist()
        for sub in dup_subjects:
            credits_used = df_valid[df_valid['과목'] == sub]['운영학점'].unique()
            duplicate_issues[sub] = list(credits_used)

        sem1_cols = [c for c in df_valid.columns if '1학년' in c]
        first_year_electives = []
        for c in sem1_cols:
            has_val = df_valid[(df_valid['과목유형'] == '일반') & (df_valid[c].notna()) & (df_valid[c].astype(str).str.strip() != '')]
            if not has_val.empty:
                first_year_electives.extend(has_val['과목'].tolist())
        
        return df_valid, duplicate_issues, list(set(first_year_electives))
        
    except Exception as e:
        st.error(f"파싱 오류: {e}")
        return None, {}, []

# --- 2. 학점 계산 함수 (교차 이수 대괄호 지원) ---
def calculate_credits(df):
    semester_cols = ['1학년_1학기', '1학년_2학기', '2학년_1학기', '2학년_2학기', '3학년_1학기', '3학년_2학기']
    designated_mask = df['구분'].str.contains('학교지정', na=False)
    
    choice_total_credits = 0
    semester_student_totals = {sem: 0.0 for sem in semester_cols}
    
    for sem in semester_cols:
        if sem in df.columns:
            # 1. 학교지정 학기별 합산 (대괄호 '[1]' 등 교차이수 문자열의 숫자 추출 기능 추가)
            des_sem_vals = df[designated_mask][sem].dropna().astype(str).str.strip()
            des_sem_sum = 0
            for val in des_sem_vals:
                # '[1]', '2' 등에서 순수 숫자만 추출하여 합산
                match = re.search(r'\[?(\d+(\.\d+)?)\]?', val)
                if match:
                    des_sem_sum += float(match.group(1))
            semester_student_totals[sem] += des_sem_sum
            
            # 2. 학생선택 '택N' 합산
            unique_choice = df[~designated_mask][sem].dropna().unique()
            sem_choice_sum = 0
            for val in unique_choice:
                match = re.search(r'(\d+)\s*\(택', str(val))
                if match:
                    sem_choice_sum += int(match.group(1))
            
            semester_student_totals[sem] += sem_choice_sum
            choice_total_credits += sem_choice_sum
                    
    # 학교지정 총 학점은 각 학기 합계의 절반 (교차 이수 '[1]'과 '1'이 양쪽에 있으면 두번 더해지므로)
    # 하지만 엑셀의 '운영학점' 열이 정확하다면 그걸 쓰는 것이 가장 안전합니다.
    designated_total = df[designated_mask]['운영학점'].sum()
    
    course_total = designated_total + choice_total_credits
    grand_total = course_total + 18
    
    return int(course_total), int(grand_total), semester_student_totals

# --- 3. 종합 검토 함수 (체육 교차 이수 검증 강화) ---
def validate_curriculum(df, actual_total_credits, semester_student_totals, duplicate_issues):
    results = []
    
    group_comp = df.groupby('교과(군)')['이수학점_합계'].max()
    group_req = df.groupby('교과(군)')['필수이수학점_기준'].max()

    # 1. 총 이수 학점
    if actual_total_credits >= 174:
        results.append(("총 교과 이수 학점", f"총 {actual_total_credits}학점 (기준 174 이상)", "성공"))
    else:
        results.append(("총 교과 이수 학점", f"총 {actual_total_credits}학점 (기준 174 미달)", "실패"))

    # 2. 필수 이수 학점
    req_sum = 0
    failed_req = []
    for subj in group_req.index:
        req_val = group_req[subj]
        comp_val = group_comp[subj]
        if req_val > 0:
            if req_val <= comp_val:
                req_sum += req_val
            else:
                failed_req.append(f"{subj.strip()}(필수 {int(req_val)}/이수 {int(comp_val)})")

    if not failed_req and req_sum >= 84:
        results.append(("필수 이수 학점 충족 여부", f"필수 학점 합계 {int(req_sum)}학점 충족 완료", "성공"))
    else:
        results.append(("필수 이수 학점 충족 여부", f"미달 발생: {', '.join(failed_req)}", "실패"))

    # 3. 국수영 편중 방지
    basic_credits = 0
    for subj in group_comp.index:
        if any(keyword in subj for keyword in ['국어', '수학', '영어', '국 어', '수 학', '영 어']):
            basic_credits += group_comp[subj]
            
    if basic_credits <= 81:
        results.append(("국수영 편중 방지", f"국수영 이수학점 총합 {int(basic_credits)}학점 (81 상한 준수)", "성공"))
    else:
        results.append(("국수영 편중 과다", f"국수영 이수학점 총합 {int(basic_credits)}학점 (81 상한 초과)", "실패"))

    # 4. [개선] 체육 교과 매 학기 편성 검토 (교차 이수 완벽 인식)
    pe_df = df[df['교과(군)'].str.contains('체육', na=False)]
    pe_miss = []
    semester_cols = ['1학년_1학기', '1학년_2학기', '2학년_1학기', '2학년_2학기', '3학년_1학기', '3학년_2학기']
    
    for sem in semester_cols:
        if sem in pe_df.columns:
            # 해당 학기에 1, 2, [1] 등 숫자를 포함한 유효한 배당 문자가 있는지 정규식으로 검사
            valid_entries = pe_df[sem].dropna().astype(str).str.strip().str.contains(r'\[?\d+\]?', regex=True)
            if not valid_entries.any():
                pe_miss.append(sem.replace('_', ' '))
                
    if not pe_miss:
        results.append(("체육 매 학기 편성", "교차 이수 포함, 6개 학기 모두 체육 과목 배당 정상 확인", "성공"))
    else:
        results.append(("체육 편성 누락", f"누락 학기: {', '.join(pe_miss)}", "실패"))

    # 5. 학기 격차
    actual_sem_vals = [v for v in semester_student_totals.values() if v > 0]
    if actual_sem_vals:
        diff = max(actual_sem_vals) - min(actual_sem_vals)
        if diff <= 5:
            results.append(("학기 격차", f"최대-최소 차이 {diff}학점 (5 이내 준수)", "성공"))
        else:
            results.append(("학기 격차 초과", f"학기 간 차이 {diff}학점 (5 초과)", "실패"))

    # 6. 동일 과목 이중 학점
    if not duplicate_issues:
        results.append(("동일 과목 동일 학점", "이중 학점 편성 없음", "성공"))
    else:
        results.append(("동일 과목 이중 학점", f"{len(duplicate_issues)}건 오류 발견", "실패"))

    return results

# --- 4. 엑셀 다운로드 파일 생성 함수 ---
def to_excel(summary_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_df.to_excel(writer, index=False, sheet_name='종합_검토결과')
    processed_data = output.getvalue()
    return processed_data

# --- 5. UI 메인 로직 ---
def main():
    st.title("📊 고교 교육과정 다중 점검 대시보드")
    st.markdown("여러 학교의 배당표를 한 번에 업로드하여 검증하고 결과를 엑셀로 다운로드하세요.")
    
    uploaded_files = st.file_uploader("배당표 업로드 (여러 파일 선택 가능)", type=['xlsx'], accept_multiple_files=True)
    
    if uploaded_files:
        summary_data = []
        
        for file in uploaded_files:
            df, duplicate_issues, first_year_electives = parse_and_audit(file)
            
            if df is not None:
                course_credits, total_credits, sem_totals = calculate_credits(df)
                validation_results = validate_curriculum(df, course_credits, sem_totals, duplicate_issues)
                
                has_error = any(status == "실패" for _, _, status in validation_results)
                
                summary_row = {
                    "학교/파일명": file.name,
                    "교과 이수 학점": course_credits,
                    "총 이수 학점": total_credits,
                    "총평": "⚠️ 점검 필요" if has_error else "✅ 정상",
                }
                for title, msg, status in validation_results:
                    summary_row[title] = msg
                summary_data.append(summary_row)
                
                with st.expander(f"🏫 {file.name} 검토 리포트 ({'⚠️ 확인 요망' if has_error else '✅ 정상'})", expanded=False):
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("총 이수 학점", f"{total_credits} 학점", delta="충족" if total_credits>=192 else "미달", delta_color="normal" if total_credits>=192 else "inverse")
                    col2.metric("교과 이수 학점", f"{course_credits} 학점", delta="충족" if course_credits>=174 else "미달", delta_color="normal" if course_credits>=174 else "inverse")
                    col3.metric("창체", "18 학점")
                    col4.metric("데이터 오류", f"{len(duplicate_issues)} 건", delta="수정 요망" if duplicate_issues else "정상", delta_color="inverse")
                    
                    st.divider()
                    
                    t1, t2 = st.tabs(["📋 자율점검 상세 결과", "🚨 이슈 항목"])
                    
                    with t1:
                        for title, message, status in validation_results:
                            if status == "성공":
                                st.success(f"**{title}**: {message}")
                            elif status == "경고":
                                st.warning(f"**{title}**: {message}")
                            else:
                                st.error(f"**{title}**: {message}")
                                
                    with t2:
                        if duplicate_issues:
                            st.error("❌ **동일 과목이 학년별로 다른 학점으로 편성됨**")
                            for sub, credits in duplicate_issues.items():
                                st.write(f"- {sub}: {', '.join(map(str, credits))}학점 이중 편성")
                        else:
                            st.write("✅ 데이터 기입 상의 중복/충돌 오류가 없습니다.")
                            
                        if first_year_electives:
                            st.info("💡 **1학년 일반선택 과목 편성 지양** (학교 사정에 따른 예외 수용 가능)")
                            st.write(f"- 대상 과목: {', '.join(first_year_electives)}")

        if summary_data:
            st.markdown("### 📥 일괄 검토 결과 다운로드")
            summary_df = pd.DataFrame(summary_data)
            excel_data = to_excel(summary_df)
            
            st.download_button(
                label="📊 전체 학교 검토 결과 다운로드 (.xlsx)",
                data=excel_data,
                file_name="교육과정_일괄검토_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

if __name__ == '__main__':
    main()
