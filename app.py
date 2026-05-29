import streamlit as st
import pandas as pd
import numpy as np
import re

# --- 1. 데이터 파싱 및 클렌징 함수 ---
def parse_curriculum(file):
    try:
        # '2026입학생' 시트 로드 (헤더 없음)
        df_raw = pd.read_excel(file, sheet_name='2026입학생', header=None)
        
        # 1. 윗줄 병합 셀 처리를 위해 빈칸을 앞의 값으로 채움 (Forward Fill)
        row2 = df_raw.iloc[2].replace(r'^\s*$', np.nan, regex=True).ffill().fillna('')
        row3 = df_raw.iloc[3].fillna('')
        
        raw_headers = []
        for a, b in zip(row2, row3):
            a_str = str(a).strip().replace('\n', '')
            b_str = str(b).strip().replace('\n', '')
            header = f"{a_str}_{b_str}".strip('_')
            if not header:
                header = "빈칸"
            raw_headers.append(header)
            
        # 2. 컬럼명 중복 제거 (PyArrow 에러 방지)
        headers = []
        seen = {}
        for h in raw_headers:
            if h in seen:
                seen[h] += 1
                headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 0
                headers.append(h)
        
        # 7행부터 실제 데이터 데이터프레임 생성
        df = pd.DataFrame(df_raw.values[6:], columns=headers)
        
        # 3. 텍스트 정규화 및 병합 셀 빈칸 채우기
        df['구분'] = df['구분'].replace('', np.nan).ffill().str.replace('\n', '').str.strip()
        
        subject_col = [col for col in df.columns if '교과' in col][0]
        df[subject_col] = df[subject_col].replace('', np.nan).ffill().str.replace('\n', ' ').str.strip()
        
        course_col = [col for col in df.columns if '과목' in col and '유형' not in col and '개설' not in col][0]
        df = df.dropna(subset=[course_col])
        
        credit_col = [col for col in df.columns if '운영' in col and '학점' in col][0]
        df[credit_col] = pd.to_numeric(df[credit_col], errors='coerce').fillna(0)
        
        # 컬럼명 표준화 정의
        df = df.rename(columns={subject_col: '교과(군)', course_col: '과목', credit_col: '운영학점'})
        
        return df
    
    except Exception as e:
        st.error(f"파일을 분석하는 중 오류가 발생했습니다: {e}")
        return None

# --- 2. 교육과정 자율점검표 종합 검토 함수 ---
def validate_curriculum(df):
    results = []
    semester_cols = ['1학년_1학기', '1학년_2학기', '2학년_1학기', '2학년_2학기', '3학년_1학기', '3학년_2학기']
    
    # [개선 사항 1] 실제 학생 1인의 학기별/총 이수 학점 정밀 계산 (택N 반영)
    designated_mask = df['구분'].str.contains('학교지정', na=False)
    designated_total = df[designated_mask]['운영학점'].sum()
    
    # 학기별 이수 학점 추적용 딕셔너리
    semester_student_totals = {sem: 0.0 for sem in semester_cols}
    choice_total_credits = 0
    
    for sem in semester_cols:
        if sem in df.columns:
            # A. 학교지정 과목 중 해당 학기 운영 학점 합산
            des_sem_sum = pd.to_numeric(df[designated_mask][sem], errors='coerce').fillna(0).sum()
            semester_student_totals[sem] += des_sem_sum
            
            # B. 학생선택 과목 중 해당 학기 '택N' 문구에서 실제 이수 학점 추출
            unique_choice_strings = df[~designated_mask][sem].dropna().unique()
            sem_choice_sum = 0
            for val in unique_choice_strings:
                match = re.search(r'(\d+)\s*\(택', str(val))
                if match:
                    sem_choice_sum += int(match.group(1))
            
            semester_student_totals[sem] += sem_choice_sum
            choice_total_credits += sem_choice_sum

    actual_total_credits = designated_total + choice_total_credits

    # ----------------------------------------------------
    # 기준 1: 총 이수 학점 검토 (교과 174학점 이상)
    # ----------------------------------------------------
    if actual_total_credits >= 174:
        results.append(("✅ 총 교과 이수 학점", f"실제 학생 이수 기준 총 {actual_total_credits}학점으로 지침(174학점 이상)을 충족합니다.", "성공"))
    else:
        results.append(("❌ 총 교과 이수 학점", f"실제 학생 이수 기준 총 {actual_total_credits}학점으로 지침(174학점)에 미달합니다. 편성을 재확인하세요.", "실패"))

    # ----------------------------------------------------
    # 기준 2: 교과(군)별 필수 이수 학점 (84학점) 검토
    # ----------------------------------------------------
    req_map = {'국어': 8, '수학': 8, '영어': 8, '과학': 10, '체육': 10, '예술': 10}
    mandatory_failed = []
    
    # 학교지정 기준으로 기초적인 필수 이수 학점 충족 체크
    df_des = df[designated_mask]
    for k, v in req_map.items():
        current_sum = df_des[df_des['교과(군)'].str.contains(k, na=False)]['운영학점'].sum()
        if current_sum < v:
            # 선택과목에 포함되어 있을 수 있으므로 알림 형태로 전달
            mandatory_failed.append(f"{k}(기준 {v}학점 / 지정 {int(current_sum)}학점)")
            
    if not mandatory_failed:
        results.append(("✅ 교과(군)별 필수 이수 학점", "국어(8), 수학(8), 영어(8), 과학(10), 체육(10), 예술(10) 등 기본 필수 학점이 학교지정만으로 모두 충족됩니다.", "성공"))
    else:
        results.append(("⚠️ 교과(군)별 필수 이수 학점(확인 필요)", f"일부 교과군이 학교지정만으로는 부족합니다. 학생 선택군에서 충족되는지 확인하세요: {', '.join(mandatory_failed)}", "경고"))

    # ----------------------------------------------------
    # 기준 3: 기초 교과(국수영) 편중 방지 (81학점 초과 금지)
    # ----------------------------------------------------
    basic_credits = df_des[df_des['교과(군)'].str.contains('국어|수학|영어', na=False)]['운영학점'].sum()
    if basic_credits <= 81:
        results.append(("✅ 기초 교과(국수영) 편중 방지", f"학교지정 기준 국수영 학점 총합이 {int(basic_credits)}학점으로 상한선(81학점) 이하입니다.", "성공"))
    else:
        results.append(("❌ 기초 교과(국수영) 편중 과다", f"학교지정 국수영 학점 총합이 {int(basic_credits)}학점으로 상한선(81학점)을 초과했습니다.", "실패"))

    # ----------------------------------------------------
    # 기준 4: 체육 교과 매 학기 편성 검토
    # ----------------------------------------------------
    pe_df = df[df['교과(군)'].str.contains('체육', na=False)]
    pe_sem_check = []
    for sem in semester_cols:
        if sem in pe_df.columns:
            has_pe = pe_df[sem].dropna().astype(str).str.strip().str.len().gt(0).any()
            if not has_pe:
                pe_sem_check.append(sem.replace('_', ' '))
                
    if not pe_sem_check:
        results.append(("✅ 체육 교과 매 학기 편성", "1학년 1학기부터 3학년 2학기까지 매 학기 체육 과목이 누락 없이 편성되었습니다.", "성공"))
    else:
        results.append(("❌ 체육 교과 편성 누락", f"다음 학기에 체육 과목이 편성되지 않았습니다: {', '.join(pe_sem_check)}", "실패"))

    # ----------------------------------------------------
    # 기준 5: 학기당 이수 학점 격차 검토 (5학점 이내)
    # ----------------------------------------------------
    actual_sem_vals = [v for v in semester_student_totals.values() if v > 0]
    if actual_sem_vals:
        credit_diff = max(actual_sem_vals) - min(actual_sem_vals)
        if credit_diff <= 5:
            results.append(("✅ 학기 간 이수 학점 격차", f"최대 학기({max(actual_sem_vals)}학점)와 최소 학기({min(actual_sem_vals)}학점)의 차이가 {credit_diff}학점으로 기준(5학점 이내)을 만족합니다.", "성공"))
        else:
            results.append(("❌ 학기 간 이수 학점 격차 초과", f"학기 간 학점 차이가 {credit_diff}학점입니다. 격차가 5학점을 초과하여 균형 편성에 위배됩니다.", "실패"))

    # ----------------------------------------------------
    # 기준 6: 동일 과목 동일 학점 검토
    # ----------------------------------------------------
    duplicate_errors = []
    subject_groups = df.groupby('과목')['운영학점'].nunique()
    for sub, count in subject_groups.items():
        if count > 1:
            diff_credits = df[df['과목'] == sub]['운영학점'].unique()
            duplicate_errors.append(f"{sub}({', '.join(map(str, diff_credits))}학점)")
            
    if not duplicate_errors:
        results.append(("✅ 동일 과목 동일 학점 준수", "동일한 과목이 서로 다른 학점 수로 이중 편성된 사례가 없습니다.", "성공"))
    else:
        results.append(("❌ 동일 과목 이중 학점 오류", f"동일 과목의 학점이 다르게 입력되었습니다: {', '.join(duplicate_errors)}", "실패"))

    # ----------------------------------------------------
    # 기준 7 & 8: 지정 과목(한국사, 과학탐구실험) 고정 학점 검토
    # ----------------------------------------------------
    history_err = []
    sci_exp_err = []
    for idx, row in df.iterrows():
        sub_name = str(row['과목']).strip()
        if '한국사' in sub_name and row['운영학점'] != 3:
            history_err.append(f"{sub_name}({int(row['운영학점'])}학점)")
        if '과학탐구실험' in sub_name and row['운영학점'] != 1:
            sci_exp_err.append(f"{sub_name}({int(row['운영학점'])}학점)")
            
    if not history_err:
        results.append(("✅ 한국사 고정 학점 준수", "한국사 1, 2 과목이 지정 기준대로 각각 3학점으로 편성되었습니다.", "성공"))
    else:
        results.append(("❌ 한국사 학점 오류", f"한국사는 3학점 고정 편성이 원칙입니다: {', '.join(history_err)}", "실패"))
        
    if not sci_exp_err:
        results.append(("✅ 과학탐구실험 학점 준수", "과학탐구실험 1, 2 과목이 지정 기준대로 각각 1학점으로 편성되었습니다.", "성공"))
    else:
        results.append(("❌ 과학탐구실험 학점 오류", f"과학탐구실험은 1학점 고정 편성이 원칙입니다: {', '.join(sci_exp_err)}", "실패"))

    return results

# --- 3. 웹앱 화면 레이아웃 구성 ---
def main():
    st.set_page_config(page_title="고등학교 교육과정 자동 검토기", layout="wide")
    st.title("🏫 고등학교 교육과정 편성·운영 자동 검토 시스템")
    st.markdown("교육청 표준 서식인 **'교육과정 학점 배당표'**를 분석하여 자율점검표 기준 검증을 일괄 수행합니다.")
    
    uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=['xlsx'])
    
    if uploaded_file is not None:
        st.success("파일을 성공적으로 업로드했습니다. 정밀 데이터 검증을 진행합니다.")
        
        df = parse_curriculum(uploaded_file)
        
        if df is not None:
            st.subheader("📊 정제된 교육과정 데이터 테이블 (상위 15개 행)")
            st.dataframe(df[['구분', '교과(군)', '과목', '운영학점']].head(15), use_container_width=True)
            
            st.subheader("📋 고등학교 교육과정 자율점검표 검토 결과")
            validation_results = validate_curriculum(df)
            
            # 검토 항목 결과 대시보드 출력
            for title, message, status in validation_results:
                if status == "성공":
                    st.success(f"**{title}** \n{message}")
                elif status == "경고":
                    st.warning(f"**{title}** \n{message}")
                else:
                    st.error(f"**{title}** \n{message}")

if __name__ == '__main__':
    main()
