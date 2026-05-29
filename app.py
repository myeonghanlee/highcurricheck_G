import streamlit as st
import pandas as pd
import numpy as np

# --- 1. 데이터 파싱 함수 ---
def parse_curriculum(file):
    try:
        # '2026입학생' 시트 로드 (헤더 없음)
        df_raw = pd.read_excel(file, sheet_name='2026입학생', header=None)
        
        # --- [수정된 부분] 헤더 병합 및 중복 방지 처리 ---
        # 1. 윗줄(1학년, 2학년 등) 병합 셀 처리를 위해 빈칸을 앞의 값으로 채움(ffill)
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
            
        # 2. 컬럼명 중복 제거 (예: 빈칸, 빈칸_1, 2학기, 2학기_1 방지)
        headers = []
        seen = {}
        for h in raw_headers:
            if h in seen:
                seen[h] += 1
                headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 0
                headers.append(h)
        # --------------------------------------------------
        
        # 7행부터 실제 데이터
        df = pd.DataFrame(df_raw.values[6:], columns=headers)
        
        # 병합된 셀(NaN) 채우기 (Forward Fill)
        df['구분'] = df['구분'].replace('', np.nan).ffill()
        
        # 파일마다 '교과(군)' 띄어쓰기가 다를 수 있어 포함 문자로 유연하게 찾기
        subject_col = [col for col in df.columns if '교과' in col][0]
        df[subject_col] = df[subject_col].replace('', np.nan).ffill()
        
        # 불필요한 빈 행 제거
        course_col = [col for col in df.columns if '과목' in col and '유형' not in col and '개설' not in col][0]
        df = df.dropna(subset=[course_col])
        
        # 운영학점 숫자형 변환
        credit_col = [col for col in df.columns if '운영' in col and '학점' in col][0]
        df[credit_col] = pd.to_numeric(df[credit_col], errors='coerce').fillna(0)
        
        # 뒤의 검토 함수(validate_curriculum)가 잘 작동하도록 핵심 컬럼명 강제 통일
        df = df.rename(columns={subject_col: '교과(군)', course_col: '과목', credit_col: '운영학점'})
        
        return df
    
    except Exception as e:
        st.error(f"파일을 파싱하는 중 오류가 발생했습니다: {e}")
        return None

# --- 2. 검토 함수 ---
def validate_curriculum(df):
    results = []
    
    # [검토 1] 교과 총 이수 학점 검토 (기준: 174학점 이상)
    total_credits = df['운영학점'].sum()
    if total_credits >= 174:
        results.append(("✅ 총 교과 이수 학점", f"총 {total_credits}학점으로 기준(174학점 이상)을 충족합니다.", "성공"))
    else:
        results.append(("❌ 총 교과 이수 학점", f"총 {total_credits}학점으로 기준(174학점)에 미달합니다.", "실패"))
        
    # [검토 2] 기초 교과(국수영) 편중 방지 (기준: 81학점 초과 금지)
    basic_subjects = ['국어', '수학', '영어', '국 어', '수 학', '영 어']
    basic_credits = df[df['교과(군)'].isin(basic_subjects)]['운영학점'].sum()
    
    if basic_credits <= 81:
        results.append(("✅ 기초 교과(국수영) 편중", f"총 {basic_credits}학점으로 기준(81학점 이하)을 충족합니다.", "성공"))
    else:
        results.append(("❌ 기초 교과(국수영) 편중", f"총 {basic_credits}학점으로 기준(81학점)을 초과했습니다.", "실패"))

    return results

# --- 3. 웹앱 메인 화면 ---
def main():
    st.set_page_config(page_title="고등학교 교육과정 자동 검토기", layout="wide")
    st.title("🏫 고등학교 교육과정 자동 검토 앱")
    st.markdown("교육청 표준 양식인 **'2026학년도 입학생 교육과정 학점 배당표'** 엑셀 파일을 업로드해주세요.")
    
    uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=['xlsx'])
    
    if uploaded_file is not None:
        st.info("파일을 성공적으로 불러왔습니다. 데이터 분석을 시작합니다...")
        
        df = parse_curriculum(uploaded_file)
        
        if df is not None:
            st.subheader("📊 파싱된 데이터 미리보기")
            st.dataframe(df.head(10)) 
            
            st.subheader("📋 자율 점검표 검토 결과")
            validation_results = validate_curriculum(df)
            
            for title, message, status in validation_results:
                if status == "성공":
                    st.success(f"**{title}** : {message}")
                else:
                    st.error(f"**{title}** : {message}")

if __name__ == '__main__':
    main()
