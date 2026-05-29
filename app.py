import streamlit as st
import pandas as pd
import numpy as np

# --- 1. 데이터 파싱 함수 ---
def parse_curriculum(file):
    try:
        # '2026입학생' 시트 로드 (헤더 없음)
        df_raw = pd.read_excel(file, sheet_name='2026입학생', header=None)
        
        # 3행, 4행을 결합하여 컬럼명 생성
        row2 = df_raw.iloc[2].fillna('')
        row3 = df_raw.iloc[3].fillna('')
        headers = [f"{str(a).strip()}_{str(b).strip()}".strip('_').replace('\n', '') 
                   for a, b in zip(row2, row3)]
        
        # 7행부터 실제 데이터
        df = pd.DataFrame(df_raw.values[6:], columns=headers)
        
        # 병합된 셀(NaN) 채우기 (Forward Fill)
        df['구분'] = df['구분'].replace('', np.nan).ffill()
        df['교과(군)'] = df['교과(군)'].replace('', np.nan).ffill()
        
        # 불필요한 행 제거
        df = df.dropna(subset=['과목'])
        
        # 운영학점 숫자형 변환
        df['운영학점'] = pd.to_numeric(df['운영학점'], errors='coerce').fillna(0)
        
        return df
    
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
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
