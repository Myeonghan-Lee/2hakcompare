import streamlit as st
import pandas as pd
import re
import io

# -----------------------------------------------------------------------------
# 1. 공통 유틸리티 함수
# -----------------------------------------------------------------------------

def load_data(uploaded_file):
    """파일 로드 (CSV, Excel)"""
    file_ext = uploaded_file.name.split('.')[-1].lower()
    try:
        if file_ext == 'csv':
            return pd.read_csv(uploaded_file, header=None)
        elif file_ext in ['xlsx', 'xls']:
            return pd.read_excel(uploaded_file, header=None, engine='openpyxl')
        else:
            return None
    except Exception as e:
        st.error(f"파일 오류 ({uploaded_file.name}): {e}")
        return None

def extract_grade_class(df_raw):
    """학년 반 추출"""
    limit = min(20, len(df_raw))
    for i in range(limit):
        row_values = df_raw.iloc[i].astype(str).values
        for val in row_values:
            match = re.search(r"(\d+)학년\s*(\d+)반", val)
            if match:
                return match.group(0)
    return "미상"

def detect_file_type(df_raw):
    """파일 유형 감지 (행특 / 세특 / 창체)"""
    limit = min(20, len(df_raw))
    text_sample = df_raw.iloc[:limit].astype(str).to_string()
    
    if "창의적" in text_sample and ("체험활동" in text_sample or "자율" in text_sample):
        return "CHANG"
    elif "행 동 특 성" in text_sample or "행동특성" in text_sample or "종합의견" in text_sample:
        return "HANG"
    elif "세부능력" in text_sample or "특기사항" in text_sample or "과 목" in text_sample:
        return "KYO"
    else:
        return "UNKNOWN"

# -----------------------------------------------------------------------------
# 2. 데이터 처리 로직 (행특 / 세특 / 창체)
# -----------------------------------------------------------------------------

def process_hang(df_raw, grade_class):
    header_idx = -1
    for i, row in df_raw.iterrows():
        row_str = row.astype(str).values
        if any('번' in s and '호' in s for s in row_str) and any('성' in s and '명' in s for s in row_str):
            header_idx = i
            break
            
    if header_idx == -1: return None

    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = df_raw.iloc[header_idx].astype(str).str.replace(" ", "")
    
    rename_map = {}
    for col in df.columns:
        if '번호' in col: rename_map[col] = '번호'
        elif '행동특성' in col: rename_map[col] = '내용'
        elif '종합의견' in col: rename_map[col] = '내용'
    df = df.rename(columns=rename_map)
    
    if '번호' not in df.columns or '내용' not in df.columns: return None
        
    df['번호'] = pd.to_numeric(df['번호'], errors='coerce')
    df = df[df['내용'].notna()]
    df = df[~df['내용'].astype(str).str.contains('행 동 특 성', na=False)]
    df = df[~df['내용'].astype(str).str.contains('종 합 의 견', na=False)]
    
    df['번호'] = df['번호'].ffill()
    df = df.dropna(subset=['번호'])
    df['번호'] = df['번호'].astype(int) # [요구사항 3] 번호 정수형 변환
    
    df_grouped = df.groupby('번호')['내용'].apply(lambda x: ' '.join(x.astype(str))).reset_index()
    
    df_grouped['학년 반'] = grade_class
    df_grouped['학기'] = ''
    df_grouped['과목/영역'] = '행동특성'
    df_grouped['시수'] = ''
    
    return df_grouped

def process_kyo(df_raw, grade_class):
    header_idx = -1
    for i, row in df_raw.iterrows():
        row_str = row.astype(str).values
        if any('과' in s and '목' in s for s in row_str) and any('세부능력' in s for s in row_str):
            header_idx = i
            break
            
    if header_idx == -1: return None
        
    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = df_raw.iloc[header_idx].astype(str).str.replace(" ", "")
    
    rename_map = {}
    for col in df.columns:
        if '과목' in col: rename_map[col] = '과목/영역'
        elif '학기' in col: rename_map[col] = '학기'
        elif '번호' in col: rename_map[col] = '번호'
        elif '세부능력' in col: rename_map[col] = '내용'
        elif '특기사항' in col: rename_map[col] = '내용'
    df = df.rename(columns=rename_map)
    
    if '내용' not in df.columns or '과목/영역' not in df.columns: return None

    df['번호'] = pd.to_numeric(df['번호'], errors='coerce')
    df = df[df['과목/영역'] != '과 목']
    df = df[df['과목/영역'] != '과목']
    df['번호'] = df['번호'].ffill()
    df['과목/영역'] = df['과목/영역'].ffill()
    df['학기'] = df['학기'].ffill()
    
    df = df.dropna(subset=['번호', '내용'])
    df['번호'] = df['번호'].astype(int) # [요구사항 3] 번호 정수형 변환
    
    df_grouped = df.groupby(['번호', '학기', '과목/영역'])['내용'].apply(lambda x: ' '.join(x.astype(str))).reset_index()
    
    df_grouped['학년 반'] = grade_class
    df_grouped['시수'] = '' 
    
    return df_grouped

def process_chang(df_raw, grade_class):
    header_idx = -1
    for i, row in df_raw.iterrows():
        row_str = row.astype(str).values
        if any('영' in s and '역' in s for s in row_str) and any('시' in s and '간' in s for s in row_str):
            header_idx = i
            break
            
    if header_idx == -1: return None
    
    cols = df_raw.iloc[header_idx].fillna('').astype(str).values.tolist()
    
    if header_idx > 0:
        upper_row = df_raw.iloc[header_idx - 1].fillna('').astype(str).values.tolist()
        for i in range(len(cols)):
            if cols[i].strip() == '' or cols[i].lower() == 'nan':
                if i < len(upper_row) and upper_row[i].strip() != '' and upper_row[i].lower() != 'nan':
                    cols[i] = upper_row[i]
    
    cols = [c.replace(" ", "") for c in cols]
    
    df = df_raw.iloc[header_idx+1:].copy()
    df.columns = cols
    
    rename_map = {}
    for col in df.columns:
        if '번호' in col: rename_map[col] = '번호'
        elif '영역' in col: rename_map[col] = '과목/영역'
        elif '시간' in col: rename_map[col] = '시수'
        elif '특기사항' in col: rename_map[col] = '내용'
    
    df = df.rename(columns=rename_map)
    
    if '번호' not in df.columns or '내용' not in df.columns or '과목/영역' not in df.columns:
        return None

    df['번호'] = pd.to_numeric(df['번호'], errors='coerce')
    df = df[df['과목/영역'] != '영 역']
    df = df[df['과목/영역'] != '영역']
    
    df['번호'] = df['번호'].ffill()
    df['과목/영역'] = df['과목/영역'].ffill()
    df['시수'] = df['시수'].ffill()
    
    df = df.dropna(subset=['번호'])
    df['번호'] = df['번호'].astype(int) # [요구사항 3] 번호 정수형 변환
    
    df = df[df['내용'].astype(str) != '희망분야']
    df = df[~df['내용'].astype(str).str.contains('희망분야', na=False)]
    df = df.dropna(subset=['내용'])

    df_grouped = df.groupby(['번호', '과목/영역', '시수'])['내용'].apply(lambda x: ' '.join(x.astype(str))).reset_index()
    
    df_grouped['학년 반'] = grade_class
    df_grouped['학기'] = '' 
    
    return df_grouped

# -----------------------------------------------------------------------------
# 3. 중복 탐지 로직 (색상 매핑 추가)
# -----------------------------------------------------------------------------

# 시각적으로 구분이 잘 되는 파스텔톤 색상 팔레트
COLOR_PALETTE = [
    '#ffadad', '#ffd6a5', '#fdffb6', '#caffbf', '#9bf6ff', '#a0c4ff', '#bdb2ff', '#ffc6ff', '#fffffc'
]

@st.cache_data
def detect_duplicates(df):
    """복붙(중복) 문장 탐지 및 색상 매핑"""
    if df.empty: return df
    
    df['중복여부'] = False
    df['복붙 의심 문장'] = ''
    df['색상'] = '' # 엑셀/웹 스타일링을 위한 숨김 컬럼
    df['과목/영역'] = df['과목/영역'].fillna('기타')
    
    color_idx = 0
    duplicate_color_map = {}
    
    for subject, group in df.groupby('과목/영역'):
        if len(group) < 2: continue
        
        sentence_counts = {}
        # [개선] 마침표가 누락된 경우도 대비해 줄바꿈 및 문장부호로 분리 후 10글자 이상만 추출
        for idx, row in group.iterrows():
            content = str(row['내용'])
            sentences = [s.strip() for s in re.split(r'[.!?\n]+', content) if len(s.strip()) >= 10]
            for s in sentences:
                sentence_counts[s] = sentence_counts.get(s, 0) + 1
        
        duplicate_sentences = {s for s, count in sentence_counts.items() if count > 1}
        
        # 중복 문장마다 고유한 색상 부여
        for dup_sent in duplicate_sentences:
            if dup_sent not in duplicate_color_map:
                duplicate_color_map[dup_sent] = COLOR_PALETTE[color_idx % len(COLOR_PALETTE)]
                color_idx += 1
        
        for idx, row in group.iterrows():
            content = str(row['내용'])
            sentences = [s.strip() for s in re.split(r'[.!?\n]+', content) if len(s.strip()) >= 10]
            found_duplicates = [s for s in sentences if s in duplicate_sentences]
            
            if found_duplicates:
                df.at[idx, '중복여부'] = True
                unique_dupes = list(set(found_duplicates))
                df.at[idx, '복붙 의심 문장'] = " / ".join(unique_dupes)
                # 첫 번째 발견된 중복 문장의 색상을 해당 행의 테마색으로 지정
                df.at[idx, '색상'] = duplicate_color_map[unique_dupes[0]]

    # [요구사항 1] 열 표시 순서 정렬
    ordered_cols = ['학년 반', '학기', '과목/영역', '번호', '시수', '내용', '복붙 의심 문장', '중복여부', '색상']
    # 혹시 누락된 컬럼이 있을 수 있으므로 교집합으로 순서 재정렬
    final_cols = [c for c in ordered_cols if c in df.columns] 
    return df[final_cols]

def style_dataframe(df_to_style):
    """Pandas Styler를 사용해 과목/영역, 내용, 복붙의심문장 컬럼에 배경색 적용"""
    def row_style(row):
        styles = [''] * len(row)
        if row.get('중복여부', False) and row.get('색상', '') != '':
            bg_color = f"background-color: {row['색상']}; color: black;"
            
            # [요구사항 2] 과목/영역, 내용, 복붙 의심 문장에만 색상 적용
            for target_col in ['과목/영역', '내용', '복붙 의심 문장']:
                if target_col in row.index:
                    styles[row.index.get_loc(target_col)] = bg_color
        return styles

    # UI 및 엑셀 출력 시 불필요한 보조 컬럼은 숨김 처리할 수 있도록 셋업
    display_cols = [c for c in df_to_style.columns if c not in ['중복여부', '색상']]
    return df_to_style.style.apply(row_style, axis=1), display_cols

@st.cache_data
def to_excel_with_style(df):
    """엑셀 스타일링 및 저장 (캐싱 적용)"""
    output = io.BytesIO()
    styler, save_cols = style_dataframe(df)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        styler.to_excel(writer, index=False, columns=save_cols, sheet_name='정리결과')
        worksheet = writer.sheets['정리결과']
        for idx, col in enumerate(save_cols):
            width = 50 if '내용' in col or '문장' in col else 12
            worksheet.column_dimensions[chr(65 + idx)].width = width
            
    return output.getvalue()

# -----------------------------------------------------------------------------
# 4. 메인 앱 UI
# -----------------------------------------------------------------------------
st.set_page_config(page_title="학생부 점검 도우미", layout="wide")

st.title("🏫 학생부 점검 도우미")
st.markdown("""
**지원내용:** 행특, 세특(교과), 창체(자율/진로)

**기능:**
  1. xlsx_data 파일 다운로드 및 업로드 시 **자동 분류 및 정리**
  2. **복붙 의심 문장 색상 분류 표시** (같은 중복 문장끼리 같은 색상)
""")

# 세션 상태(Session State)를 활용해 엑셀 다운로드 시 데이터가 증발하거나 재연산되는 현상 방지
if 'final_df' not in st.session_state:
    st.session_state.final_df = None

uploaded_files = st.file_uploader(
    "처리할 파일들을 모두 올려주세요", 
    accept_multiple_files=True,
    type=['xlsx', 'xls', 'csv']
)

# 파일이 업로드되었고, 버튼을 누르거나 파일이 변경되었을 때 실행
if uploaded_files:
    if st.button("파일 분석 시작", type="primary"):
        all_results = []
        
        with st.status("파일 분석 및 처리 중...", expanded=True) as status:
            for file in uploaded_files:
                df_raw = load_data(file)
                if df_raw is None:
                    st.error(f"{file.name}: 읽기 실패")
                    continue
                    
                grade_class = extract_grade_class(df_raw)
                file_type = detect_file_type(df_raw)
                
                processed_df = None
                type_label = ""
                
                if file_type == 'HANG':
                    processed_df = process_hang(df_raw, grade_class)
                    type_label = "행동특성"
                elif file_type == 'KYO':
                    processed_df = process_kyo(df_raw, grade_class)
                    type_label = "세부능력"
                elif file_type == 'CHANG':
                    processed_df = process_chang(df_raw, grade_class)
                    type_label = "창의적체험"
                else:
                    st.warning(f"⚠️ {file.name}: 알 수 없는 형식 (건너뜀)")
                    continue
                    
                if processed_df is not None and not processed_df.empty:
                    all_results.append(processed_df)
                    st.write(f"✅ {file.name} ({type_label} / {grade_class}) - {len(processed_df)}명 처리")
                else:
                    st.warning(f"⚠️ {file.name}: 데이터 추출 실패")

            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)
                final_df = final_df.sort_values(by=['과목/영역', '번호'])
                final_df = detect_duplicates(final_df)
                st.session_state.final_df = final_df # 세션에 저장
            else:
                st.session_state.final_df = None

            status.update(label="모든 파일 처리 완료!", state="complete", expanded=False)

# 세션에 저장된 결과가 있으면 출력 및 다운로드 버튼 제공
if st.session_state.final_df is not None:
    final_df = st.session_state.final_df
    
    st.divider()
    st.subheader("📊 결과 미리보기")
    
    styler, display_cols = style_dataframe(final_df)
    
    st.dataframe(
        styler,
        column_order=display_cols,
        column_config={
            "번호": st.column_config.NumberColumn("번호", format="%d"),
            "시수": st.column_config.TextColumn("시수", width="small"),
            "복붙 의심 문장": st.column_config.TextColumn("⚠️ 복붙 의심 문장", width="large")
        },
        use_container_width=True,
        hide_index=True
    )
    
    excel_data = to_excel_with_style(final_df)
    
    st.download_button(
        label="📥 통합 엑셀 파일 다운로드 (.xlsx)",
        data=excel_data,
        file_name="생기부_통합_정리결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
