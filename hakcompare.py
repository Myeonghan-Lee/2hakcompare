import streamlit as st
import pandas as pd
import re
import io

# -----------------------------------------------------------------------------
# 0. 페이지 기본 설정 (항상 최상단에 위치)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="학생부 점검 도우미", layout="wide")

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
        return "CHANG" # 창의적 체험활동
    elif "행 동 특 성" in text_sample or "행동특성" in text_sample or "종합의견" in text_sample:
        return "HANG" # 행동특성
    elif "세부능력" in text_sample or "특기사항" in text_sample or "과 목" in text_sample:
        return "KYO" # 세부능력(교과)
    else:
        return "UNKNOWN"

# -----------------------------------------------------------------------------
# 2. 데이터 처리 로직 (행특 / 세특 / 창체)
# -----------------------------------------------------------------------------

def process_hang(df_raw, grade_class):
    """행동특성 처리"""
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
    
    # 필수 컬럼 확인
    if '번호' not in df.columns or '내용' not in df.columns: return None
        
    df['번호'] = pd.to_numeric(df['번호'], errors='coerce')
    df = df[df['내용'].notna()]
    df = df[~df['내용'].str.contains('행 동 특 성', na=False)]
    df = df[~df['내용'].str.contains('종 합 의 견', na=False)]
    
    df['번호'] = df['번호'].ffill()
    df = df.dropna(subset=['번호'])
    
    df_grouped = df.groupby('번호')['내용'].apply(lambda x: ' '.join(x.astype(str))).reset_index()
    
    df_grouped['학년 반'] = grade_class
    df_grouped['학기'] = ''
    df_grouped['과목/영역'] = '행동특성'
    df_grouped['시수'] = ''
    
    return df_grouped[['학년 반', '번호', '학기', '과목/영역', '시수', '내용']]

def process_kyo(df_raw, grade_class):
    """세부능력(교과) 처리"""
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
    
    df_grouped = df.groupby(['번호', '학기', '과목/영역'])['내용'].apply(lambda x: ' '.join(x.astype(str))).reset_index()
    
    df_grouped['학년 반'] = grade_class
    df_grouped['시수'] = '' 
    
    return df_grouped[['학년 반', '번호', '학기', '과목/영역', '시수', '내용']]

def process_chang(df_raw, grade_class):
    """창의적 체험활동(자율/진로) 처리"""
    header_idx = -1
    for i, row in df_raw.iterrows():
        row_str = row.astype(str).values
        if any('영' in s and '역' in s for s in row_str) and any('시' in s and '간' in s for s in row_str):
            header_idx = i
            break
            
    if header_idx == -1: return None
    
    # 2단 헤더 병합 로직
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
    
    df = df[df['내용'].astype(str) != '희망분야']
    df = df[~df['내용'].astype(str).str.contains('희망분야', na=False)]
    df = df.dropna(subset=['내용'])

    df_grouped = df.groupby(['번호', '과목/영역', '시수'])['내용'].apply(lambda x: ' '.join(x.astype(str))).reset_index()
    
    df_grouped['학년 반'] = grade_class
    df_grouped['학기'] = '' 
    
    return df_grouped[['학년 반', '번호', '학기', '과목/영역', '시수', '내용']]

def detect_duplicates(df):
    """복붙(중복) 문장 탐지 및 그룹별 색상 할당"""
    sentence_pattern = re.compile(r'[^.!?]+[.!?]')
    df['중복여부'] = False
    df['비고(중복문장)'] = ''
    df['중복색상'] = '' 
    
    # 🎨 파스텔톤 컬러 팔레트
    color_palette = [
        '#ffb3ba', '#ffdfba', '#ffffba', '#baffc9', '#bae1ff', 
        '#e8baff', '#ffbaff', '#ffc4e1', '#e2f0cb', '#ffcfd2',
        '#d4f0f0', '#f3e8ff', '#ffebd6', '#e6fffa', '#ffe6f2'
    ]
    
    df['과목/영역'] = df['과목/영역'].fillna('기타')
    
    for subject, group in df.groupby('과목/영역'):
        if len(group) < 2: continue
        
        sentence_counts = {}
        for idx, row in group.iterrows():
            content = str(row['내용'])
            sentences = [s.strip() for s in sentence_pattern.findall(content)]
            for s in sentences:
                if len(s) < 10: continue
                sentence_counts[s] = sentence_counts.get(s, 0) + 1
        
        duplicate_sentences = {s for s, count in sentence_counts.items() if count > 1}
        
        # 중복 문장별 고유 색상 매핑
        color_map = {}
        for i, dup_sent in enumerate(duplicate_sentences):
            color_map[dup_sent] = color_palette[i % len(color_palette)]
            
        for idx, row in group.iterrows():
            content = str(row['내용'])
            sentences = [s.strip() for s in sentence_pattern.findall(content)]
            found_duplicates = [s for s in sentences if s in duplicate_sentences]
            
            if found_duplicates:
                df.at[idx, '중복여부'] = True
                unique_dupes = list(set(found_duplicates))
                df.at[idx, '비고(중복문장)'] = " / ".join(unique_dupes)
                df.at[idx, '중복색상'] = color_map[unique_dupes[0]]

    return df

def to_excel_with_style(df):
    """엑셀 스타일링 및 저장 (특정 열만 색상 반영)"""
    output = io.BytesIO()
    save_cols = [c for c in df.columns if c not in ['중복여부', '중복색상']]
    
    def style_duplicate_excel(row):
        styles = [''] * len(row)
        if row.get('중복여부', False) and row.get('중복색상', ''):
            bg_color = row['중복색상']
            # 🎨 과목/영역, 번호, 내용에만 배경색 적용
            for col in ['과목/영역', '번호', '내용']:
                if col in row.index:
                    try:
                        idx = row.index.get_loc(col)
                        styles[idx] = f'background-color: {bg_color}'
                    except KeyError: pass
            
            # 비고(중복문장) 열은 빨간색 텍스트
            if '비고(중복문장)' in row.index:
                try:
                    note_idx = row.index.get_loc('비고(중복문장)')
                    styles[note_idx] = 'color: red;'
                except KeyError: pass
                
        return styles

    styler = df.style.apply(style_duplicate_excel, axis=1)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        styler.to_excel(writer, index=False, columns=save_cols, sheet_name='정리결과')
        worksheet = writer.sheets['정리결과']
        for idx, col in enumerate(save_cols):
            width = 50 if '내용' in col or '비고' in col else 12
            worksheet.column_dimensions[chr(65 + idx)].width = width
            
    return output.getvalue()

# -----------------------------------------------------------------------------
# 3. 메인 앱 UI
# -----------------------------------------------------------------------------

st.title("🏫 학생부 점검 도우미")
st.markdown("""
**지원내용:** 행특, 세특(교과), 창체(자율/진로)

**기능:**
  1. xlsx_data 파일 다운로드 및 업로드 시 **자동 분류 및 정리**
  2. **복붙 의심 문장 그룹별 다른 색상 표시 (과목, 번호, 내용 강조)**
""")

uploaded_files = st.file_uploader(
    "처리할 파일들을 모두 올려주세요", 
    accept_multiple_files=True,
    type=['xlsx', 'xls', 'csv']
)

if uploaded_files:
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

        status.update(label="모든 파일 처리 완료!", state="complete", expanded=False)

    if all_results:
        final_df = pd.concat(all_results, ignore_index=True)
        final_df = final_df.sort_values(by=['과목/영역', '번호'])
        final_df = detect_duplicates(final_df)
        
        # 🔢 번호를 정수형(int)으로 변환
        final_df['번호'] = pd.to_numeric(final_df['번호']).astype(int)
        
        # 📌 요청하신 컬럼 순서 지정
        ordered_cols = ['학년 반', '학기', '과목/영역', '번호', '시수', '내용', '비고(중복문장)', '중복여부', '중복색상']
        final_df = final_df[ordered_cols]
        
        st.divider()
        st.subheader("📊 결과 미리보기")
        
        # 🎨 웹 화면 스타일링 함수
        def highlight_row_web(row):
            styles = [''] * len(row)
            if row.get('중복여부', False) and row.get('중복색상', ''):
                bg_color = row['중복색상']
                for col in ['과목/영역', '번호', '내용']:
                    if col in row.index:
                        try:
                            idx = row.index.get_loc(col)
                            styles[idx] = f'background-color: {bg_color}'
                        except KeyError: pass
            return styles
            
        st.dataframe(
            final_df.style.apply(highlight_row_web, axis=1),
            column_config={
                "시수": st.column_config.TextColumn("시수", width="small"),
                "비고(중복문장)": st.column_config.TextColumn("⚠️ 복붙 의심 문장", width="medium"),
                "중복여부": None, # 화면에서 숨김
                "중복색상": None  # 화면에서 숨김
            },
            use_container_width=True
        )
        
        excel_data = to_excel_with_style(final_df)
        
        st.download_button(
            label="📥 통합 엑셀 파일 다운로드 (.xlsx)",
            data=excel_data,
            file_name="생기부_통합_정리결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("처리할 데이터가 없습니다.")
