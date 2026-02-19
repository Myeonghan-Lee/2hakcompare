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
    """파일 유형 감지 (행특 / 세특 / 창체) - 헤더 기반 정확한 판정"""
    limit = min(20, len(df_raw))
    
    # 방법: 실제 헤더 행의 키워드로 판정 (본문 내용 제외)
    for i in range(limit):
        row_str = " ".join(df_raw.iloc[i].astype(str).values)
        
        # 창체: 헤더에 "영역"과 "시간"/"시수"가 함께 있는 경우
        if ("영" in row_str and "역" in row_str) and ("시" in row_str and "간" in row_str):
            return "CHANG"
        
        # 행특: 헤더/제목에 "행동특성" 또는 "종합의견"이 있는 경우
        if "행 동 특 성" in row_str or "행동특성" in row_str or "종합의견" in row_str:
            return "HANG"
        
        # 세특: 헤더에 "과목"과 "세부능력"이 함께 있는 경우  
        if ("과" in row_str and "목" in row_str) and "세부능력" in row_str:
            return "KYO"
    
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
    df['번호'] = df['번호'].astype(int) 
    
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
    df['번호'] = df['번호'].astype(int) 
    
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
    df['번호'] = df['번호'].astype(int)
    
    df = df[df['내용'].astype(str) != '희망분야']
    df = df[~df['내용'].astype(str).str.contains('희망분야', na=False)]
    df = df.dropna(subset=['내용'])

    df_grouped = df.groupby(['번호', '과목/영역', '시수'])['내용'].apply(lambda x: ' '.join(x.astype(str))).reset_index()
    
    df_grouped['학년 반'] = grade_class
    df_grouped['학기'] = '' 
    
    return df_grouped

# -----------------------------------------------------------------------------
# 3. 중복 탐지 및 교차 검증 로직
# -----------------------------------------------------------------------------

COLOR_PALETTE = [
    '#ffadad', '#ffd6a5', '#fdffb6', '#caffbf', '#9bf6ff', '#a0c4ff', '#bdb2ff', '#ffc6ff', '#fffffc'
]

@st.cache_data
def detect_duplicates(df):
    if df.empty: return df
    
    df['중복여부'] = False
    df['복붙 의심 문장'] = ''
    df['색상'] = '' 
    df['과목/영역'] = df['과목/영역'].fillna('기타')
    
    color_idx = 0
    duplicate_color_map = {}
    
    for subject, group in df.groupby(['유형', '과목/영역']):
        if len(group) < 2: continue
        
        sentence_counts = {}
        for idx, row in group.iterrows():
            content = str(row['내용'])
            sentences = [s.strip() for s in re.split(r'[.!?\n]+', content) if len(s.strip()) >= 10]
            for s in sentences:
                sentence_counts[s] = sentence_counts.get(s, 0) + 1
        
        duplicate_sentences = {s for s, count in sentence_counts.items() if count > 1}
        
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
                df.at[idx, '색상'] = duplicate_color_map[unique_dupes[0]]

    ordered_cols = ['학년 반', '학기', '과목/영역', '번호', '시수', '내용', '복붙 의심 문장', '중복여부', '색상', '유형']
    final_cols = [c for c in ordered_cols if c in df.columns] 
    return df[final_cols]

def get_sentence_map(df):
    """데이터프레임 내의 문장별 사용 내역을 해시맵으로 추출 (교차 검증용)"""
    sent_map = {}
    for idx, row in df.iterrows():
        subj = (row.get('유형', ''), row.get('과목/영역', ''))
        content = str(row['내용'])
        grade_class = row['학년 반']
        num = row['번호']
        sentences = [s.strip() for s in re.split(r'[.!?\n]+', content) if len(s.strip()) >= 10]
        for s in sentences:
            if subj not in sent_map:
                sent_map[subj] = {}
            if s not in sent_map[subj]:
                sent_map[subj][s] = {}
            if grade_class not in sent_map[subj][s]:
                sent_map[subj][s][grade_class] = []
            if num not in sent_map[subj][s][grade_class]:
                sent_map[subj][s][grade_class].append(num)
    return sent_map

@st.cache_data
def run_cross_validation(df1, df2):
    """그룹1과 그룹2 사이의 동일 유형 데이터 교차 검증"""
    if df1 is None or df2 is None or df1.empty or df2.empty:
        return None
    
    map1 = get_sentence_map(df1)
    map2 = get_sentence_map(df2)
    
    cross_results = []
    
    for subj in set(map1.keys()).intersection(set(map2.keys())):
        type_val, subject = subj
        sentences1 = map1[subj]
        sentences2 = map2[subj]
        
        common_sentences = set(sentences1.keys()).intersection(set(sentences2.keys()))
        
        for s in common_sentences:
            g1_usage = []
            for gc, nums in sentences1[s].items():
                nums_str = ", ".join([f"{n}번" for n in sorted(nums)])
                g1_usage.append(f"[{gc}] {nums_str}")
            g1_str = " \n ".join(g1_usage)
            
            g2_usage = []
            for gc, nums in sentences2[s].items():
                nums_str = ", ".join([f"{n}번" for n in sorted(nums)])
                g2_usage.append(f"[{gc}] {nums_str}")
            g2_str = " \n ".join(g2_usage)
            
            cross_results.append({
                '과목/영역': subject,
                '복붙 의심 문장': s,
                '그룹1 파일의 학년 반': g1_str,
                '그룹 2 파일의 학년 반': g2_str
            })
            
    if cross_results:
        return pd.DataFrame(cross_results)
    return None

def style_dataframe(df_to_style):
    def row_style(row):
        styles = [''] * len(row)
        if row.get('중복여부', False) and row.get('색상', '') != '':
            bg_color = f"background-color: {row['색상']}; color: black;"
            for target_col in ['과목/영역', '내용', '복붙 의심 문장']:
                if target_col in row.index:
                    styles[row.index.get_loc(target_col)] = bg_color
        return styles

    display_cols = [c for c in df_to_style.columns if c not in ['중복여부', '색상', '유형']]
    return df_to_style.style.apply(row_style, axis=1), display_cols

@st.cache_data
def to_excel_with_style(df):
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
# 4. 메인 앱 UI (멀티 파일 업로드 및 탭 구조)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="학생부 점검 도우미", layout="wide")

st.title("🏫 학생부 점검 도우미")
st.markdown("""
**지원내용:** 행특, 세특(교과), 창체(자율/진로)

**기능:**
  1. xlsx_data 파일 업로드 시 **자동 분류 및 정리**
  2. **그룹 내 복붙 의심 문장 색상 표시**
  3. **두 그룹 간의 교차 검증 지원** (다른 파일에 복붙한 사례 색출)
""")

# 두 그룹의 결과 저장을 위한 세션 상태 초기화
if 'final_df_1' not in st.session_state: st.session_state.final_df_1 = None
if 'final_df_2' not in st.session_state: st.session_state.final_df_2 = None

# -----------------------------------------------------------------------------
# 파일 변경 시 호출될 콜백 함수 추가
# -----------------------------------------------------------------------------
def reset_group1():
    """그룹 1 파일 업로더에 변경(추가/삭제)이 발생하면 그룹1 결과 초기화"""
    st.session_state.final_df_1 = None

def reset_group2():
    """그룹 2 파일 업로더에 변경(추가/삭제)이 발생하면 그룹2 결과 초기화"""
    st.session_state.final_df_2 = None

col1, col2 = st.columns(2)
with col1:
    st.subheader("📁 그룹 1 파일")
    uploaded_files_1 = st.file_uploader(
        "그룹 1에 처리할 파일을 올려주세요", 
        accept_multiple_files=True, 
        type=['xlsx', 'xls', 'csv'], 
        key="uploader_1",
        on_change=reset_group1  # 상태 변경 시 초기화 콜백
    )
with col2:
    st.subheader("📁 그룹 2 파일")
    uploaded_files_2 = st.file_uploader(
        "그룹 2에 처리할 파일을 올려주세요", 
        accept_multiple_files=True, 
        type=['xlsx', 'xls', 'csv'], 
        key="uploader_2",
        on_change=reset_group2  # 상태 변경 시 초기화 콜백
    )

def process_uploaded_files(files):
    all_results = []
    for file in files:
        df_raw = load_data(file)
        if df_raw is None:
            continue
            
        grade_class = extract_grade_class(df_raw)
        file_type = detect_file_type(df_raw)
        
        processed_df = None
        if file_type == 'HANG':
            processed_df = process_hang(df_raw, grade_class)
        elif file_type == 'KYO':
            processed_df = process_kyo(df_raw, grade_class)
        elif file_type == 'CHANG':
            processed_df = process_chang(df_raw, grade_class)
            
        if processed_df is not None and not processed_df.empty:
            processed_df['유형'] = file_type 
            all_results.append(processed_df)

    if all_results:
        final_df = pd.concat(all_results, ignore_index=True)
        final_df = final_df.sort_values(by=['과목/영역', '번호'])
        return detect_duplicates(final_df)
    return None

if st.button("🚀 전체 파일 분석 시작", type="primary", use_container_width=True):
    if not uploaded_files_1 and not uploaded_files_2:
        st.warning("분석할 파일을 하나 이상 업로드해주세요.")
    else:
        with st.status("파일 분석 및 처리 중...", expanded=True) as status:
            if uploaded_files_1:
                st.write("진행중: 그룹 1 분석...")
                st.session_state.final_df_1 = process_uploaded_files(uploaded_files_1)
            
            if uploaded_files_2:
                st.write("진행중: 그룹 2 분석...")
                st.session_state.final_df_2 = process_uploaded_files(uploaded_files_2)
                
            status.update(label="모든 파일 처리 완료!", state="complete", expanded=False)

# 하나 이상의 그룹 데이터가 분석 완료되었을 경우에만 결과 표시
if st.session_state.final_df_1 is not None or st.session_state.final_df_2 is not None:
    st.divider()
    
    tab1, tab2, tab3 = st.tabs(["📊 그룹 1 결과보기", "📊 그룹 2 결과보기", "🔄 교차 검증 결과 (그룹1 ↔ 그룹2)"])
    
    def render_result_tab(df, group_name):
        if df is not None:
            styler, display_cols = style_dataframe(df)
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
            
            excel_data = to_excel_with_style(df)
            st.download_button(
                label=f"📥 {group_name} 엑셀 파일 다운로드 (.xlsx)",
                data=excel_data,
                file_name=f"생기부_{group_name}_정리결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_btn_{group_name}" 
            )
        else:
            st.info(f"{group_name}에 처리할 수 있는 정상적인 데이터가 없거나 분석되지 않았습니다.")

    with tab1:
        render_result_tab(st.session_state.final_df_1, "그룹1")
        
    with tab2:
        render_result_tab(st.session_state.final_df_2, "그룹2")
        
    with tab3:
        if st.session_state.final_df_1 is not None and st.session_state.final_df_2 is not None:
            cross_df = run_cross_validation(st.session_state.final_df_1, st.session_state.final_df_2)
            if cross_df is not None and not cross_df.empty:
                st.success(f"⚠️ 두 그룹 사이에서 총 **{len(cross_df)}개**의 동일 문장이 발견되었습니다.")
                st.dataframe(
                    cross_df,
                    column_config={
                        "복붙 의심 문장": st.column_config.TextColumn("복붙 의심 문장", width="large"),
                        "그룹1 파일의 학년 반": st.column_config.TextColumn("그룹1 파일의 학년 반", width="medium"),
                        "그룹 2 파일의 학년 반": st.column_config.TextColumn("그룹 2 파일의 학년 반", width="medium"),
                    },
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.balloons()
                st.success("🎉 두 그룹 간에 교차되는 중복(복붙) 문장이 발견되지 않았습니다!")
        else:
            st.warning("교차 검증을 진행하려면 그룹 1과 그룹 2 모두 업로드 및 분석이 완료되어야 합니다.")
