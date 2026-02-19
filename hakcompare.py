import streamlit as st
import pandas as pd
import re
import io
import itertools

# -----------------------------------------------------------------------------
# 0. 페이지 기본 설정
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
    
    df = df[df['내용'].astype(str) != '희망분야']
    df = df[~df['내용'].astype(str).str.contains('희망분야', na=False)]
    df = df.dropna(subset=['내용'])

    df_grouped = df.groupby(['번호', '과목/영역', '시수'])['내용'].apply(lambda x: ' '.join(x.astype(str))).reset_index()
    
    df_grouped['학년 반'] = grade_class
    df_grouped['학기'] = '' 
    
    return df_grouped[['학년 반', '번호', '학기', '과목/영역', '시수', '내용']]

def detect_duplicates(df):
    """단일 파일 내 복붙(중복) 문장 탐지"""
    sentence_pattern = re.compile(r'[^.!?]+[.!?]')
    df['중복여부'] = False
    df['비고(중복문장)'] = ''
    df['중복배경색'] = '' 
    df['중복글자색'] = ''
    
    color_pairs = [
        ('#ffe6e6', '#cc0000'), ('#e6f2ff', '#004080'), ('#e6ffe6', '#006600'),
        ('#fff2e6', '#cc6600'), ('#f2e6ff', '#4d0099'), ('#ffffe6', '#808000'),
        ('#e6ffff', '#006666'), ('#ffe6f2', '#99004d'), ('#f2ffe6', '#4d9900'),
        ('#ebebe0', '#333333')
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
        
        color_map = {}
        for i, dup_sent in enumerate(duplicate_sentences):
            color_map[dup_sent] = color_pairs[i % len(color_pairs)]
            
        for idx, row in group.iterrows():
            content = str(row['내용'])
            sentences = [s.strip() for s in sentence_pattern.findall(content)]
            found_duplicates = [s for s in sentences if s in duplicate_sentences]
            
            if found_duplicates:
                df.at[idx, '중복여부'] = True
                unique_dupes = list(set(found_duplicates))
                df.at[idx, '비고(중복문장)'] = " / ".join(unique_dupes)
                
                bg_color, text_color = color_map[unique_dupes[0]]
                df.at[idx, '중복배경색'] = bg_color
                df.at[idx, '중복글자색'] = text_color

    return df

def cross_validate_files(df1, df2, name1, name2):
    """두 파일 간의 교차 점검 (동일 과목 내 중복 문장 탐색)"""
    sentence_pattern = re.compile(r'[^.!?]+[.!?]')
    cross_results = []
    
    # 두 파일에 공통으로 존재하는 과목/영역 찾기
    subjects1 = set(df1['과목/영역'].dropna().unique())
    subjects2 = set(df2['과목/영역'].dropna().unique())
    common_subjects = subjects1.intersection(subjects2)
    
    for subj in common_subjects:
        group1 = df1[df1['과목/영역'] == subj]
        group2 = df2[df2['과목/영역'] == subj]
        
        # 파일 1의 문장들 수집 (문장 -> 학생정보 리스트)
        sent_map1 = {}
        for _, row in group1.iterrows():
            content = str(row['내용'])
            student_info = f"{row['학년 반']} {row['번호']}번"
            for s in [s.strip() for s in sentence_pattern.findall(content)]:
                if len(s) < 10: continue # 10자 미만 무시
                if s not in sent_map1: sent_map1[s] = []
                sent_map1[s].append(student_info)
                
        # 파일 2의 문장들 수집
        sent_map2 = {}
        for _, row in group2.iterrows():
            content = str(row['내용'])
            student_info = f"{row['학년 반']} {row['번호']}번"
            for s in [s.strip() for s in sentence_pattern.findall(content)]:
                if len(s) < 10: continue
                if s not in sent_map2: sent_map2[s] = []
                sent_map2[s].append(student_info)
                
        # 교차 중복된 문장 찾기 (교집합)
        common_sentences = set(sent_map1.keys()).intersection(set(sent_map2.keys()))
        
        # 결과 리스트에 추가 (과목 내에서 동일 문장이 여러 개면 행 추가)
        for s in common_sentences:
            students1 = ", ".join(list(set(sent_map1[s])))
            students2 = ", ".join(list(set(sent_map2[s])))
            cross_results.append({
                '과목/영역': subj,
                '동일 문장': s,
                f'첫번째 파일({name1}) 학생반 번호': students1,
                f'두번째 파일({name2}) 학생반 번호': students2
            })
            
    if cross_results:
        return pd.DataFrame(cross_results).sort_values(by=['과목/영역'])
    else:
        # 중복이 없을 경우 빈 데이터프레임 반환
        return pd.DataFrame(columns=['과목/영역', '동일 문장', f'첫번째 파일({name1}) 학생반 번호', f'두번째 파일({name2}) 학생반 번호'])

def to_excel_multiple_sheets(df_dict, cross_df=None):
    """여러 데이터프레임과 교차점검 결과를 엑셀에 저장"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for file_name, df in df_dict.items():
            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', file_name)[:31]
            save_cols = [c for c in df.columns if c not in ['중복여부', '중복배경색', '중복글자색']]
            
            def style_duplicate_excel(row):
                styles = [''] * len(row)
                if row.get('중복여부', False) and row.get('중복배경색', ''):
                    bg_color = row['중복배경색']
                    txt_color = row['중복글자색']
                    
                    for col in ['과목/영역', '번호', '내용']:
                        if col in row.index:
                            try:
                                idx = row.index.get_loc(col)
                                styles[idx] = f'background-color: {bg_color}; color: {txt_color}; font-weight: bold;'
                            except KeyError: pass
                    
                    if '비고(중복문장)' in row.index:
                        try:
                            note_idx = row.index.get_loc('비고(중복문장)')
                            styles[note_idx] = 'color: red;'
                        except KeyError: pass
                return styles

            styler = df.style.apply(style_duplicate_excel, axis=1)
            styler.to_excel(writer, index=False, columns=save_cols, sheet_name=safe_sheet_name)
            
            worksheet = writer.sheets[safe_sheet_name]
            for idx, col in enumerate(save_cols):
                width = 50 if '내용' in col or '비고' in col else 12
                worksheet.column_dimensions[chr(65 + idx)].width = width
                
        # 교차 점검 결과 시트 추가
        if cross_df is not None and not cross_df.empty:
            cross_sheet_name = "교차점검결과"
            cross_df.to_excel(writer, index=False, sheet_name=cross_sheet_name)
            worksheet = writer.sheets[cross_sheet_name]
            # 열 너비 조정
            worksheet.column_dimensions['A'].width = 15 # 과목/영역
            worksheet.column_dimensions['B'].width = 60 # 동일 문장
            worksheet.column_dimensions['C'].width = 20 # 첫번째 파일 학생
            worksheet.column_dimensions['D'].width = 20 # 두번째 파일 학생

    return output.getvalue()

# -----------------------------------------------------------------------------
# 3. 메인 앱 UI
# -----------------------------------------------------------------------------

st.title("🏫 학생부 점검 도우미")
st.markdown("""
**지원내용:** 행특, 세특(교과), 창체(자율/진로)

**기능:**
  1. xlsx_data 파일 다운로드 및 업로드 시 **자동 분류 및 정리**
  2. **복붙 의심 문장 그룹별 배경색/글자색 다르게 표시 (과목, 번호, 내용 강조)**
  3. **여러 파일 업로드 시 탭(Tab)으로 구분하여 표시**
  4. **두 개 이상의 파일 업로드 시 파일 간 복붙(교차 점검) 자동 탐지** 🚀
""")

uploaded_files = st.file_uploader(
    "처리할 파일들을 모두 올려주세요", 
    accept_multiple_files=True,
    type=['xlsx', 'xls', 'csv']
)

if uploaded_files:
    processed_data_dict = {}
    
    with st.status("파일 분석 및 처리 중...", expanded=True) as status:
        for file in uploaded_files:
            df_raw = load_data(file)
            if df_raw is None:
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
                processed_df = processed_df.sort_values(by=['과목/영역', '번호'])
                processed_df = detect_duplicates(processed_df)
                processed_df['번호'] = pd.to_numeric(processed_df['번호']).astype(int)
                
                ordered_cols = ['학년 반', '학기', '과목/영역', '번호', '시수', '내용', '비고(중복문장)', '중복여부', '중복배경색', '중복글자색']
                processed_df = processed_df[ordered_cols]
                
                processed_data_dict[file.name] = processed_df
                st.write(f"✅ {file.name} ({type_label} / {grade_class}) - {len(processed_df)}명 처리")

        status.update(label="모든 파일 처리 완료!", state="complete", expanded=False)

    if processed_data_dict:
        st.divider()
        st.subheader("📊 결과 미리보기")
        
        # 교차 점검 로직 실행 (파일이 2개 이상일 때, 처음 두 파일 기준)
        cross_df = None
        file_names = list(processed_data_dict.keys())
        
        if len(file_names) >= 2:
            name1, name2 = file_names[0], file_names[1]
            df1, df2 = processed_data_dict[name1], processed_data_dict[name2]
            cross_df = cross_validate_files(df1, df2, name1, name2)
            
        # 탭 구성 (파일별 탭 + 교차점검 탭)
        tab_names = file_names.copy()
        if cross_df is not None:
            tab_names.append("🚨 교차 점검 결과")
            
        tabs = st.tabs(tab_names)
        
        def highlight_row_web(row):
            styles = [''] * len(row)
            if row.get('중복여부', False) and row.get('중복배경색', ''):
                bg_color = row['중복배경색']
                txt_color = row['중복글자색']
                for col in ['과목/영역', '번호', '내용']:
                    if col in row.index:
                        try:
                            idx = row.index.get_loc(col)
                            styles[idx] = f'background-color: {bg_color}; color: {txt_color}; font-weight: bold;'
                        except KeyError: pass
            return styles
        
        # 탭 콘텐츠 채우기
        for i, tab in enumerate(tabs):
            with tab:
                if i < len(file_names):
                    # 개별 파일 탭
                    file_name = file_names[i]
                    df_to_show = processed_data_dict[file_name]
                    st.dataframe(
                        df_to_show.style.apply(highlight_row_web, axis=1),
                        column_config={
                            "시수": st.column_config.TextColumn("시수", width="small"),
                            "비고(중복문장)": st.column_config.TextColumn("⚠️ 복붙 의심 문장", width="medium"),
                            "중복여부": None, "중복배경색": None, "중복글자색": None
                        },
                        use_container_width=True
                    )
                else:
                    # 교차 점검 결과 탭
                    if cross_df is not None and not cross_df.empty:
                        st.warning(f"⚠️ {name1} 과(와) {name2} 사이에 내용이 중복된 문장들입니다.")
                        st.dataframe(
                            cross_df,
                            column_config={
                                "동일 문장": st.column_config.TextColumn("동일 문장", width="large"),
                                f"첫번째 파일({name1}) 학생반 번호": st.column_config.TextColumn(f"{name1} 학생", width="medium"),
                                f"두번째 파일({name2}) 학생반 번호": st.column_config.TextColumn(f"{name2} 학생", width="medium")
                            },
                            use_container_width=True
                        )
                    else:
                        st.success("🎉 두 파일 사이에 교차 중복된 문장이 없습니다!")
        
        st.divider()
        excel_data = to_excel_multiple_sheets(processed_data_dict, cross_df=cross_df)
        
        st.download_button(
            label="📥 통합 엑셀 다운로드 (개별시트 + 교차점검시트 포함)",
            data=excel_data,
            file_name="생기부_정리결과_전체.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("처리할 데이터가 없습니다.")
