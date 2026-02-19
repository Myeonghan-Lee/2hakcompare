import streamlit as st
import pandas as pd
import io
import re
import xlsxwriter

# 웹앱 기본 설정
st.set_page_config(page_title="세특/행특 데이터 전처리 및 교차 검증 도구", layout="wide")
st.title("📄 나이스 세특/행특 데이터 종합 분석기")
st.write("나이스 파일(세특 또는 행특)을 업로드하면 **파일 종류를 자동 인식**하여 정제 규격을 통일하고, **내부 중복 검사** 및 **파일 간 복붙 의심(교차 검증)**을 수행합니다.")

# 1. 단일 파일 정제 및 내부 중복 검사 함수
def process_single_file(uploaded_file, file_name):
    # (1) 파일 종류 판별 및 '반' 정보 추출을 위해 파일의 첫 5줄만 먼저 읽기
    uploaded_file.seek(0)
    header_df = pd.read_excel(uploaded_file, nrows=5, header=None)
    header_text = "".join(header_df.astype(str).values.flatten())
    
    is_haengteuk = False
    class_num = ""
    
    if "행동특성" in header_text.replace(" ", ""):
        is_haengteuk = True
        # 메타데이터에서 'N학년 N반' 중 '반' 숫자 추출
        for val in header_df.astype(str).values.flatten():
            match = re.search(r'(\d+)\s*반', val)
            if match:
                class_num = int(match.group(1))
                break

    # (2) 실제 데이터 읽기
    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file, skiprows=4)
    
    # [수정된 부분] 열 이름에 결측치나 숫자가 섞여 있을 경우를 대비해 문자로 변환 후 필터링
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False)]
    
    # (3) 세특 vs 행특 맞춤형 전처리 로직
    if not is_haengteuk:
        # --- 세특 처리 ---
        if '과 목' in df.columns:
            df = df[~df['과 목'].astype(str).str.contains('과 목|1학년|2학년|3학년', na=False)]
            
        target_col_raw = [col for col in df.columns if '세부능력' in col.replace(" ", "")][0]
        df = df.dropna(subset=[target_col_raw])
        
        fill_cols = [col for col in ['과 목', '학 년', '학기', '번 호'] if col in df.columns]
        df[fill_cols] = df[fill_cols].ffill()
        
        # 통합 처리를 위해 내용 열 이름 통일
        df.rename(columns={target_col_raw: '세부능력 및 특기사항'}, inplace=True)
        
    else:
        # --- 행특 처리 ---
        # 타겟 열 찾기
        target_col_raw = [col for col in df.columns if '행동특성' in col.replace(" ", "")][0]
        num_col_raw = [col for col in df.columns if '번' in col][0]
        
        # 데이터 중간에 낀 반복 헤더 제거
        df = df[~df[num_col_raw].astype(str).str.contains('번 호|1학년|2학년|3학년|/', na=False)]
        df = df.dropna(subset=[target_col_raw])
        
        fill_cols = [col for col in ['학 년', '번 호'] if col in df.columns]
        df[fill_cols] = df[fill_cols].ffill()
        
        # 행특 전용 열 추가 및 맵핑
        df['과 목'] = '행동특성'
        df['학기'] = class_num if class_num else 1  # '반' 정보를 '학기' 열에 삽입
        
        df.rename(columns={target_col_raw: '세부능력 및 특기사항'}, inplace=True)

    # (4) 공통 전처리: 타입 변환 및 이름(성명) 열 삭제
    for col in ['학 년', '학기', '번 호']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')
            
    name_col = [col for col in df.columns if '성' in col and '명' in col]
    if name_col:
        df = df.drop(columns=[name_col[0]])
        
    subject_col = '과 목'
    num_col = [col for col in df.columns if '번' in col and '호' in col][0]
    target_col = '세부능력 및 특기사항' 
    
    # (5) 끊어진 내용 병합
    groupby_cols = [col for col in [subject_col, '학 년', '학기', num_col] if col in df.columns]
    df = df.groupby(groupby_cols, as_index=False).agg({
        target_col: lambda x: "".join(x.astype(str))
    })
    
    # 정렬
    df = df.sort_values(by=[subject_col, num_col]).reset_index(drop=True)
    
    # (6) 문장 추출 및 내부 중복 검사
    sentences_map = {}
    for _, row in df.iterrows():
        subj = row[subject_col]
        num = str(row[num_col])
        text = str(row[target_col])
        sentences = [s.strip() for s in re.findall(r'[^.!?\n]+[.!?]+', text) if s.strip()]
        
        if subj not in sentences_map:
            sentences_map[subj] = {}
            
        for s in sentences:
            if len(s) > 5:
                if s not in sentences_map[subj]:
                    sentences_map[subj][s] = set()
                sentences_map[subj][s].add(num)
                
    internal_dups = {}
    for subj, sents in sentences_map.items():
        dups = {s: len(nums) for s, nums in sents.items() if len(nums) > 1}
        if dups:
            internal_dups[subj] = dups

    df['중복 문장'] = ""
    for idx, row in df.iterrows():
        subj = row[subject_col]
        text = str(row[target_col])
        found_dups = [dup for dup in internal_dups.get(subj, {}).keys() if dup in text]
        if found_dups:
            df.at[idx, '중복 문장'] = "\n".join(found_dups)

    # (7) 컬럼 순서 재배치
    ordered_cols = ['학 년', '학기', subject_col, num_col, target_col, '중복 문장']
    ordered_cols = [col for col in ordered_cols if col in df.columns] 
    df = df[ordered_cols]

    # (8) 미리보기 스타일링 및 엑셀 파일 생성
    bg_colors = ['#ffe6e6', '#e6ffe6', '#e6e6ff', '#ffffe6', '#ffe6ff', '#e6ffff', '#fff2e6', '#f2e6ff', '#e6f2ff', '#e6fffa']
    subject_dup_bg = {}
    for subj, dups in internal_dups.items():
        subject_dup_bg[subj] = {}
        for i, dup in enumerate(sorted(dups.keys(), key=len, reverse=True)):
            subject_dup_bg[subj][dup] = bg_colors[i % len(bg_colors)]

    def highlight_dup(row):
        styles = [''] * len(row)
        subj = row.get(subject_col, "")
        text = str(row.get(target_col, ""))
        found_dups = [dup for dup in internal_dups.get(subj, {}).keys() if dup in text]
        if found_dups:
            bg_color = subject_dup_bg[subj][found_dups[0]]
            highlight = f'background-color: {bg_color}; color: #333; font-weight: bold;'
            if num_col in df.columns: styles[df.columns.get_loc(num_col)] = highlight
            if target_col in df.columns: styles[df.columns.get_loc(target_col)] = highlight
        return styles
    
    styled_df = df.style.apply(highlight_dup, axis=1)

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('정제_결과')
    wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter'})
    text_colors = ['#FF0000', '#0000FF', '#008000', '#FF8C00', '#800080', '#FF00FF', '#008080', '#A52A2A', '#D2691E']
    
    format_cache = {}
    def get_format(color):
        if color not in format_cache:
            format_cache[color] = workbook.add_format({'color': color, 'text_wrap': True, 'valign': 'vcenter'})
        return format_cache[color]
    
    header_format = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1, 'align': 'center'})
    
    for col_num, header in enumerate(df.columns):
        display_header = header
        if is_haengteuk and header == '학기':
            display_header = '반'
        worksheet.write(0, col_num, display_header, header_format)
        
    row_num = 1
    for _, row in df.iterrows():
        subj = row[subject_col]
        duplicates = internal_dups.get(subj, {})
        dup_colors = {}
        c_idx = 0
        for dup_s in sorted(duplicates.keys(), key=len, reverse=True):
            dup_colors[dup_s] = text_colors[c_idx % len(text_colors)]
            c_idx += 1
            
        for col_num, header in enumerate(df.columns):
            val = row[header]
            if pd.isna(val) or val == "":
                worksheet.write(row_num, col_num, "", wrap_format)
                continue
                
            val_str = str(val)
            if header == target_col and duplicates and row['중복 문장'] != "":
                from re import escape
                import re as regex
                pattern = regex.compile('(' + '|'.join(map(escape, dup_colors.keys())) + ')')
                parts = pattern.split(val_str)
                rich_string_args = []
                for part in parts:
                    if not part: continue
                    if part in dup_colors: rich_string_args.extend([get_format(dup_colors[part]), part])
                    else: rich_string_args.append(part)
                
                if len(rich_string_args) > 1: worksheet.write_rich_string(row_num, col_num, *rich_string_args, wrap_format)
                elif len(rich_string_args) == 1: worksheet.write(row_num, col_num, rich_string_args[0], wrap_format)
                else: worksheet.write(row_num, col_num, "", wrap_format)
            else:
                if isinstance(val, (int, float)): worksheet.write_number(row_num, col_num, val, wrap_format)
                else: worksheet.write_string(row_num, col_num, val_str, wrap_format)
        row_num += 1

    for idx, col_name in enumerate(df.columns):
        if col_name in ['학 년', '학기', num_col]: worksheet.set_column(idx, idx, 6)
        elif col_name == subject_col: worksheet.set_column(idx, idx, 16)
        elif col_name == target_col: worksheet.set_column(idx, idx, 70)
        elif col_name == '중복 문장': worksheet.set_column(idx, idx, 40)
    
    workbook.close()
    excel_data = output.getvalue()
    
    return styled_df, excel_data, sentences_map

# --- 메인 UI 구성 ---
col1, col2 = st.columns(2)
with col1:
    file1 = st.file_uploader("첫 번째 파일 업로드 (세특 또는 행특)", type=['xlsx'])
with col2:
    file2 = st.file_uploader("두 번째 파일 업로드 (세특 또는 행특)", type=['xlsx'])

st.divider()

if file1 is not None and file2 is not None:
    with st.spinner('파일 양식을 판별하여 데이터를 정제 및 비교 분석 중입니다...'):
        style1, excel1, map1 = process_single_file(file1, "첫 번째 파일")
        style2, excel2, map2 = process_single_file(file2, "두 번째 파일")
        
        cross_data = []
        common_subjects = set(map1.keys()).intersection(set(map2.keys()))
        
        for subj in common_subjects:
            common_sentences = set(map1[subj].keys()).intersection(set(map2[subj].keys()))
            for sent in common_sentences:
                nums1 = ", ".join(sorted(list(map1[subj][sent]), key=lambda x: int(x) if x.isdigit() else x))
                nums2 = ", ".join(sorted(list(map2[subj][sent]), key=lambda x: int(x) if x.isdigit() else x))
                cross_data.append({
                    "과목": subj,
                    "동일 문장": sent,
                    "첫번째 파일 번호": nums1,
                    "두번째 파일 번호": nums2
                })
        
        cross_df = pd.DataFrame(cross_data)
        if not cross_df.empty:
            cross_df = cross_df.sort_values(by=["과목", "동일 문장"]).reset_index(drop=True)
            
            cross_output = io.BytesIO()
            with pd.ExcelWriter(cross_output, engine='xlsxwriter') as writer:
                cross_df.to_excel(writer, index=False, sheet_name='교차검증_결과')
                workbook = writer.book
                worksheet = writer.sheets['교차검증_결과']
                wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'vcenter'})
                header_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
                for col_num, value in enumerate(cross_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                for row_num in range(1, len(cross_df) + 1):
                    for col_num in range(len(cross_df.columns)):
                        worksheet.write(row_num, col_num, cross_df.iloc[row_num - 1, col_num], wrap_format)
                worksheet.set_column(0, 0, 15)
                worksheet.set_column(1, 1, 80)
                worksheet.set_column(2, 3, 20)
            cross_excel_data = cross_output.getvalue()
            
    tab1, tab2, tab3 = st.tabs(["📊 첫 번째 파일 정제 결과", "📊 두 번째 파일 정제 결과", "🔍 교차 검증(두 파일 비교) 결과"])
    
    with tab1:
        st.subheader("첫 번째 파일 분석 내역")
        st.download_button(label="📥 첫 번째 파일 다운로드 (XLSX)", data=excel1, file_name="cleaned_file1.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(style1, use_container_width=True)
        
    with tab2:
        st.subheader("두 번째 파일 분석 내역")
        st.download_button(label="📥 두 번째 파일 다운로드 (XLSX)", data=excel2, file_name="cleaned_file2.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(style2, use_container_width=True)
        
    with tab3:
        st.subheader("교차 검증 분석 (두 파일 간 동일 문장 사용 내역)")
        if not cross_df.empty:
            st.download_button(label="📥 교차 검증 결과 다운로드 (XLSX)", data=cross_excel_data, file_name="cross_check_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.dataframe(cross_df, use_container_width=True)
        else:
            st.success("✅ 교차 검증 완료! 두 파일 간에 복사된 동일 문장이 없습니다.")

elif file1 is not None or file2 is not None:
    st.warning("분석을 시작하려면 두 개의 파일을 모두 업로드해 주세요.")
