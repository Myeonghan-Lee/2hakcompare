import streamlit as st
import pandas as pd
import re
import io
import matplotlib.colors as mcolors
import matplotlib.pyplot as plt
import numpy as np

# ... (상단 load_data, extract_grade_class, detect_file_type 등은 기존과 동일) ...

# -----------------------------------------------------------------------------
# [수정] 중복 감지 및 색상 할당 로직
# -----------------------------------------------------------------------------

def detect_duplicates_with_colors(df):
    """중복 문장별로 고유 색상을 할당"""
    sentence_pattern = re.compile(r'[^.!?]+[.!?]')
    df['중복여부'] = False
    df['색상정보'] = None  # {문장: 색상} 형태의 딕셔너리를 저장할 열
    
    df['과목/영역'] = df['과목/영역'].fillna('기타')
    
    # 중복 문장 추출용
    all_duplicate_info = {} # 과목별 중복 문장 색상 관리

    for subject, group in df.groupby('과목/영역'):
        sentence_counts = {}
        for _, row in group.iterrows():
            sentences = [s.strip() for s in sentence_pattern.findall(str(row['내용']))]
            for s in sentences:
                if len(s) < 10: continue
                sentence_counts[s] = sentence_counts.get(s, 0) + 1
        
        # 2회 이상 등장한 문장들
        dupes = [s for s, count in sentence_counts.items() if count > 1]
        
        if dupes:
            # 중복 문장 개수만큼 컬러맵 생성 (너무 밝지 않은 색상 위주)
            cmap = plt.get_cmap('Pastel1', len(dupes))
            color_map = {s: mcolors.to_hex(cmap(i)) for i, s in enumerate(dupes)}
            all_duplicate_info[subject] = color_map

    # 각 행에 색상 정보 매핑
    for idx, row in df.iterrows():
        subj = row['과목/영역']
        if subj in all_duplicate_info:
            content = str(row['내용'])
            subj_dupes = all_duplicate_info[subj]
            found = {s: color for s, color in subj_dupes.items() if s in content}
            if found:
                df.at[idx, '중복여부'] = True
                df.at[idx, '색상정보'] = found # 해당 행에 포함된 중복문장과 색상 저장

    return df, all_duplicate_info

# -----------------------------------------------------------------------------
# [수정] 화면 표시 및 엑셀 스타일링
# -----------------------------------------------------------------------------

def style_df(df):
    """화면 출력용 스타일링"""
    def apply_color(row):
        styles = [''] * len(row)
        if row['색상정보']:
            # 가장 먼저 발견된 중복 문장의 색상을 배경색으로 지정
            first_color = list(row['색상정보'].values())[0]
            content_idx = row.index.get_loc('내용')
            styles[content_idx] = f'background-color: {first_color}; color: black;'
        return styles
    return df.style.apply(apply_color, axis=1)

def to_excel_with_multi_color(df):
    """엑셀 파일에 중복별 배경색 적용"""
    output = io.BytesIO()
    save_cols = [c for c in df.columns if c not in ['중복여부', '색상정보']]
    
    # 스타일 적용
    styler = df.style.apply(lambda row: [
        f'background-color: {list(row["색상정보"].values())[0]}' if row['색상정보'] and col == '내용' else ''
        for col in df.columns
    ], axis=1)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        styler.to_excel(writer, index=False, columns=save_cols, sheet_name='정리결과')
    return output.getvalue()

# -----------------------------------------------------------------------------
# 메인 앱 UI (수정 부분 위주)
# -----------------------------------------------------------------------------

# ... (파일 업로드 및 process_xxx 호출 부분은 동일) ...

    if all_results:
        final_df = pd.concat(all_results, ignore_index=True)
        final_df = final_df.sort_values(by=['과목/영역', '번호'])
        
        # [변경] 중복 분석 실행
        final_df, color_info_master = detect_duplicates_with_colors(final_df)
        
        st.divider()
        st.subheader("📊 결과 미리보기")
        st.caption("💡 같은 색상으로 표시된 셀은 서로 동일한 문장을 포함하고 있습니다.")
        
        # [변경] 스타일이 적용된 데이터프레임 표시
        st.dataframe(
            style_df(final_df),
            column_config={
                "시수": st.column_config.TextColumn("시수", width="small"),
                "중복여부": None,
                "색상정보": None
            },
            use_container_width=True
        )
        
        excel_data = to_excel_with_multi_color(final_df)
        st.download_button(
            label="📥 컬러 중복 체크 엑셀 다운로드",
            data=excel_data,
            file_name="생기부_중복점검_결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
