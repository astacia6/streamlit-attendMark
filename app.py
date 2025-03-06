import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import zipfile
import os
import time

# 페이지 설정
st.set_page_config(page_title="출결 서류 처리", layout="wide")

# 제목 및 설명 추가
st.title("출결 서류 항목 표시")
st.markdown("""
    <style>
    .main {background-color: #f0f2f6;}
    </style>
    """, unsafe_allow_html=True)

st.markdown("### 반별 출결 엑셀 파일을 업로드하세요.")
st.markdown("업로드된 파일에서 '질병결석' 또는 '인정결석'이 포함된 행을 음영 처리합니다.")

# 레이아웃 조정
col1, col2 = st.columns([1, 3])

with col1:
    # 파일 업로드
    uploaded_files = st.file_uploader("엑셀 파일 선택", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    # 결과 파일을 저장할 임시 디렉토리 생성
    os.makedirs("temp", exist_ok=True)
    
    for uploaded_file in uploaded_files:
        # 엑셀 파일 읽기
        df = pd.read_excel(uploaded_file)

        # 특정 열의 값을 검사하여 조건에 맞는 행에 음영 처리
        def highlight_rows(workbook):
            sheet = workbook.active  # 첫 번째 시트를 선택
            fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
            
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                cell = row[3]  # 4번째 열 (0부터 시작하므로 인덱스 3)
                if "질병결석" in str(cell.value) or "인정결석" in str(cell.value):
                    for c in row:
                        c.fill = fill

        # 엑셀 파일을 다시 저장
        timestamp = int(time.time())
        output_path = f"temp/{timestamp}_{uploaded_file.name}"
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            workbook = writer.book
            highlight_rows(workbook)

    # ZIP 파일로 압축
    zip_filename = "highlighted_files.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for file in os.listdir("temp"):
            zipf.write(os.path.join("temp", file), file)

    # 안내 문구 및 이모지 추가
    st.markdown("<h3>출결 필수 서류 항목 표시가 완료되었습니다. :point_down:</h3>", unsafe_allow_html=True)

    # 다운로드 링크 제공
    with open(zip_filename, "rb") as file:
        btn = st.download_button(
            label="다운로드",
            data=file,
            file_name=zip_filename,
            mime="application/zip"
        )

    # 임시 디렉토리 정리
    for file in os.listdir("temp"):
        os.remove(os.path.join("temp", file))
    os.rmdir("temp") 