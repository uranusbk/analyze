# -*- coding: utf-8 -*-
"""Streamlit 웹앱: 불량율 분석"""

from datetime import datetime

import streamlit as st

from analyzer import RAW_HEADERS, analyze_uploaded_files

APP_TITLE = "불량율 분석"
CREATOR_TEXT = "created by 김지연"

st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="wide")

st.title(APP_TITLE)
st.caption(CREATOR_TEXT)

st.markdown(
    """
PDF 제품평가보고서 파일을 여러 개 업로드하면, 내용을 읽어 하나의 엑셀 파일로 통합합니다.

- PDF가 아닌 파일은 자동으로 무시합니다.
- 정해진 양식으로 읽지 못한 PDF는 처리 로그에 실패 이유를 표시합니다.
- 분석 결과가 있으면 엑셀 파일을 다운로드할 수 있습니다.
"""
)

uploaded_files = st.file_uploader(
    "PDF 파일 선택",
    type=["pdf"],
    accept_multiple_files=True,
    help="여러 PDF 파일을 한 번에 선택할 수 있습니다.",
)

col1, col2 = st.columns([1, 4])
with col1:
    run = st.button("분석 실행", type="primary", use_container_width=True, disabled=not uploaded_files)

if "result" not in st.session_state:
    st.session_state.result = None

if run:
    with st.spinner("PDF 파일을 분석하는 중입니다..."):
        st.session_state.result = analyze_uploaded_files(uploaded_files)

result = st.session_state.result

if result is not None:
    records = result["records"]
    failed_files = result["failed_files"]
    logs = result["logs"]
    excel_bytes = result["excel_bytes"]

    total = len([f for f in uploaded_files if f.name.lower().endswith(".pdf")]) if uploaded_files else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("PDF 파일 수", total)
    m2.metric("성공", len(records))
    m3.metric("실패", len(failed_files))

    st.subheader("처리 로그")
    st.text_area(
        "로그",
        value="\n".join(logs),
        height=300,
        label_visibility="collapsed",
    )

    if failed_files:
        st.subheader("실패 파일 및 이유")
        for filename, reason in failed_files:
            with st.expander(filename):
                st.code(reason, language="text")

    if records:
        st.subheader("추출 결과 미리보기")
        preview_rows = []
        for rec in records:
            row = {h: rec.get(h) for h in RAW_HEADERS[:23]}
            preview_rows.append(row)
        st.dataframe(preview_rows, use_container_width=True, hide_index=True)

        output_name = f"불량율_분석_통합_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            label="엑셀 파일 다운로드",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    else:
        st.error("엑셀로 저장할 수 있는 PDF가 없습니다. 처리 로그의 실패 이유를 확인해 주세요.")
