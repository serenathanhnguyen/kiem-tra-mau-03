import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Kiểm tra Mẫu 03 Medinet",
    page_icon="📋",
    layout="wide"
)

st.title("KIỂM TRA FILE MẪU 03 MEDINET")

st.write(
    "Công cụ kiểm tra dữ liệu Excel M03 trước khi nhập lên hệ thống Medinet."
)

uploaded_file = st.file_uploader(
    "Chọn file Excel M03",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    st.success("Đã nhận file Excel.")

    excel = pd.ExcelFile(uploaded_file)

    st.write("Các sheet có trong file:")
    st.write(excel.sheet_names)
