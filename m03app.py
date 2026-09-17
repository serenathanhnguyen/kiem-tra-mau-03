from io import BytesIO

import pandas as pd
import streamlit as st

from validators import validate_workbook


st.set_page_config(
    page_title="Kiểm tra Mẫu 03 Medinet",
    page_icon="📋",
    layout="wide",
)


def make_report(issues_df, structural_notes, rows_checked):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if issues_df.empty:
            pd.DataFrame([
                {
                    "Kết quả": "Không phát hiện lỗi hoặc cảnh báo",
                    "Số dòng đã kiểm tra": rows_checked,
                }
            ]).to_excel(writer, sheet_name="KET_QUA", index=False)
        else:
            issues_df.to_excel(writer, sheet_name="LOI_CHI_TIET", index=False)

        pd.DataFrame([
            {
                "Số dòng đã kiểm tra": rows_checked,
                "Số lỗi": int((issues_df["Mức độ"] == "Lỗi").sum()) if not issues_df.empty else 0,
                "Số cảnh báo": int((issues_df["Mức độ"] == "Cảnh báo").sum()) if not issues_df.empty else 0,
            }
        ]).to_excel(writer, sheet_name="TONG_HOP", index=False)

        if structural_notes:
            pd.DataFrame({"Lưu ý cấu trúc": structural_notes}).to_excel(
                writer, sheet_name="LUU_Y_CAU_TRUC", index=False
            )

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column in ws.columns:
                width = min(max(len(str(cell.value or "")) for cell in column) + 2, 70)
                ws.column_dimensions[column[0].column_letter].width = max(width, 12)

    return output.getvalue()


st.title("KIỂM TRA FILE EXCEL MẪU 03 MEDINET")
st.write(
    "Tải file Excel Mẫu 03 lên để kiểm tra dữ liệu trước khi nhập vào Medinet. "
    "Công cụ không sửa dữ liệu trong file gốc."
)

with st.expander("Công cụ đang kiểm tra những gì?"):
    st.markdown(
        """
        - Ô bắt buộc, ngày khám, ngày sinh, CCCD và số điện thoại.
        - CCCD trùng trong cùng file.
        - Các giá trị Có/Không, giới tính, nhóm máu, phân loại sức khỏe.
        - Danh mục tỉnh, phường/xã, nghề nghiệp, nơi công tác và đối tượng khám.
        - Khoảng trắng ẩn, trường số, dấu thập phân và mã ICD-10.
        - Logic giữa “chưa phát hiện bất thường”, chẩn đoán ICD và phân loại.
        - Quy tắc Sản khoa, Phụ khoa theo giới tính và lựa chọn từ chối khám.
        """
    )

uploaded_file = st.file_uploader(
    "Chọn file Excel Mẫu 03",
    type=["xlsx"],
    help="File phải có sheet ThongTinHanhChinh và các sheet danh mục đi kèm.",
)

if uploaded_file is None:
    st.info("Chưa có file được tải lên.")
else:
    try:
        with st.spinner("Đang kiểm tra file..."):
            issues, structural_notes, rows_checked, _ = validate_workbook(
                uploaded_file.getvalue()
            )

        if issues.empty:
            error_count = 0
            warning_count = 0
        else:
            error_count = int((issues["Mức độ"] == "Lỗi").sum())
            warning_count = int((issues["Mức độ"] == "Cảnh báo").sum())

        c1, c2, c3 = st.columns(3)
        c1.metric("Dòng đã kiểm tra", rows_checked)
        c2.metric("Lỗi", error_count)
        c3.metric("Cảnh báo", warning_count)

        if structural_notes:
            st.warning("Phát hiện lưu ý về cấu trúc file:")
            for note in structural_notes:
                st.write(f"- {note}")

        if rows_checked == 0:
            st.warning("Không tìm thấy hồ sơ dữ liệu từ dòng 5 trở đi.")
        elif error_count == 0:
            st.success("Không phát hiện lỗi bắt buộc trong file.")
        else:
            st.error(f"Có {error_count} lỗi cần sửa trước khi nhập Medinet.")

        if not issues.empty:
            tab_errors, tab_warnings, tab_all = st.tabs(
                ["Lỗi cần sửa", "Cảnh báo", "Tất cả kết quả"]
            )
            with tab_errors:
                error_df = issues[issues["Mức độ"] == "Lỗi"]
                if error_df.empty:
                    st.info("Không có lỗi.")
                else:
                    st.dataframe(error_df, use_container_width=True, hide_index=True)
            with tab_warnings:
                warning_df = issues[issues["Mức độ"] == "Cảnh báo"]
                if warning_df.empty:
                    st.info("Không có cảnh báo.")
                else:
                    st.dataframe(warning_df, use_container_width=True, hide_index=True)
            with tab_all:
                st.dataframe(issues, use_container_width=True, hide_index=True)

        report = make_report(issues, structural_notes, rows_checked)
        output_name = uploaded_file.name.rsplit(".", 1)[0] + "_BAO_CAO_KIEM_TRA.xlsx"
        st.download_button(
            "Tải báo cáo kiểm tra Excel",
            data=report,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    except Exception as exc:
        st.error("Không thể kiểm tra file Excel này.")
        st.exception(exc)
        st.info(
            "Hãy kiểm tra lại tên sheet ThongTinHanhChinh và bảo đảm file đúng cấu trúc Mẫu 03."
        )
