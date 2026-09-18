from io import BytesIO

import pandas as pd
import streamlit as st

from validators import validate_workbook, annotate_workbook, SHEET_MAIN

st.set_page_config(page_title="Kiểm tra file Mẫu 03 - Medinet", page_icon="✅", layout="wide")
st.title("✅ Kiểm tra file Excel Mẫu 03 (KSKDK) trước khi nhập Medinet")
st.caption(
    "Tải lên file Excel dữ liệu bệnh nhân (đúng cấu trúc file mẫu Medinet — sheet "
    f"**{SHEET_MAIN}**, dòng 2 là nhãn, dòng 4 là mã field, dữ liệu từ dòng 5). "
    "Công cụ sẽ kiểm tra định dạng, ô bắt buộc, khớp danh mục (Tỉnh/Phường xã, Nghề nghiệp, "
    "Nơi công tác, Đối tượng khám...), CCCD trùng, và khoảng trắng ẩn."
)

uploaded = st.file_uploader("Chọn file Excel (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        with st.spinner("Đang kiểm tra..."):
            issues_df, structural_notes, n_rows_checked, col_defs, theluc_by_row = validate_workbook(uploaded.getvalue())
    except Exception as e:
        st.error(f"Không đọc được file: {e}")
        st.stop()

    n_errors = int((issues_df["Mức độ"] == "Lỗi").sum()) if not issues_df.empty else 0
    n_warnings = int((issues_df["Mức độ"] == "Cảnh báo").sum()) if not issues_df.empty else 0
    n_bad_rows = issues_df.loc[issues_df["Mức độ"] == "Lỗi", "Dòng Excel"].nunique() if not issues_df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Số dòng đã kiểm tra", n_rows_checked)
    c2.metric("Số lỗi", n_errors)
    c3.metric("Số cảnh báo", n_warnings)
    c4.metric("Số dòng có lỗi", n_bad_rows)

    if structural_notes:
        with st.expander("⚠ Ghi chú về cấu trúc file (dòng nhãn/mã)", expanded=True):
            for note in structural_notes:
                st.warning(note)

    if issues_df.empty:
        st.success("Không phát hiện lỗi hay cảnh báo nào. 🎉")
    else:
        tab1, tab2 = st.tabs(["🔴 Lỗi (chặn nhập liệu)", "🟡 Cảnh báo (nên xem lại)"])
        with tab1:
            df_err = issues_df[issues_df["Mức độ"] == "Lỗi"].drop(columns=["Mức độ"])
            st.dataframe(df_err, use_container_width=True, hide_index=True)
        with tab2:
            df_warn = issues_df[issues_df["Mức độ"] == "Cảnh báo"].drop(columns=["Mức độ"])
            st.dataframe(df_warn, use_container_width=True, hide_index=True)

    # File Excel đã kiểm tra: đã điền phân loại thể lực + tô màu/ghi chú ô lỗi (luôn có, kể cả khi không lỗi)
    try:
        annotated = annotate_workbook(uploaded.getvalue(), issues_df, theluc_by_row)
        st.download_button(
            "⬇ Tải file Excel đã kiểm tra (đã điền phân loại thể lực + đánh dấu ô lỗi)",
            data=annotated,
            file_name="mau03_da_kiem_tra.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(
            f"Đã tự tính & điền **Phân Loại thể lực** cho {len(theluc_by_row)} dòng. "
            "Trong file tải về: ô **đỏ** là lỗi, ô **vàng** là cảnh báo — di chuột vào ô để xem ghi chú chi tiết."
        )
    except Exception as e:
        st.warning(f"Không tạo được file Excel đã đánh dấu: {e}")

    if not issues_df.empty:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            issues_df.to_excel(writer, index=False, sheet_name="Ket_qua_kiem_tra")
        st.download_button(
            "⬇ Tải báo cáo lỗi dạng bảng (Excel)",
            data=buf.getvalue(),
            file_name="bao_cao_kiem_tra_mau03.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with st.expander("ℹ Công cụ đang kiểm tra những gì?"):
    st.markdown(
        """
- **Ô bắt buộc** (nhãn có dấu `*`) không được để trống.
- **Ngày tháng** (`ngay_kham`, `ngay_sinh`) đúng định dạng `dd/MM/yyyy`.
- **CCCD** đủ 12 chữ số, và **không trùng** với dòng khác trong cùng file. Ngoài ra suy từ CCCD để
  đối chiếu: ký tự thứ 4 (chẵn → Nam, lẻ → Nữ) so với ô **giới tính**; ký tự thứ 4-6 (thế kỷ + 2 số
  cuối năm sinh) so với năm của ô **ngày sinh** — lệch thì báo lỗi.
- **Số điện thoại** đúng dạng số Việt Nam thông thường (cảnh báo, không chặn).
- **Các ô chọn 1 giá trị cố định**: giới tính, nhóm máu, yếu tố Rh, loại khám, phân loại thể lực (1-5),
  các ô "Chưa phát hiện bất thường" (0/1), các câu tiền sử bệnh Có/Không...
- **Mã ICD**: kiểm tra định dạng giống ICD-10 (cảnh báo nếu lạ, không có danh mục đầy đủ để đối chiếu).
- **Các ô số** (chiều cao, cân nặng, mạch, huyết áp, xét nghiệm...) phải là số hợp lệ, và nếu nhập
  dạng chữ thì **phần thập phân phải dùng dấu phẩy (,)** — dấu chấm (.) chỉ chấp nhận khi rõ ràng là
  phân cách hàng nghìn của số nguyên, còn lại bị coi là sai định dạng theo đúng chuẩn Medinet.
- **Khớp danh mục chặt**: Đối tượng khám, Tỉnh, Phường/Xã (đối chiếu Phường/Xã có thuộc đúng Tỉnh),
  Nghề nghiệp, Nơi công tác, Bệnh tiền sử gia đình — đối chiếu với các sheet danh mục có trong
  chính file Excel (`DoiTuongKham`, `Tinh`, `PhuongXa`, `NgheNghiep`, `NoiLamViec`, `TienSuGiaDinh`).
- **Ô Kết luận (`danh_muc_de_nghi`)** phải là 1 trong 5 giá trị cố định (Bình thường hẹn khám định kỳ /
  Có yếu tố nguy cơ / Đã có bệnh mạn tính / Chuyển tuyến / Khác).
- **Đối tượng khám = 1 (Sinh viên, học viên) hoặc = 2 (Người lao động chính thức)**: bắt buộc nhập
  `nghenghiep_code`, `noi_cong_tac`, `noi_cong_tac_xa_phuong`; riêng `noi_cong_tac` lúc này chỉ cần
  có dữ liệu, không đối chiếu danh mục `NoiLamViec` (với đối tượng khác thì `noi_cong_tac` không bắt
  buộc, có thể để trống).
- **13 khối chuyên khoa ở tab Khám lâm sàng** (4 ô: chưa phát hiện bất thường / chẩn đoán sơ bộ /
  chẩn đoán xác định / phân loại): chọn "chưa phát hiện bất thường" thì không được có ICD và phân
  loại phải là Loại 1; có ICD (sơ bộ hoặc xác định) thì phân loại phải từ Loại 2 trở lên; ô phân
  loại không được để trống.
- **Sản khoa / Phụ khoa** (có thêm ô "từ chối khám" ở đầu): nếu chọn "từ chối khám" thì không bắt
  buộc chọn phân loại; nếu không từ chối khám thì áp dụng đúng quy tắc 4 ô như các khối chuyên khoa
  khác ở trên.
- **Giới tính Nam**: cảnh báo nếu vẫn có dữ liệu ở các ô chỉ dành cho nữ (tiền sử thai sản, toàn bộ
  khối Sản khoa/Phụ khoa).
- **Phân loại thể lực (cột "Phân Loại thể lực")**: công cụ tự tính theo QĐ 1613/BYT từ chiều cao +
  cân nặng (theo giới tính; bảng học sinh/SV nếu Đối tượng khám = mã 1, còn lại dùng bảng người lao
  động) rồi **điền sẵn vào cột này** trong file tải về. Vì file không có cột vòng ngực nên chỉ dùng 2
  chỉ số, lấy loại kém hơn.
- **3 cặp đo thị lực Mắt** (không kính / kính lỗ / có kính, mỗi cặp gồm mắt phải + mắt trái): phải
  điền theo từng cặp (cùng có hoặc cùng trống); cặp "không kính" loại trừ với 2 cặp "kính lỗ" và
  "có kính" — điền cặp này thì không điền cặp kia.
- **`giadinh_macbenh` và `giadinh_danhsachbenh_icd`** chỉ ở mức **cảnh báo**, không chặn nhập liệu.
- **Khoảng trắng ẩn** (dấu cách không ngắt `\\xa0`, ký tự rộng-0...) trong bất kỳ ô chữ nào — dấu vết
  hay gặp khi copy dữ liệu từ web/PDF, từng gây lỗi "điền thành công giả" ở Nơi công tác.
- **Cấu trúc file**: phát hiện nếu 2 cột vô tình dùng trùng 1 mã field (lỗi hiếm gặp trong file gốc).

Công cụ **không** kiểm tra: nội dung mô tả tự do (ghi chú, mô tả lâm sàng), và không có danh mục
đầy đủ cho Dân tộc / Hình thức chi trả nên các ô đó chỉ được kiểm tra "có dữ liệu hay không", không
đối chiếu danh mục.
        """
    )
