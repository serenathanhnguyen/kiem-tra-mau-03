from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

from validators import validate_workbook, annotate_workbook, SHEET_MAIN, DE_NGHI_DEFAULT_VALUE
from merger import read_source_file, merge_mau03_files, default_unit_label

st.set_page_config(page_title="Kiểm tra & ghép file Mẫu 03 - Medinet", page_icon="✅", layout="wide")
st.title("✅ Công cụ Mẫu 03 (KSKDK) - Medinet")

tab_check, tab_merge = st.tabs(["🔍 Kiểm tra file", "🔗 Ghép nhiều file"])

# ============================================================
# TAB 1 — KIỂM TRA FILE (giữ nguyên toàn bộ logic cũ)
# ============================================================
with tab_check:
    st.caption(
        "Tải lên file Excel dữ liệu bệnh nhân (đúng cấu trúc file mẫu Medinet — sheet "
        f"**{SHEET_MAIN}**, có dòng nhãn và dòng mã field/keyword như file mẫu — công cụ tự nhận diện "
        "dòng mã field dù nằm ở dòng số mấy, không bắt buộc phải là dòng 4). "
        "Công cụ sẽ kiểm tra định dạng, ô bắt buộc, khớp danh mục (Tỉnh/Phường xã, Nghề nghiệp, "
        "Nơi công tác, Đối tượng khám...), CCCD trùng, và khoảng trắng ẩn."
    )

    uploaded = st.file_uploader("Chọn file Excel (.xlsx)", type=["xlsx"], key="check_uploader")

    if uploaded is not None:
        try:
            with st.spinner("Đang kiểm tra..."):
                (issues_df, structural_notes, n_rows_checked, col_defs, theluc_by_row,
                 danhmucdenghi_by_row, denghi_by_row, data_fixes_by_row) = validate_workbook(uploaded.getvalue())
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
            itab1, itab2 = st.tabs(["🔴 Lỗi (chặn nhập liệu)", "🟡 Cảnh báo (nên xem lại)"])
            with itab1:
                df_err = issues_df[issues_df["Mức độ"] == "Lỗi"].drop(columns=["Mức độ"])
                st.dataframe(df_err, use_container_width=True, hide_index=True)
            with itab2:
                df_warn = issues_df[issues_df["Mức độ"] == "Cảnh báo"].drop(columns=["Mức độ"])
                st.dataframe(df_warn, use_container_width=True, hide_index=True)

        # File Excel đã kiểm tra: đã điền phân loại thể lực + tô màu/ghi chú ô lỗi (luôn có, kể cả khi không lỗi)
        try:
            annotated = annotate_workbook(
                uploaded.getvalue(), issues_df, theluc_by_row, danhmucdenghi_by_row, denghi_by_row,
                data_fixes_by_row
            )
            st.download_button(
                "⬇ Tải file Excel đã kiểm tra (đã tự sửa dữ liệu + canh giữa + đánh dấu ô lỗi)",
                data=annotated,
                file_name="mau03_da_kiem_tra.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_annotated",
            )
            n_fixed_rows = len(data_fixes_by_row)
            st.caption(
                f"Đã tự tính & điền **Phân Loại thể lực** cho {len(theluc_by_row)} dòng, "
                f"đề xuất **Kết luận (danh_muc_de_nghi)** cho {len(danhmucdenghi_by_row)} dòng đang trống, "
                f"điền mặc định **'{DE_NGHI_DEFAULT_VALUE}'** cho ô **de_nghi** ở {len(denghi_by_row)} dòng đang trống, "
                f"và áp các quy tắc tự sửa dữ liệu khác (giới tính, tiền sử bệnh 0/1, ICD=0, loại khám, hồng cầu...) "
                f"cho {n_fixed_rows} dòng. Toàn bộ vùng dữ liệu đã được **canh giữa**. "
                "Trong file tải về: ô **đỏ** là lỗi, ô **vàng** là cảnh báo — di chuột vào ô để xem ghi chú chi tiết. "
                "File không bị khoá — vẫn filter, xoá, copy, paste bình thường."
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
                key="dl_report",
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
- **Các ô số không thuộc cận lâm sàng** (chiều cao, cân nặng, mạch, huyết áp, nhịp thở, thị lực mắt)
  phải là số hợp lệ, và nếu nhập dạng chữ thì **phần thập phân phải dùng dấu phẩy (,)** — dấu chấm (.)
  chỉ chấp nhận khi rõ ràng là phân cách hàng nghìn của số nguyên, còn lại bị coi là sai định dạng
  theo đúng chuẩn Medinet.
- **Các ô số thuộc cận lâm sàng** (xét nghiệm máu, sinh hóa máu, xét nghiệm nước tiểu): nếu có nhập
  thì **chỉ kiểm tra đúng lỗi dùng dấu chấm (.) thay cho dấu phẩy (,)** ở phần thập phân — không kiểm
  tra hay cảnh báo gì khác cho các ô này.
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
- **Tự đề xuất điền khi đang để trống** (chỉ áp dụng trong file Excel tải về, không tính là lỗi/cảnh báo):
  - Ô **Kết luận** (`danh_muc_de_nghi`): xét các ô `*_phanloai`, `*_chandoansobo_icd`,
    `*_chandoanxacdinh_icd` của 13 khối chuyên khoa (và Sản khoa/Phụ khoa, trừ khối đã chọn "từ chối
    khám"). Theo thứ tự — luôn điền ra 1 trong 3 giá trị, không để trống:
    1) có ít nhất 1 khối `*_phanloai` > 1 **và** có ít nhất 1 `*_chandoanxacdinh_icd` được điền →
       "Đã có bệnh mạn tính, tiếp tục điều trị theo phác đồ/toa cũ";
    2) ngược lại, có ít nhất 1 `*_chandoansobo_icd` được điền → "Có yếu tố nguy cơ, cần theo dõi
       thêm"; 3) còn lại → "Bình thường, hẹn khám định kỳ lần sau".
  - Ô **`de_nghi`** (Đề nghị, ghi rõ): nếu đang trống → điền mặc định "Tái khám định kỳ".
- **Tự động sửa dữ liệu khác trong file Excel tải về** (không tính là lỗi/cảnh báo, chỉ áp dụng khi xuất file):
  - **Canh giữa**: toàn bộ vùng dữ liệu (từ dòng đầu tiên có dữ liệu đến dòng cuối, các cột đã map
    được mã field) được canh giữa theo cả chiều ngang và dọc.
  - **Giới tính Nam** (cột `gioi_tinh` = 1): các ô chỉ dành cho nữ (thai sản, toàn bộ khối Sản khoa/
    Phụ khoa — cột AY, AZ, CV đến DE) được **xoá trắng** nếu lỡ có dữ liệu.
  - **Giới tính Nữ** (cột `gioi_tinh` = 2): ô "Có thai sản không" (cột AY) nếu đang là 0 hoặc 1 thì
    giữ nguyên; nếu là chữ "Không" thì chỉnh về **0**.
  - **Cột AA đến AU** (21 câu tiền sử bệnh dạng Có/Không): chỉ nhận 0 hoặc 1 — ô đã điền giá trị
    khác "1" thì chỉnh về **0**.
  - **Chỉ số sinh tồn** (cột BA-BF: chiều cao 120-210, cân nặng 25-200, nhịp thở 12-20, mạch 60-100,
    huyết áp tâm thu 90-120, huyết áp tâm trương 60-80): nếu giá trị nằm **ngoài khoảng cho phép** thì
    báo **Cảnh báo** kèm ghi chú, và ô được tô màu trong file tải về.
  - **Các cột `*_chandoansobo_icd` / `*_chandoanxacdinh_icd`** (phạm vi cột BI đến DB): nếu giá trị là
    **0** thì chuyển thành **trống (null)**.
  - **Loại khám** (cột ED): nếu đang trống → điền mặc định **= 2**.
  - **Số lượng hồng cầu** (cột EE, `kskdk_xnm_slhc`): nếu đang trống → điền mặc định **= 0**.
  - **`*_tuchoikham`** (cột CV: `sankhoa_tuchoikham`, cột DA: `phukhoa_tuchoikham`) = 1: cột phân loại
    tương ứng (CZ: `sankhoa_phanloai`, DE: `phukhoa_phanloai`) chuyển thành **trống (null)**.
  - **Giới tính Nữ (`gioi_tinh` = 2)** và cột CV, CX, CY (`sankhoa_tuchoikham`, `sankhoa_chandoansobo_icd`,
    `sankhoa_chandoanxacdinh_icd`) đều đang trống → điền CW (`sankhoa_chuaphathienbatthuong`) = **1**
    và CZ (`sankhoa_phanloai`) = **1**.
  - **Giới tính Nữ (`gioi_tinh` = 2)** và cột DC, DD (`phukhoa_chandoansobo_icd`,
    `phukhoa_chandoanxacdinh_icd`) đều đang trống → điền DB (`phukhoa_chuaphathienbatthuong`) = **1**
    và DE (`phukhoa_phanloai`) = **1**.
  - **`nghenghiep_code`** (cột Q) hoặc **`noi_cong_tac`** (cột R) đang trống → điền **`doi_tuong_kham`**
    (cột C) = **3**.
- **Nhận diện cấu trúc file linh hoạt**: công cụ tự tìm dòng "mã field" (keyword) trong 15 dòng đầu
  của sheet thay vì cố định ở dòng 4 — chỉ cần file upload có dòng keyword giống file mẫu (dựa trên
  các mã quen thuộc như `ho_ten`, `dinh_danh_ca_nhan`, `ngay_kham`, `gioi_tinh`), dữ liệu sẽ được
  mapping đúng theo keyword bất kể nằm ở dòng số mấy. Nếu không tự nhận diện được, công cụ dùng mặc
  định dòng 4 và báo ghi chú ở phần "Ghi chú về cấu trúc file".
- **File Excel tải về không bị khoá/bảo vệ**: vẫn có thể filter, xoá dòng/cột, copy, paste bình
  thường như file gốc.
- **3 cặp đo thị lực Mắt** (không kính / kính lỗ / có kính, mỗi cặp gồm mắt phải + mắt trái): phải
  điền theo từng cặp (cùng có hoặc cùng trống); cặp "không kính" loại trừ với 2 cặp "kính lỗ" và
  "có kính" — điền cặp này thì không điền cặp kia.
- **`giadinh_macbenh` và `giadinh_danhsachbenh_icd`** chỉ ở mức **cảnh báo**, không chặn nhập liệu.
- **Khoảng trắng ẩn** (dấu cách không ngắt `\\xa0`, ký tự rộng-0...) trong bất kỳ ô chữ nào — dấu vết
  hay gặp khi copy dữ liệu từ web/PDF, từng gây lỗi "điền thành công giả" ở Nơi công tác. (Không áp
  dụng cho các ô cận lâm sàng — xem mục riêng ở trên.)
- **Cấu trúc file**: phát hiện nếu 2 cột vô tình dùng trùng 1 mã field (lỗi hiếm gặp trong file gốc).

Công cụ **không** kiểm tra: nội dung mô tả tự do (ghi chú, mô tả lâm sàng), và không có danh mục
đầy đủ cho Dân tộc / Hình thức chi trả nên các ô đó chỉ được kiểm tra "có dữ liệu hay không", không
đối chiếu danh mục.
            """
        )

# ============================================================
# TAB 2 — GHÉP NHIỀU FILE (tính năng mới)
# ============================================================
with tab_merge:
    st.caption(
        "Dùng khi có **nhiều file dữ liệu Mẫu 03** (ví dụ mỗi khóm/tổ/phường 1 file, cùng đúng khuôn "
        "mẫu Medinet) và muốn **gộp lại thành 1 file duy nhất** để chạy script Tampermonkey 1 lần cho "
        "cả khu vực, thay vì chạy tay từng file nhỏ. Các sheet danh mục (Nghề nghiệp, Nơi làm việc, "
        "Tỉnh, Phường xã...) trong file kết quả lấy theo **file ĐẦU TIÊN** trong danh sách tải lên — "
        "vì đây là danh mục dùng chung của Medinet, không phải dữ liệu riêng của từng đơn vị."
    )

    merge_files = st.file_uploader(
        "Chọn nhiều file Excel Mẫu 03 cần ghép (giữ Ctrl hoặc Cmd để chọn nhiều file cùng lúc)",
        type=["xlsx"], accept_multiple_files=True, key="merge_uploader",
    )

    if merge_files:
        if len(merge_files) < 2:
            st.info("Chọn ít nhất 2 file để ghép.")
        else:
            st.write("**Tên đơn vị cho từng file** (tự đoán từ tên file — sửa lại cho đúng nếu cần):")
            unit_labels = []
            for i, f in enumerate(merge_files):
                label = st.text_input(f"📄 {f.name}", value=default_unit_label(f.name), key=f"unit_label_{i}")
                unit_labels.append(label)

            if st.button("🔗 Ghép các file này lại", type="primary"):
                parsed, failed = [], []
                for f, label in zip(merge_files, unit_labels):
                    try:
                        parsed.append(read_source_file(f.getvalue(), f.name, label))
                    except Exception as e:
                        failed.append((f.name, str(e)))

                if failed:
                    st.error("Các file sau **không ghép được**, cần kiểm tra lại (có thể lỡ tải nhầm mẫu khác):")
                    for name, err in failed:
                        st.write(f"- **{name}**: {err}")

                if len(parsed) < 2:
                    st.warning("Chưa đủ ít nhất 2 file hợp lệ để ghép.")
                else:
                    if failed:
                        st.warning(f"Vẫn tiếp tục ghép **{len(parsed)} file hợp lệ**, bỏ qua {len(failed)} file lỗi ở trên.")
                    try:
                        merged_bytes, summary = merge_mau03_files(parsed)
                    except Exception as e:
                        st.error(f"Không ghép được: {e}")
                        st.stop()

                    mc1, mc2, mc3 = st.columns(3)
                    mc1.metric("Tổng số dòng đã ghép", summary["total_rows"])
                    mc2.metric("Số file đã ghép", len(parsed))
                    mc3.metric("Số CCCD trùng phát hiện", len(summary["duplicates"]))

                    st.write("**Số dòng lấy từ mỗi file:**")
                    st.dataframe(
                        pd.DataFrame(summary["per_file"], columns=["Tên file", "Đơn vị", "Số dòng"]),
                        use_container_width=True, hide_index=True,
                    )

                    if summary["layout_notes"]:
                        with st.expander("⚠ Ghi chú về cấu trúc file", expanded=False):
                            for name, note in summary["layout_notes"]:
                                st.warning(f"**{name}**: {note}")

                    if summary["field_gaps"]:
                        with st.expander(
                            "⚠ Có file thiếu 1 số mã field so với file nền (các dòng đó sẽ để trống ở đúng ô thiếu)",
                            expanded=True,
                        ):
                            for name, gaps in summary["field_gaps"]:
                                st.warning(f"**{name}**: thiếu `{'`, `'.join(gaps)}`")

                    if summary["duplicates"]:
                        st.warning(
                            f"⚠ Phát hiện **{len(summary['duplicates'])} số CCCD bị trùng** — file vẫn ghép "
                            "đầy đủ, các ô CCCD trùng được tô vàng + ghi chú trong file tải về. Tự kiểm tra "
                            "lại trước khi chạy script điền Medinet (CCCD trùng thật sẽ bị Medinet từ chối lưu)."
                        )
                        dup_rows = []
                        for d in summary["duplicates"]:
                            for r, unit_label, ho_ten in d["rows"]:
                                dup_rows.append({
                                    "CCCD": d["cccd"], "Họ tên": ho_ten, "Đơn vị": unit_label,
                                    "Dòng Excel (sau ghép)": r,
                                })
                        st.dataframe(pd.DataFrame(dup_rows), use_container_width=True, hide_index=True)
                    else:
                        st.success("Không phát hiện CCCD trùng giữa các file. 🎉")

                    st.download_button(
                        "⬇ Tải file đã ghép",
                        data=merged_bytes,
                        file_name=f"Mau03_da_ghep_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_merged",
                    )
                    st.caption(
                        "Có thể đưa file này qua tab \"🔍 Kiểm tra file\" ở trên để kiểm tra lại lần nữa "
                        "trước khi dùng script Tampermonkey."
                    )

    with st.expander("ℹ Công cụ ghép đang làm gì?"):
        st.markdown(
            """
- Đọc từng file theo **mã field** (không theo số cột) — dùng đúng cách nhận diện cấu trúc (dòng
  nhãn/dòng mã field/dòng dữ liệu) như tab "Kiểm tra file", nên vẫn ghép đúng dù cột ở 1 file nguồn
  có lệch vị trí đôi chút so với file nền.
- **File đầu tiên trong danh sách tải lên là NỀN**: giữ nguyên định dạng, 4 dòng đầu và toàn bộ các
  sheet danh mục (`DoiTuongKham`, `Tinh`, `PhuongXa`, `NgheNghiep`, `NoiLamViec`, `TienSuGiaDinh`...).
- Gộp dữ liệu bệnh nhân của **tất cả** file (kể cả file nền) theo đúng thứ tự đã tải lên, đánh lại
  **STT liên tục** từ 1.
- File nào **thiếu mã field bắt buộc** (`ho_ten`, `dinh_danh_ca_nhan`, `ngay_kham`, `gioi_tinh`) —
  ví dụ lỡ tải nhầm file của mẫu khác — sẽ **bị loại**, có báo lỗi rõ tên file, các file còn lại vẫn
  ghép bình thường.
- File thiếu 1 vài mã field khác (không thuộc nhóm bắt buộc trên) vẫn được ghép, chỉ để **trống**
  đúng những ô đó — có cảnh báo rõ file nào thiếu mã gì.
- **CCCD trùng** (giữa các file, hoặc trùng ngay trong 1 file): **không chặn**, vẫn ghép đầy đủ, chỉ
  tô vàng + ghi chú ô CCCD trong file tải về và liệt kê trong bảng cảnh báo trên màn hình.
- File kết quả có thêm **1 sheet phụ "Nguon_Ghep"** ghi rõ mỗi dòng bệnh nhân đến từ đơn vị/file nào
  — không đụng vào cấu trúc sheet dữ liệu chính (`ThongTinHanhChinh`), nên không ảnh hưởng tới script
  Tampermonkey hay tab "Kiểm tra file".
            """
        )
