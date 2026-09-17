# Công cụ kiểm tra file Excel Mẫu 03 (KSKDK) — Medinet

Kiểm tra file Excel dữ liệu bệnh nhân (đúng cấu trúc file mẫu Medinet dùng để nhập/import Mẫu 03)
**trước khi** dùng script Tampermonkey điền vào Medinet — giúp phát hiện sớm CCCD sai/trùng, ngày
sai định dạng, ô bắt buộc bị bỏ trống, giá trị không khớp danh mục (Tỉnh/Phường xã, Nghề nghiệp,
Nơi công tác, Đối tượng khám...), khoảng trắng ẩn trong chữ...

## File trong repo

- `app.py` — giao diện Streamlit (upload file, hiển thị kết quả).
- `validators.py` — toàn bộ logic kiểm tra (không phụ thuộc Streamlit, có thể test riêng).
- `requirements.txt` — danh sách thư viện cần cài khi deploy.

## Đưa lên GitHub

1. Vào https://github.com → **New repository** (nút xanh "New" hoặc dấu **+** góc trên phải).
2. Đặt tên, ví dụ `mau03-excel-checker`. Chọn **Public** (Streamlit Cloud miễn phí cần repo public,
   trừ khi tài khoản Streamlit của bạn có gói trả phí). Không cần tick "Add a README" vì mình đã có.
3. Bấm **Create repository**.
4. Ở trang repo vừa tạo, bấm **"uploading an existing file"** (hoặc **Add file → Upload files**).
5. Kéo thả (hoặc chọn) cả 3 file `app.py`, `validators.py`, `requirements.txt` vào, rồi bấm
   **Commit changes**.

*(Nếu bạn quen dùng GitHub Desktop hoặc git dòng lệnh như các app trước, làm theo cách cũ cũng
được — không có gì khác biệt.)*

## Deploy lên Streamlit Community Cloud

1. Vào https://share.streamlit.io, đăng nhập bằng tài khoản GitHub (tài khoản bạn đã dùng cho các
   app trước — sales-dashboard, tyt-nhieu-loc-tools...).
2. Bấm **Create app** (hoặc **New app**).
3. Chọn **"Deploy a public app from GitHub"**, sau đó chọn:
   - **Repository**: `<tên-tài-khoản>/mau03-excel-checker`
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. Bấm **Deploy**. Chờ khoảng 1–2 phút để Streamlit cài thư viện và khởi động app.
5. Xong sẽ có 1 link dạng `https://<tên-app>.streamlit.app` — dùng link này để mở công cụ mọi lúc,
   không cần cài gì trên máy.

## Cách dùng

1. Mở link app.
2. Kéo thả (hoặc chọn) file Excel dữ liệu bệnh nhân (file `.xlsx` có sheet `ThongTinHanhChinh`
   đúng cấu trúc mẫu Medinet) vào ô tải file.
3. App tự chạy kiểm tra và hiện:
   - Số dòng đã kiểm tra / số lỗi / số cảnh báo.
   - Bảng chi tiết từng lỗi (tab đỏ — cần sửa trước khi nhập Medinet) và từng cảnh báo (tab vàng —
     nên xem lại nhưng không nhất thiết chặn).
   - Nút **"Tải báo cáo lỗi (Excel)"** để tải về, tiện gửi cho người nhập liệu sửa.

## Cập nhật sau này

Muốn sửa quy tắc kiểm tra (ví dụ thêm cột mới, đổi ngưỡng...), chỉ cần sửa `validators.py` ngay
trên GitHub (bấm vào file → biểu tượng bút chì để sửa → Commit changes) — Streamlit Cloud tự phát
hiện thay đổi và khởi động lại app sau vài chục giây, không cần làm gì thêm ở phía Streamlit.

## Những gì công cụ đang kiểm tra (và chưa kiểm tra)

Xem mục "ℹ Công cụ đang kiểm tra những gì?" ngay trong giao diện app — liệt kê đầy đủ các quy tắc
đang áp dụng, và các phần chưa có danh mục đầy đủ để đối chiếu (Dân tộc, Hình thức chi trả) nên chỉ
kiểm tra "có dữ liệu hay không" chứ chưa đối chiếu danh mục chặt.
