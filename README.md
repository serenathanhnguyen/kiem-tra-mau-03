# Công cụ kiểm tra & ghép file Excel Mẫu 03 (KSKDK) — Medinet

App có **2 tab**:

1. **🔍 Kiểm tra file** — kiểm tra 1 file Excel dữ liệu bệnh nhân (đúng cấu trúc file mẫu Medinet
   dùng để nhập/import Mẫu 03) **trước khi** dùng script Tampermonkey điền vào Medinet — giúp phát
   hiện sớm CCCD sai/trùng, ngày sai định dạng, ô bắt buộc bị bỏ trống, giá trị không khớp danh mục
   (Tỉnh/Phường xã, Nghề nghiệp, Nơi công tác, Đối tượng khám...), khoảng trắng ẩn trong chữ...
2. **🔗 Ghép nhiều file** — gộp nhiều file Mẫu 03 (ví dụ mỗi khóm/tổ/phường 1 file) thành 1 file
   duy nhất, giữ nguyên cấu trúc, để chạy script Tampermonkey 1 lần cho cả khu vực. Đặt tên đơn vị
   cho từng file, tự đánh lại STT, cảnh báo (không chặn) nếu có CCCD trùng giữa các file, và ghi rõ
   mỗi dòng đến từ đơn vị nào trong 1 sheet phụ của file kết quả.

## File trong repo

- `m03app.py` — giao diện Streamlit (2 tab: upload file để kiểm tra, upload nhiều file để ghép).
- `validators.py` — toàn bộ logic kiểm tra (không phụ thuộc Streamlit, có thể test riêng).
- `merger.py` — toàn bộ logic ghép file, tái dùng phần nhận diện cấu trúc file của `validators.py`.
- `requirements.txt` — danh sách thư viện cần cài khi deploy (không đổi so với trước — tính năng
  ghép không cần thêm thư viện nào).

> ⚠ **Tên file phải viết đúng chữ thường `validators.py`.** Trong `m03app.py` có dòng
> `from validators import ...` — máy chủ Streamlit Cloud chạy Linux, phân biệt chữ hoa/thường, nên
> nếu file bị đặt tên `Validators.py` (V hoa) app sẽ báo lỗi `ModuleNotFoundError` và không chạy được,
> dù bạn tự thử trên Windows/Mac vẫn thấy bình thường (2 hệ đó không phân biệt hoa/thường).

## Đưa lên GitHub

1. Vào https://github.com → **New repository** (nút xanh "New" hoặc dấu **+** góc trên phải).
2. Đặt tên repo. Chọn **Public** (Streamlit Cloud miễn phí cần repo public, trừ khi tài khoản
   Streamlit của bạn có gói trả phí). Không cần tick "Add a README" nếu đã có sẵn.
3. Bấm **Create repository**.
4. Ở trang repo vừa tạo, bấm **"uploading an existing file"** (hoặc **Add file → Upload files**).
5. Kéo thả (hoặc chọn) cả 3 file `m03app.py`, `validators.py`, `requirements.txt` vào, rồi bấm
   **Commit changes**.

*(Nếu bạn quen dùng GitHub Desktop hoặc git dòng lệnh như các app trước, làm theo cách cũ cũng
được — không có gì khác biệt.)*

## Deploy lên Streamlit Community Cloud

1. Vào https://share.streamlit.io, đăng nhập bằng tài khoản GitHub (tài khoản bạn đã dùng cho các
   app trước — sales-dashboard, tyt-nhieu-loc-tools...).
2. Bấm **Create app** (hoặc **New app**).
3. Chọn **"Deploy a public app from GitHub"**, sau đó chọn:
   - **Repository**: `<tên-tài-khoản>/<tên-repo>`
   - **Branch**: `main`
   - **Main file path**: `m03app.py`
4. Bấm **Deploy**. Chờ khoảng 1–2 phút để Streamlit cài thư viện và khởi động app.
5. Xong sẽ có 1 link dạng `https://<tên-app>.streamlit.app` — dùng link này để mở công cụ mọi lúc,
   không cần cài gì trên máy.

## Cách dùng — tab 🔍 Kiểm tra file

1. Mở link app, đứng ở tab "🔍 Kiểm tra file" (mặc định).
2. Kéo thả (hoặc chọn) file Excel dữ liệu bệnh nhân (file `.xlsx` có sheet `ThongTinHanhChinh`
   đúng cấu trúc mẫu Medinet) vào ô tải file.
3. App tự chạy kiểm tra và hiện:
   - Số dòng đã kiểm tra / số lỗi / số cảnh báo.
   - Bảng chi tiết từng lỗi (tab đỏ — cần sửa trước khi nhập Medinet) và từng cảnh báo (tab vàng —
     nên xem lại nhưng không nhất thiết chặn).
   - Nút **"Tải báo cáo lỗi (Excel)"** để tải về, tiện gửi cho người nhập liệu sửa.

## Cách dùng — tab 🔗 Ghép nhiều file

1. Chuyển sang tab "🔗 Ghép nhiều file".
2. Chọn nhiều file Excel Mẫu 03 cùng lúc (giữ Ctrl/Cmd khi chọn file trong hộp thoại).
3. Sửa lại **tên đơn vị** cho từng file nếu tên tự đoán từ tên file chưa đúng (ví dụ "KP19",
   "KP2 - Đăng"...) — tên này chỉ để tra cứu sau này, không ảnh hưởng dữ liệu.
4. Bấm **"🔗 Ghép các file này lại"**. App hiện tổng số dòng đã ghép, số dòng lấy từ mỗi file, và
   cảnh báo nếu có CCCD trùng giữa các file (vẫn ghép đầy đủ, chỉ tô vàng để tự kiểm tra lại).
5. Bấm **"⬇ Tải file đã ghép"**. Có thể đưa file này quay lại tab "🔍 Kiểm tra file" để kiểm tra lần
   cuối trước khi dùng script Tampermonkey.

File đầu tiên trong danh sách tải lên được dùng làm nền (giữ định dạng + các sheet danh mục) — nếu
muốn 1 file cụ thể làm nền, tải file đó lên đầu tiên.

## Cập nhật sau này

Muốn sửa quy tắc kiểm tra (ví dụ thêm cột mới, đổi ngưỡng...), chỉ cần sửa `validators.py` ngay
trên GitHub (bấm vào file → biểu tượng bút chì để sửa → Commit changes) — Streamlit Cloud tự phát
hiện thay đổi và khởi động lại app sau vài chục giây, không cần làm gì thêm ở phía Streamlit.

## Những gì công cụ đang kiểm tra (và chưa kiểm tra)

Xem mục "ℹ Công cụ đang kiểm tra những gì?" ngay trong giao diện app — liệt kê đầy đủ các quy tắc
đang áp dụng, và các phần chưa có danh mục đầy đủ để đối chiếu (Dân tộc, Hình thức chi trả) nên chỉ
kiểm tra "có dữ liệu hay không" chứ chưa đối chiếu danh mục chặt.
