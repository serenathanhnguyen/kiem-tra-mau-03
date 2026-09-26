import re
from datetime import datetime, date, timedelta
from io import BytesIO

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment
from openpyxl.comments import Comment

# ============================================================
# CẤU HÌNH CẤU TRÚC FILE MẪU 03
# ============================================================
SHEET_MAIN = "ThongTinHanhChinh"
LABEL_ROW = 2          # dòng nhãn mặc định (có dấu * cho cột bắt buộc) — dùng khi không tự nhận diện được
CODE_ROW = 4            # dòng mã field (keyword) mặc định — dùng khi không tự nhận diện được
DATA_START_ROW = 5      # dữ liệu bệnh nhân bắt đầu từ dòng này (mặc định)

REF_SHEETS = {
    "DoiTuongKham": {"name_col": "Đối Tượng", "id_col": "Mã"},
    "Tinh": {"name_col": "CityName", "id_col": "CityId"},
    "PhuongXa": {"name_col": "WardName", "id_col": "WardId", "parent_col": "CityId"},
    "NgheNghiep": {"name_col": "Nghề Nghiệp", "id_col": "Mã"},
    "NoiLamViec": {"name_col": "Nơi làm việc", "id_col": "Mã"},
    "TienSuGiaDinh": {"name_col": "Bệnh Tiền Sử Gia Đình", "id_col": "ID"},
}

# ---- Các nhóm mã field theo quy tắc kiểm tra ----

CHOICE_ONE_OF = {
    "gioi_tinh": ["1", "2", "3"],
    "nhom_mau_id": ["A", "B", "O", "AB"],
    "yeu_to_nhom_mau_id": ["1", "2"],
    "loai_kham": ["1", "2"],
    "can_lam_sang_khac": ["1", "2"],
    "kskdk_xnnt_nitrit": ["0", "1"],
    "sankhoa_tuchoikham": ["0", "1"],
    "phukhoa_tuchoikham": ["0", "1"],
}

# Các câu hỏi tiền sử bệnh bản thân / gia đình dạng Có(1)/Không(0)
# LƯU Ý QUAN TRỌNG: đã ĐẢO NGƯỢC lại lần đổi tên trước (ts_than_kinh_dau, ts_mat...) — đối chiếu
# trực tiếp với file Excel THẬT Jo gửi (10_09_26_FILE_MAU_M3...) thì cột AB-AU vẫn dùng ĐÚNG các
# mã field CŨ này (benh_than_kinh, benh_mat...). Ảnh chụp + file "Danh_sach_ten_cot_keyword.xlsx"
# Jo gửi trước đó là một bảng đặt tên KHÁC, KHÔNG khớp với file thật đang dùng — đổi theo bảng đó
# đã khiến toàn bộ 20 câu này ngừng được kiểm tra trên file thật (mã field không khớp cột nào cả).
BINARY_01_FIELDS = [
    "benh_5nam", "benh_than_kinh", "benh_mat", "benh_tai", "benh_tim", "pt_tim_mach",
    "tang_ha", "kho_tho", "benh_phoi", "benh_than", "nghien_ruou_bia", "dai_thao_duong",
    "benh_tam_than", "mat_y_thuc", "ngat_chong_mat", "benh_tieu_hoa", "roi_loan_giac_ngu",
    "tai_bien_mach_mau_nao", "cot_song", "su_dung_ruou_bia", "su_dung_ma_tuy",
    "dieu_tri_benh_co_khong", "thai_san_co_khong",
]

# Các trường số KHÔNG thuộc cận lâm sàng (chỉ số sinh tồn + thị lực) — kiểm tra đầy đủ:
# vừa bắt lỗi dấu chấm/phẩy, vừa phải là số hợp lệ.
# LƯU Ý: đã ĐẢO NGƯỢC lại tên mat_thiluc_khongkinh_*/mat_thiluc_cokinh_* — file thật vẫn dùng
# mat_khongkinh_*/mat_cokinh_* (không có "thiluc_"); mat_kinhlo_* giữ nguyên vì vốn đã đúng.
NUMERIC_FIELDS = [
    "chieucao", "cannang", "nhiptho", "mach", "huyetaptamthu", "huyetaptamtruong",
    "mat_khongkinh_mp", "mat_khongkinh_mt", "mat_kinhlo_mp", "mat_kinhlo_mt",
    "mat_cokinh_mp", "mat_cokinh_mt", "mat_docau_mp", "mat_docau_mt",
    "mat_dotru_mp", "mat_dotru_mt", "mat_truc_mt",
]



# Các trường số THUỘC cận lâm sàng (xét nghiệm máu, sinh hóa máu, xét nghiệm nước tiểu).
# Theo yêu cầu của Jo: nếu có nhập dữ liệu thì CHỈ kiểm tra lỗi dùng dấu chấm (.) thay cho
# dấu phẩy (,) ở phần thập phân — không cảnh báo/báo lỗi gì khác (kể cả khi không phải số hợp lệ,
# hay dính khoảng trắng ẩn).
CANLAMSANG_NUMERIC_FIELDS = [
    "kskdk_xnm_slhc", "kskdk_xnm_huyetsacto", "kskdk_xnm_hematocrit", "kskdk_xnm_mcv",
    "kskdk_xnm_mch", "kskdk_xnm_mchc", "kskdk_xnm_rdw", "kskdk_xnm_slbc",
    "kskdk_xnm_slbc_trungtinh", "kskdk_xnm_slbc_lympho", "kskdk_xnm_slbc_donnhan",
    "kskdk_xnm_slbc_aitoan", "kskdk_xnm_slbc_aikiem", "kskdk_xnm_sltc",
    "kskdk_shm_duongmau", "kskdk_shm_ure", "kskdk_shm_creatinin", "kskdk_shm_asat_got",
    "kskdk_shm_alat_gpt", "kskdk_xnnt_titrong", "kskdk_xnnt_ph", "kskdk_xnnt_bachcau",
    "kskdk_xnnt_hongcau", "kskdk_xnnt_protein", "kskdk_xnnt_glucose", "kskdk_xnnt_cetonic",
    "kskdk_xnnt_bilirubin", "kskdk_xnnt_urobilinogen",
]

DATE_FIELDS = ["ngay_kham", "ngay_sinh"]

# Mã field cần khớp DANH MỤC CHẶT (sai là script điền tự động sẽ thất bại)
STRICT_TEXT_CATEGORY = {
    # code: (tên sheet tham chiếu, tên cột chứa giá trị hiển thị)
    "nghenghiep_code": ("NgheNghiep", "Nghề Nghiệp"),
    "noi_cong_tac": ("NoiLamViec", "Nơi làm việc"),
}

STRICT_ID_CATEGORY = {
    # code: (tên sheet tham chiếu, tên cột ID, cho phép nhiều giá trị cách nhau bởi dấu phẩy)
    "doi_tuong_kham": ("DoiTuongKham", "Mã", True),
    "giadinh_macbenh": ("TienSuGiaDinh", "ID", True),
}

ICD_RE = re.compile(r"^[A-TV-Z][0-9]{2}(\.[0-9]{1,2})?$", re.IGNORECASE)
CCCD_RE = re.compile(r"^\d{12}$")
PHONE_RE = re.compile(r"^0\d{9,10}$")

# Quy tắc mới (file quy tắc Jo gửi): 'sdt' phải đủ ĐÚNG 10 chữ số — không đúng 10 số (kể cả để
# trống, thiếu số, thừa số, hoặc lẫn ký tự khác) thì hệ thống tự điền giá trị mặc định này trong
# file tải về (giá trị lấy ĐÚNG NGUYÊN VĂN theo yêu cầu của Jo, kể cả khi bản thân nó cũng không
# đủ 10 số).
SDT_10_RE = re.compile(r"^\d{10}$")
SDT_DEFAULT_VALUE = "090900202"

NBSP_CHARS = [
    "\xa0", "​", "﻿", " ", " ",   # NBSP, zero-width space, BOM, figure space, narrow NBSP
    "‌", "‍", "⁠", "­",             # zero-width non-joiner/joiner, word joiner, soft hyphen
]

# 5 giá trị hợp lệ của ô Kết luận (đúng theo danh sách Jo cung cấp)
DANH_MUC_DE_NGHI_CHOICES = [
    "Bình thường, hẹn khám định kỳ lần sau",
    "Có yếu tố nguy cơ, cần theo dõi thêm",
    "Đã có bệnh mạn tính, tiếp tục điều trị theo phác đồ/toa cũ",
    "Chuyển tuyến, khám chuyên khoa",
    "Khác",
]

# Giá trị mặc định tự điền cho ô 'de_nghi' (Đề nghị, ghi rõ) khi đang để trống — theo yêu cầu Jo
DE_NGHI_DEFAULT_VALUE = "Tái khám định kỳ"

# 12 khối chuyên khoa dạng 4 ô: _chuaphathienbatthuong / _chandoansobo_icd / _chandoanxacdinh_icd / _phanloai
# Mã lấy ĐÚNG theo file gốc — thankinh có lỗi chính tả sẵn trong file ("chuandoansobo" thay vì "chandoansobo")
# Khối "mat" (Mắt) KHÔNG còn nằm trong bảng chung này nữa — Jo cho quy tắc RIÊNG (mức Cảnh báo, có
# tự điền H52.7...) khác hẳn 12 khối kia — xem check_mat_rules() + compute_data_fixes_for_row().
# Mã field khối Mắt VẪN đúng khuôn _chuaphathienbatthuong/_chandoansobo_icd/_chandoanxacdinh_icd/
# _phanloai như 12 khối này (mat_chuaphathienbatthuong, mat_chandoansobo_icd, mat_chandoanxacdinh_icd,
# mat_phanloai) — chỉ tách riêng vì QUY TẮC khác, không phải vì tên mã khác.
# Hệ quả: khối Mắt KHÔNG còn tự động góp vào gợi ý "Kết luận" (compute_danh_muc_de_nghi_for_row)
# hay quy tắc "có ICD xác định thật thì ép Kết luận = có bệnh mạn tính" như 12 khối còn lại — nếu
# Jo muốn Mắt vẫn góp vào 2 chỗ đó, nói mình bổ sung riêng.
SPECIALTY_BLOCKS_4FIELD = {
    "noikhoa": ("noikhoa_chuaphathienbatthuong", "noikhoa_chandoansobo_icd", "noikhoa_chandoanxacdinh_icd", "noikhoa_phanloai"),
    "hohap": ("hohap_chuaphathienbatthuong", "hohap_chandoansobo_icd", "hohap_chandoanxacdinh_icd", "hohap_phanloai"),
    "tieuhoa": ("tieuhoa_chuaphathienbatthuong", "tieuhoa_chandoansobo_icd", "tieuhoa_chandoanxacdinh_icd", "tieuhoa_phanloai"),
    "thantietnieu": ("thantietnieu_chuaphathienbatthuong", "thantietnieu_chandoansobo_icd", "thantietnieu_chandoanxacdinh_icd", "thantietnieu_phanloai"),
    "noitiet": ("noitiet_chuaphathienbatthuong", "noitiet_chandoansobo_icd", "noitiet_chandoanxacdinh_icd", "noitiet_phanloai"),
    "coxuongkhop": ("coxuongkhop_chuaphathienbatthuong", "coxuongkhop_chandoansobo_icd", "coxuongkhop_chandoanxacdinh_icd", "coxuongkhop_phanloai"),
    "thankinh": ("thankinh_chuaphathienbatthuong", "thankinh_chuandoansobo_icd", "thankinh_chandoanxacdinh_icd", "thankinh_phanloai"),
    "tamthan": ("tamthan_chuaphathienbatthuong", "tamthan_chandoansobo_icd", "tamthan_chandoanxacdinh_icd", "tamthan_phanloai"),
    "ngoaikhoa": ("ngoaikhoa_chuaphathienbatthuong", "ngoaikhoa_chandoansobo_icd", "ngoaikhoa_chandoanxacdinh_icd", "ngoaikhoa_phanloai"),
    "dalieu": ("dalieu_chuaphathienbatthuong", "dalieu_chandoansobo_icd", "dalieu_chandoanxacdinh_icd", "dalieu_phanloai"),
    "tmh": ("tmh_chuaphathienbatthuong", "tmh_chandoansobo_icd", "tmh_chandoanxacdinh_icd", "tmh_phanloai"),
    "rhm": ("rhm_chuaphathienbatthuong", "rhm_chandoansobo_icd", "rhm_chandoanxacdinh_icd", "rhm_phanloai"),
}

# Sản khoa / Phụ khoa: có thêm ô "từ chối khám" ở đầu — chỉ áp dụng cho nữ
SPECIALTY_BLOCKS_5FIELD = {
    "sankhoa": ("sankhoa_tuchoikham", "sankhoa_chuaphathienbatthuong", "sankhoa_chandoansobo_icd", "sankhoa_chandoanxacdinh_icd", "sankhoa_phanloai"),
    "phukhoa": ("phukhoa_tuchoikham", "phukhoa_chuaphathienbatthuong", "phukhoa_chandoansobo_icd", "phukhoa_chandoanxacdinh_icd", "phukhoa_phanloai"),
}

# Các ô chỉ dành cho nữ — nếu gioi_tinh = 1 (Nam) thì các ô này phải để trống
MALE_EXCLUDED_FIELDS = ["thai_san_co_khong", "thai_san_liet_ke"] + [
    f for fields in SPECIALTY_BLOCKS_5FIELD.values() for f in fields
]

DECIMAL_DOT_THOUSANDS_RE = re.compile(r"^-?\d{1,3}(\.\d{3})+$")

# Các ô hạ mức từ "Lỗi" xuống "Cảnh báo" (không chặn nhập liệu) — theo yêu cầu Jo
WARN_ONLY_FIELDS = {"giadinh_macbenh", "giadinh_danhsachbenh_icd"}

# 3 cặp đo thị lực Mắt — điền theo từng cặp (mp = mắt phải, mt = mắt trái)
# LƯU Ý: đã ĐẢO NGƯỢC lại tên mat_thiluc_khongkinh_*/mat_thiluc_cokinh_* — file thật vẫn dùng
# mat_khongkinh_*/mat_cokinh_* (không có "thiluc_"); mat_kinhlo_* giữ nguyên vì vốn đã đúng.
EYE_PAIRS = [
    ("khongkinh", ("mat_khongkinh_mp", "mat_khongkinh_mt")),  # cặp 1: không kính
    ("kinhlo", ("mat_kinhlo_mp", "mat_kinhlo_mt")),           # cặp 2: kính lỗ
    ("cokinh", ("mat_cokinh_mp", "mat_cokinh_mt")),           # cặp 3: có kính
]
EYE_PAIR_LABEL = {"khongkinh": "không kính", "kinhlo": "kính lỗ", "cokinh": "có kính"}

# Bảng phân loại THỂ LỰC theo QĐ 1613/BYT (Phụ lục 2). Mỗi chỉ số cho ngưỡng DƯỚI của Loại 1,2,3,4;
# thấp hơn ngưỡng Loại 4 → Loại 5. Loại thể lực = loại KÉM NHẤT (số lớn nhất) giữa chiều cao & cân nặng
# (file không có vòng ngực nên chỉ dùng 2 chỉ số).
THE_LUC_TABLES = {
    ("student", "nam"): {"height": [160, 156, 152, 149], "weight": [48, 46, 42, 39]},
    ("student", "nu"):  {"height": [152, 149, 145, 142], "weight": [44, 42, 40, 37]},
    ("worker", "nam"):  {"height": [160, 158, 154, 150], "weight": [50, 47, 45, 41]},
    ("worker", "nu"):   {"height": [155, 151, 147, 143], "weight": [45, 43, 40, 38]},
}

FILL_ERROR = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")   # đỏ nhạt
FILL_WARN = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")     # vàng nhạt

# ============================================================
# QUY TẮC "TỰ SỬA DỮ LIỆU" TRONG FILE XUẤT RA (annotate_workbook)
# Vị trí cột theo đúng file mẫu chuẩn Medinet (cột AA-AU, BA-BF, BI-DB, ED, EE...).
# Đa số quy tắc CHỈ ảnh hưởng file Excel tải về, KHÔNG ảnh hưởng bảng lỗi/cảnh báo (vẫn tính trên
# giá trị gốc) — RIÊNG 20 câu tiền sử bệnh AB-AU (BINARY_STRICT_FIELDS, không tính benh_5nam) và
# khối Mắt (mat_*, xem check_mat_rules) là NGOẠI LỆ: theo yêu cầu Jo, để trống các ô này vẫn PHẢI
# hiện Cảnh báo trong bảng kiểm tra, không chỉ âm thầm tự điền trong file tải về.
# ============================================================

# Cột AA-AU: 21 câu tiền sử bệnh dạng Có(1)/Không(0), mã field ĐÚNG THEO FILE THẬT (xem lưu ý ở
# BINARY_01_FIELDS — đã đảo ngược lại lần đổi tên sai trước đó). Giá trị hợp lệ CHỈ là 0 hoặc 1 —
# có dữ liệu mà khác "1" thì chỉnh về 0; ĐANG ĐỂ TRỐNG cũng chỉnh về 0 (yêu cầu mới của Jo — trước
# đây bỏ qua ô trống, giờ không còn nữa).
BINARY_STRICT_FIELDS = [
    "benh_5nam", "benh_than_kinh", "benh_mat", "benh_tai", "benh_tim", "pt_tim_mach",
    "tang_ha", "kho_tho", "benh_phoi", "benh_than", "nghien_ruou_bia", "dai_thao_duong",
    "benh_tam_than", "mat_y_thuc", "ngat_chong_mat", "benh_tieu_hoa", "roi_loan_giac_ngu",
    "tai_bien_mach_mau_nao", "cot_song", "su_dung_ruou_bia", "su_dung_ma_tuy",
]

# Cột BA-BF: khoảng giá trị hợp lệ cho các chỉ số sinh tồn — ngoài khoảng thì tô màu + ghi chú cảnh báo
VITAL_SIGN_RANGES = {
    "chieucao": (120, 210),
    "cannang": (25, 200),
    "nhiptho": (12, 20),
    "mach": (60, 100),
    "huyetaptamthu": (90, 120),
    "huyetaptamtruong": (60, 80),
}

# Nhận diện dòng "mã field" (keyword) một cách linh hoạt, không cố định ở dòng 4 — file upload chỉ
# cần có dòng keyword giống file mẫu (bất kể nằm ở dòng số mấy) thì dữ liệu vẫn được mapping đúng.
CODE_PATTERN_RE = re.compile(r"^[a-z][a-z0-9_]*$")
CODE_ROW_ANCHOR_FIELDS = {"ho_ten", "dinh_danh_ca_nhan", "ngay_kham", "gioi_tinh"}
LAYOUT_SEARCH_ROWS = 15  # số dòng đầu tiên quét để tìm dòng mã field


def _classify_metric(value, bounds):
    for i, b in enumerate(bounds):  # i=0 → Loại 1
        if value >= b:
            return i + 1
    return 5


def classify_the_luc(height, weight, gender, table):
    """Trả về loại thể lực 1-5, hoặc None nếu thiếu dữ liệu để tính.
    gender: 'nam'/'nu'; table: 'student'/'worker'."""
    if height is None or weight is None or gender not in ("nam", "nu"):
        return None
    tbl = THE_LUC_TABLES.get((table, gender))
    if not tbl:
        return None
    h = int(height + 0.5)   # làm tròn 0,5 lên theo QĐ 1613
    w = int(weight + 0.5)
    return max(_classify_metric(h, tbl["height"]), _classify_metric(w, tbl["weight"]))


def compute_the_luc_for_row(raw_by_code):
    """Suy loại thể lực cho 1 dòng từ chieucao/cannang + giới tính + đối tượng khám."""
    height = to_number(raw_by_code.get("chieucao"))
    weight = to_number(raw_by_code.get("cannang"))

    # Giới tính: ưu tiên ô gioi_tinh (1=Nam,2=Nữ); nếu trống thì suy từ ký tự thứ 4 của CCCD
    gt = clean_ws(raw_by_code.get("gioi_tinh"))
    gender = {"1": "nam", "2": "nu"}.get(gt)
    if gender is None:
        cccd = clean_ws(raw_by_code.get("dinh_danh_ca_nhan"))
        if CCCD_RE.match(cccd):
            gender = "nam" if int(cccd[3]) % 2 == 0 else "nu"

    # Bảng chuẩn theo đối tượng khám: mã 1 = học sinh/SV, mã 2 = người lao động, còn lại = người lao động
    dt_codes = [p.strip() for p in clean_ws(raw_by_code.get("doi_tuong_kham")).split(",") if p.strip()]
    table = "student" if "1" in dt_codes else "worker"

    return classify_the_luc(height, weight, gender, table)


def compute_danh_muc_de_nghi_for_row(raw_by_code):
    """Tự suy giá trị cho ô Kết luận ('danh_muc_de_nghi') khi đang để trống, dựa trên 13 khối chuyên
    khoa + Sản khoa/Phụ khoa (khối đã chọn 'từ chối khám' thì không tính vào điều kiện — không có dữ
    liệu để đánh giá). Quy tắc Jo bổ sung, theo đúng thứ tự (luôn cho ra 1 trong 3 giá trị, không để
    trống nữa):
      1) Nếu có ít nhất 1 cột '*_phanloai' liên quan là loại LỚN HƠN 1, VÀ có ít nhất 1 cột
         '*_chandoanxacdinh_icd' có giá trị
         → "Đã có bệnh mạn tính, tiếp tục điều trị theo phác đồ/toa cũ".
      2) Ngược lại, nếu có ít nhất 1 cột '*_phanloai' có giá trị (>=1, tức 1..5) VÀ có ít nhất 1 cột
         '*_chandoansobo_icd' có giá trị
         → "Có yếu tố nguy cơ, cần theo dõi thêm".
      3) Còn lại (không rơi vào 2 trường hợp trên)
         → "Bình thường, hẹn khám định kỳ lần sau"."""

    def blank(c):
        return is_blank(raw_by_code.get(c))

    def val(c):
        return clean_ws(raw_by_code.get(c))

    phanloai_codes, sobo_codes, xacdinh_codes = [], [], []

    for _check_c, sobo_c, xacdinh_c, phanloai_c in SPECIALTY_BLOCKS_4FIELD.values():
        phanloai_codes.append(phanloai_c)
        sobo_codes.append(sobo_c)
        xacdinh_codes.append(xacdinh_c)

    for tuchoi_c, _check_c, sobo_c, xacdinh_c, phanloai_c in SPECIALTY_BLOCKS_5FIELD.values():
        if (not blank(tuchoi_c)) and val(tuchoi_c) == "1":
            continue  # đã từ chối khám — khối này không tính vào điều kiện
        phanloai_codes.append(phanloai_c)
        sobo_codes.append(sobo_c)
        xacdinh_codes.append(xacdinh_c)

    has_phanloai_gt1 = any(val(c) in ("2", "3", "4", "5") for c in phanloai_codes)
    has_phanloai_ge1 = any(val(c) in ("1", "2", "3", "4", "5") for c in phanloai_codes)
    has_xacdinh = any(not blank(c) for c in xacdinh_codes)
    has_sobo = any(not blank(c) for c in sobo_codes)

    if has_phanloai_gt1 and has_xacdinh:
        return "Đã có bệnh mạn tính, tiếp tục điều trị theo phác đồ/toa cũ"

    if has_phanloai_ge1 and has_sobo:
        return "Có yếu tố nguy cơ, cần theo dõi thêm"

    return "Bình thường, hẹn khám định kỳ lần sau"


# ============================================================
# QUY TẮC RIÊNG CHO KHỐI MẮT
# Mã field ĐÚNG THEO FILE THẬT: mat_chuaphathienbatthuong (DF) / mat_chandoansobo_icd (DG) /
# mat_chandoanxacdinh_icd (DH) / mat_phanloai (DI) — cùng khuôn 4 ô như 12 khối kia (KHÔNG phải
# mat_binhthuong/mat_chandoansobo/mat_chandoanxacdinh như lần đổi tên trước — đã đảo ngược lại,
# xem lưu ý ở BINARY_01_FIELDS). Jo bổ sung quy tắc RIÊNG (mức Cảnh báo, có tự điền H52.7...) khác
# hẳn 12 khối kia nên tách hàm riêng — dùng chung cho cả bước validate (check_mat_rules) lẫn bước
# tự sửa dữ liệu (compute_data_fixes_for_row) để không có 2 nơi định nghĩa khác nhau cùng 1 điều kiện.
# ============================================================

def _mat_phanloai_num(raw_by_code):
    return to_number(raw_by_code.get("mat_phanloai"))


def _mat_icd_both_blank(raw_by_code):
    return is_blank(raw_by_code.get("mat_chandoansobo_icd")) and is_blank(raw_by_code.get("mat_chandoanxacdinh_icd"))


def check_mat_rules(raw_by_code):
    """Quy tắc riêng cho khối Mắt (theo file quy tắc mới của Jo, 24/09/2026 — ĐÃ BỎ HẲN điều kiện
    thị lực có kính < 8/10 dùng trước đây) — TẤT CẢ đều ở mức Cảnh báo, không chặn Lưu:
    - mat_phanloai (DI) bắt buộc phải có giá trị (1-5); để trống → Cảnh báo.
    - mat_phanloai (DI) = 2 VÀ cả 2 ô chẩn đoán (sơ bộ + xác định) đều đang trống → Cảnh báo, hệ
      thống sẽ tự xoá trắng mat_chuaphathienbatthuong và tự điền 'H52.7' vào mat_chandoansobo_icd
      trong file tải về — dùng chung điều kiện với compute_data_fixes_for_row() để 2 bên luôn khớp.
    - mat_phanloai (DI) > 2 mà cả 2 ô chẩn đoán đều trống → Cảnh báo riêng, cần bác sĩ tự bổ sung
      (hệ thống không tự điền được, quy tắc chỉ tự điền sẵn cho đúng trường hợp =2)."""
    issues = []
    di_raw = raw_by_code.get("mat_phanloai")
    if is_blank(di_raw):
        issues.append({
            "code": "mat_phanloai", "level": "Cảnh báo",
            "message": "Bắt buộc phải có giá trị (1 đến 5) — đang để trống",
        })
        return issues

    di_num = _mat_phanloai_num(raw_by_code)
    di_text = clean_ws(di_raw)
    icd_both_blank = _mat_icd_both_blank(raw_by_code)

    if di_num == 2 and icd_both_blank:
        issues.append({
            "code": "mat_chandoansobo_icd", "level": "Cảnh báo",
            "message": (f"Đã chọn phân loại {di_text} và cả 2 ô chẩn đoán đang trống — hệ thống sẽ "
                        "tự xoá trắng 'Chưa phát hiện bất thường' và tự điền 'H52.7' vào 'Chẩn đoán "
                        "sơ bộ' trong file tải về, tự kiểm tra lại cho đúng"),
        })
    elif di_num is not None and di_num > 2 and icd_both_blank:
        issues.append({
            "code": "mat_chandoansobo_icd", "level": "Cảnh báo",
            "message": (f"Đã chọn phân loại {di_text} (lớn hơn 2) nhưng 'Chẩn đoán sơ bộ'/"
                        "'Chẩn đoán xác định' đang trống — cần bác sĩ tự bổ sung mã ICD, hệ thống "
                        "không tự điền được cho trường hợp này"),
        })
    return issues


# ============================================================
# QUY TẮC RIÊNG: SỬA GIỚI TÍNH + NĂM SINH (file quy tắc Jo gửi 24/09/2026)
# Thứ tự ưu tiên: (1) Họ tên có chữ "Thị" -> Nữ (2), có chữ "Văn" -> Nam (1); (2) sau đó đối
# chiếu ký tự thứ 4 của CCCD với giới tính (đã áp bước 1) — lệch thì sửa theo CCCD; (3) đối
# chiếu năm sinh với CCCD — lệch thì sửa năm sinh theo CCCD. MỌI chỗ đã tự sửa đều phải hiện
# Cảnh báo (yêu cầu rõ của Jo) — dùng chung 1 hàm cho cả bước tự sửa dữ liệu
# (compute_data_fixes_for_row) lẫn bước hiện cảnh báo (check_gioitinh_ngaysinh_rules) để không
# có 2 nơi định nghĩa khác nhau.
# ============================================================

def compute_gioitinh_ngaysinh_fixes(raw_by_code):
    """Trả về (fixes, warnings, effective_gioi_tinh):
    - fixes: dict {mã field: giá trị mới} cần ghi đè (gioi_tinh và/hoặc ngay_sinh).
    - warnings: list (mã field, thông báo) cho mọi chỗ đã tự sửa.
    - effective_gioi_tinh: giá trị giới tính SAU KHI áp các bước sửa ở trên (dùng cho các quy
      tắc tự sửa dữ liệu khác phụ thuộc giới tính, vì Jo yêu cầu 'sửa giới tính trước')."""
    fixes = {}
    warnings = []

    def val(c):
        return clean_ws(raw_by_code.get(c))

    ho_ten = clean_ws(raw_by_code.get("ho_ten"))
    original_gt = val("gioi_tinh")
    effective_gt = original_gt

    # 1) Họ tên có chữ đệm "Thị" -> Nữ (2); có chữ đệm "Văn" -> Nam (1).
    ten_tokens = set(norm_key(ho_ten).split(" "))
    if "thị" in ten_tokens and effective_gt != "2":
        fixes["gioi_tinh"] = 2
        warnings.append((
            "gioi_tinh",
            f"Họ tên có chữ đệm 'Thị' — tự sửa Giới tính thành Nữ (2), giá trị gốc là "
            f"'{original_gt or '(trống)'}'",
        ))
        effective_gt = "2"
    elif "văn" in ten_tokens and effective_gt != "1":
        fixes["gioi_tinh"] = 1
        warnings.append((
            "gioi_tinh",
            f"Họ tên có chữ đệm 'Văn' — tự sửa Giới tính thành Nam (1), giá trị gốc là "
            f"'{original_gt or '(trống)'}'",
        ))
        effective_gt = "1"

    # 2) Đối chiếu với CCCD — ký tự thứ 4: chẵn -> Nam (1), lẻ -> Nữ (2).
    cccd = val("dinh_danh_ca_nhan")
    d4 = None
    if CCCD_RE.match(cccd):
        d4 = int(cccd[3])
        expected_gt = "1" if d4 % 2 == 0 else "2"
        if effective_gt in ("1", "2") and effective_gt != expected_gt:
            fixes["gioi_tinh"] = int(expected_gt)
            ten = "Nam" if expected_gt == "1" else "Nữ"
            warnings.append((
                "gioi_tinh",
                f"xem lại số CCCD ko đúng với GT (ký tự thứ 4 '{cccd[3]}' cho biết là {ten}) — tự "
                f"sửa lại Giới tính, giá trị trước khi sửa là '{effective_gt}'",
            ))
            effective_gt = expected_gt

    # 3) Đối chiếu năm sinh với CCCD (ký tự 4-6) — lệch thì sửa lại năm sinh theo CCCD.
    if d4 is not None:
        yy = int(cccd[4:6])
        century = 1900 + (d4 // 2) * 100
        year_cccd = century + yy
        d, err = parse_date_cell(raw_by_code.get("ngay_sinh"))
        if d is not None and d.year != year_cccd:
            try:
                new_date = d.replace(year=year_cccd)
            except ValueError:
                new_date = d.replace(year=year_cccd, day=28)  # phòng ca 29/2 rơi vào năm không nhuận
            fixes["ngay_sinh"] = new_date
            warnings.append((
                "ngay_sinh",
                f"Năm sinh ({d.year}) không khớp CCCD (ký tự 4-6 '{cccd[3:6]}' cho biết năm sinh "
                f"phải là {year_cccd}) — tự sửa lại Ngày sinh thành {new_date.strftime('%d/%m/%Y')}",
            ))

    return fixes, warnings, effective_gt


def check_gioitinh_ngaysinh_rules(raw_by_code):
    """Sinh Cảnh báo cho MỌI chỗ compute_gioitinh_ngaysinh_fixes() đã tự sửa — dùng chung điều
    kiện với hàm đó để bảng cảnh báo luôn khớp đúng với file Excel tự sửa tải về."""
    _fixes, warnings, _gt = compute_gioitinh_ngaysinh_fixes(raw_by_code)
    return [
        {"code": code, "level": "Cảnh báo", "message": f"{msg} (đã tự sửa trong file tải về)"}
        for code, msg in warnings
    ]


def compute_data_fixes_for_row(raw_by_code):
    """Trả về dict {mã field: giá trị mới} cần ghi đè vào file Excel xuất ra cho 1 dòng — các quy
    tắc 'tự sửa dữ liệu' Jo bổ sung (chỉ áp dụng cho file tải về, KHÔNG ảnh hưởng bảng lỗi/cảnh báo
    vốn vẫn tính trên giá trị gốc đã upload). Giá trị None trong dict nghĩa là XOÁ TRẮNG ô đó.
      - gioi_tinh = 1 (Nam): xoá trắng các ô chỉ dành cho nữ (thai_san_co_khong, thai_san_liet_ke,
        và toàn bộ khối Sản khoa/Phụ khoa — cột AY, AZ, CV..DE).
      - gioi_tinh = 2 (Nữ): ô 'thai_san_co_khong' (cột AY) nếu KHÁC 1 (để trống, "Không", 0, hay
        bất kỳ giá trị nào khác 1) thì chỉnh thành 0.
      - Cột AB-AU (20 câu tiền sử bệnh 0/1, mã field ts_*): đã điền giá trị khác "1" thì chỉnh về 0;
        ĐANG TRỐNG cũng chỉnh về 0 (yêu cầu mới — trước đây bỏ qua ô trống).
      - Các cột '*_chandoansobo_icd' / '*_chandoanxacdinh_icd' (cột BI-DB, 12 khối chuyên khoa +
        Sản khoa/Phụ khoa — KHÔNG còn gồm khối Mắt, xem quy tắc riêng bên dưới): giá trị 0 thì
        chuyển null.
      - 'loai_kham' (cột ED): trống thì điền = 2.
      - 'kskdk_xnm_slhc' — Số lượng hồng cầu (cột EE): trống thì điền = 0.
      - gioi_tinh = 2 (Nữ) và 'sankhoa_tuchoikham' (cột CV) = 1 → NULL cả 4 ô còn lại của khối Sản
        khoa: CW (chuaphathienbatthuong), CX (chandoansobo_icd), CY (chandoanxacdinh_icd), CZ
        (phanloai) — không còn cảnh báo lỗi phân loại 1-5 cho khối này nữa.
      - gioi_tinh = 2 (Nữ) và 'sankhoa_tuchoikham' <> 1 và CW, CX, CY đều đang trống → chỉnh CZ
        (sankhoa_phanloai) = 1 (còn lại giữ nguyên mặc định của file Excel).
      - gioi_tinh = 2 (Nữ) và 'phukhoa_tuchoikham' (cột DA) = 1 → NULL cả 4 ô còn lại của khối Phụ
        khoa: DB (chuaphathienbatthuong), DC (chandoansobo_icd), DD (chandoanxacdinh_icd), DE
        (phanloai) — không còn cảnh báo lỗi phân loại 1-5 cho khối này nữa.
      - gioi_tinh = 2 (Nữ) và 'phukhoa_tuchoikham' <> 1 và cột DC, DD đều đang trống → chỉnh DB
        (phukhoa_chuaphathienbatthuong) = 1 và DE (phukhoa_phanloai) = 1 (còn lại giữ nguyên mặc
        định của file Excel). Hai khối Sản khoa/Phụ khoa được xét ĐỘC LẬP với nhau — khối này từ
        chối khám không còn ảnh hưởng tới khối kia nữa.
      - sdt: phải đủ đúng 10 chữ số — không đúng 10 số (kể cả để trống) thì Cảnh báo và tự điền mặc
        định "090900202" trong file tải về (xem SDT_DEFAULT_VALUE).
      - 'nghenghiep_code' (cột Q) hoặc 'noi_cong_tac' (cột R) đang trống → điền 'doi_tuong_kham'
        (cột C) = 3.
      - 'doi_tuong_kham' (cột C) = 3 (kể cả trường hợp vừa được tự điền = 3 ở quy tắc ngay trên) →
        điền cố định 'hinh_thuc_chi_tra_khamsk' (cột V) = "Ngân sách thành phố hỗ trợ" và
        'hinh_thuc_chi_tra_khamsk_chi_tiet' (cột W) = "Khám Theo Hợp Đồng".
      - 1 khối chuyên khoa (4 hoặc 5 ô, KHÔNG gồm Mắt) có '*_chandoansobo_icd' hoặc
        '*_chandoanxacdinh_icd' khác 0 (có giá trị ICD thật) → ô '*_chuaphathienbatthuong' (ô check
        "Chưa phát hiện bất thường") của đúng khối đó chuyển null — tự xoá khi đang mâu thuẫn dữ liệu.
      - 12 khối chuyên khoa 4-ô (noikhoa, hohap, tieuhoa, thantietnieu, noitiet, coxuongkhop,
        thankinh, tamthan, ngoaikhoa, dalieu, tmh, rhm — KHÔNG gồm Mắt/Sản khoa/Phụ khoa): nếu ô
        chuaphathienbatthuong đang TRỐNG HOẶC = 0, và 3 ô còn lại (chandoansobo_icd,
        chandoanxacdinh_icd, phanloai) đều đang trống → tự điền chuaphathienbatthuong = 1 VÀ
        phanloai = 1 cho đúng khối đó.
      - Bất kỳ ô '*_chandoanxacdinh_icd' nào (12 khối chuyên khoa + Sản khoa/Phụ khoa, KHÔNG gồm
        Mắt) có giá trị KHÁC 0 và KHÁC 1 (tức có mã ICD thật, không phải giá trị rác/placeholder) →
        điền cố định 'danh_muc_de_nghi' (cột FP, "Kết Luận") = "Đã có bệnh mạn tính, tiếp tục điều
        trị theo phác đồ/toa cũ" — ghi đè cả khi ô này đã có giá trị khác.
      - Khối Mắt (mat_chuaphathienbatthuong/mat_chandoansobo_icd/mat_chandoanxacdinh_icd/
        mat_phanloai) — quy tắc RIÊNG (đã cập nhật theo file quy tắc 24/09/2026, BỎ điều kiện thị
        lực), xem check_mat_rules(): mat_phanloai=1 và cả 2 ô chẩn đoán trống → điền
        mat_chuaphathienbatthuong=1; mat_phanloai=2 và cả 2 ô chẩn đoán trống → xoá trắng
        mat_chuaphathienbatthuong và điền mat_chandoansobo_icd='H52.7'; mat_phanloai và
        mat_chuaphathienbatthuong đều đang trống → điền cả 2 = 1; các trường hợp khác (phanloai>2
        mà chẩn đoán trống) chỉ cảnh báo, không tự điền.
      - Giới tính (gioi_tinh) và Ngày sinh (ngay_sinh): áp dụng TRƯỚC TIÊN theo
        compute_gioitinh_ngaysinh_fixes() (họ tên có 'Thị'/'Văn', rồi đối chiếu CCCD) — các quy
        tắc theo giới tính bên dưới (xoá ô nữ khi Nam, tự điền Sản khoa/Phụ khoa khi Nữ...) dùng
        GIÁ TRỊ GIỚI TÍNH ĐÃ SỬA, không dùng giá trị gốc trong Excel.
    """
    fixes = {}

    def blank(c):
        return is_blank(raw_by_code.get(c))

    def val(c):
        return clean_ws(raw_by_code.get(c))

    # Sửa giới tính + năm sinh TRƯỚC (theo file quy tắc Jo gửi 24/09/2026) — các quy tắc theo giới
    # tính bên dưới đều phải dùng giá trị ĐÃ SỬA (effective_gt), không dùng giá trị gốc trong Excel.
    gt_fixes, _gt_warnings, gt = compute_gioitinh_ngaysinh_fixes(raw_by_code)
    fixes.update(gt_fixes)

    if gt == "1":
        for code in MALE_EXCLUDED_FIELDS:
            if not blank(code):
                fixes[code] = None
    elif gt == "2":
        # Yêu cầu mới của Jo: gioi_tinh=2 (Nữ) và thai_san_co_khong KHÁC 1 (kể cả để trống, "Không",
        # 0, hay bất kỳ giá trị nào khác 1) → điền = 0 (trước đây chỉ nhận diện đúng chữ "Không").
        if val("thai_san_co_khong") != "1":
            fixes["thai_san_co_khong"] = 0

    # Đã có dữ liệu mà khác "1" → chỉnh về 0; ĐANG TRỐNG cũng chỉnh về 0 (yêu cầu mới của Jo —
    # trước đây bỏ qua ô trống, xem cảnh báo tương ứng ở check_cell()).
    for code in BINARY_STRICT_FIELDS:
        if val(code) != "1":
            fixes[code] = 0

    # Quy tắc RIÊNG cho khối Mắt (theo file quy tắc Jo gửi 24/09/2026, ĐÃ BỎ điều kiện thị lực —
    # THAY THẾ hoàn toàn cách xử lý theo khuôn 4-ô chung ở dưới, mat_* không nằm trong
    # SPECIALTY_BLOCKS_4FIELD):
    #  - mat_phanloai (DI) và mat_chuaphathienbatthuong (DF) đều đang trống → điền cả 2 = 1.
    #  - mat_phanloai (DI) = 1 và cả 2 ô chẩn đoán (DG/DH) đang trống → điền DF = 1.
    #  - mat_phanloai (DI) = 2 và cả 2 ô chẩn đoán (DG/DH) đang trống → xoá trắng DF (không thể vừa
    #    "bất thường" vừa "chưa phát hiện bất thường"), và điền DG = "H52.7".
    #  - mat_phanloai > 2 mà cả DG/DH trống → KHÔNG tự điền gì (chỉ cảnh báo ở check_mat_rules(),
    #    để bác sĩ tự bổ sung — quy tắc mới chỉ tự điền sẵn cho đúng trường hợp =2).
    #  - DG/DH = 0 (giá trị rác) → chuyển null, giống quy tắc chung áp dụng cho 12 khối kia.
    di_num = _mat_phanloai_num(raw_by_code)
    mat_icd_both_blank = _mat_icd_both_blank(raw_by_code)
    if blank("mat_phanloai") and blank("mat_chuaphathienbatthuong"):
        fixes["mat_phanloai"] = 1
        fixes["mat_chuaphathienbatthuong"] = 1
    elif di_num == 1 and mat_icd_both_blank:
        fixes["mat_chuaphathienbatthuong"] = 1
    if di_num == 2 and mat_icd_both_blank:
        fixes["mat_chuaphathienbatthuong"] = None
        fixes["mat_chandoansobo_icd"] = "H52.7"
    for code in ("mat_chandoansobo_icd", "mat_chandoanxacdinh_icd"):
        if code not in fixes and _is_literal_zero(raw_by_code.get(code)):
            fixes[code] = None

    icd_fix_codes = []
    for _check_c, sobo_c, xacdinh_c, _phanloai_c in SPECIALTY_BLOCKS_4FIELD.values():
        icd_fix_codes.extend([sobo_c, xacdinh_c])
    for _tuchoi_c, _check_c, sobo_c, xacdinh_c, _phanloai_c in SPECIALTY_BLOCKS_5FIELD.values():
        icd_fix_codes.extend([sobo_c, xacdinh_c])
    for code in icd_fix_codes:
        if code in fixes:
            continue  # đã bị xoá trắng ở quy tắc giới tính (Nam) — khỏi ghi đè lại
        if _is_literal_zero(raw_by_code.get(code)):
            fixes[code] = None

    # Quy tắc mới (Jo bổ sung): 1 khối chuyên khoa có ICD (sơ bộ hoặc xác định) mang giá trị THẬT
    # (khác 0, khác trống) thì ô check "Chưa phát hiện bất thường" của đúng khối đó phải là null —
    # tự xoá nếu đang mâu thuẫn (đang chọn "Chưa phát hiện bất thường" nhưng lại có ICD).
    def has_icd_value(code):
        return (not blank(code)) and not _is_literal_zero(raw_by_code.get(code))

    icd_check_groups = [(c, s, x) for c, s, x, _p in SPECIALTY_BLOCKS_4FIELD.values()]
    icd_check_groups += [(c, s, x) for _t, c, s, x, _p in SPECIALTY_BLOCKS_5FIELD.values()]
    for check_c, sobo_c, xacdinh_c in icd_check_groups:
        if check_c in fixes:
            continue  # đã bị xoá trắng ở quy tắc giới tính (Nam) — khỏi ghi đè lại
        if has_icd_value(sobo_c) or has_icd_value(xacdinh_c):
            fixes[check_c] = None

    # Quy tắc mới (Jo bổ sung, file "1_yeu_cau.txt"): với 12 khối chuyên khoa 4-ô (KHÔNG gồm Mắt/
    # Sản khoa/Phụ khoa) — nếu ô chuaphathienbatthuong đang TRỐNG HOẶC = 0, VÀ 3 ô còn lại
    # (chandoansobo_icd, chandoanxacdinh_icd, phanloai) đều đang trống, thì tự điền
    # chuaphathienbatthuong = 1 VÀ phanloai = 1 cho đúng khối đó (áp dụng cho cả 12 khối như nhau,
    # không phân biệt giới tính).
    for check_c, sobo_c, xacdinh_c, phanloai_c in SPECIALTY_BLOCKS_4FIELD.values():
        check_blank_or_zero = blank(check_c) or _is_literal_zero(raw_by_code.get(check_c))
        if check_blank_or_zero and blank(sobo_c) and blank(xacdinh_c) and blank(phanloai_c):
            fixes[check_c] = 1
            fixes[phanloai_c] = 1

    # Quy tắc mới (Jo bổ sung): bất kỳ ô '*_chandoanxacdinh_icd' nào (13 khối + Sản khoa/Phụ khoa)
    # có giá trị KHÁC 0 VÀ KHÁC 1 (mã ICD thật, không phải 0/1 rác) → ép cứng Kết luận
    # ('danh_muc_de_nghi', cột FP) = "Đã có bệnh mạn tính...", ghi đè cả giá trị đang có sẵn.
    def has_real_xacdinh(code):
        if blank(code):
            return False
        if _is_literal_zero(raw_by_code.get(code)):
            return False
        n = to_number(raw_by_code.get(code))
        if n is not None and n == 1:
            return False
        return True

    xacdinh_codes_all = [x for _c, _s, x in icd_check_groups]
    if any(has_real_xacdinh(code) for code in xacdinh_codes_all):
        fixes["danh_muc_de_nghi"] = "Đã có bệnh mạn tính, tiếp tục điều trị theo phác đồ/toa cũ"

    if blank("loai_kham"):
        fixes["loai_kham"] = 2

    if blank("kskdk_xnm_slhc"):
        fixes["kskdk_xnm_slhc"] = 0

    # Sản khoa/Phụ khoa — quy tắc XÉT RIÊNG TỪNG KHỐI (theo file quy tắc mới Jo gửi — thay thế cách
    # gộp chung cả 2 khối trước đây, vốn lỡ null luôn phanloai của khối KHÔNG từ chối khi chỉ 1
    # trong 2 khối chọn 'Từ chối khám'), chỉ áp dụng khi Nữ (gt đã sửa ở trên):
    #  - sankhoa_tuchoikham = 1 → NULL cả 4 ô còn lại của khối Sản khoa (chuaphathienbatthuong,
    #    chandoansobo_icd, chandoanxacdinh_icd, phanloai) — không cảnh báo lỗi phân loại 1-5 nữa
    #    (xem 'skip_phanloai_range' ở check_cell()).
    #  - sankhoa_tuchoikham <> 1 → nếu sankhoa_chuaphathienbatthuong, sankhoa_chandoansobo_icd VÀ
    #    sankhoa_chandoanxacdinh_icd đều đang trống thì điền sankhoa_phanloai = 1, còn lại (đã có dữ
    #    liệu ở 1 trong 3 ô đó) giữ nguyên giá trị mặc định của file Excel, không tự điền gì.
    #  - phukhoa_tuchoikham = 1 → NULL cả 4 ô còn lại của khối Phụ khoa, tương tự Sản khoa ở trên.
    #  - phukhoa_tuchoikham <> 1 → nếu phukhoa_chandoansobo_icd VÀ phukhoa_chandoanxacdinh_icd đều
    #    đang trống (không xét ô chuaphathienbatthuong) thì điền phukhoa_chuaphathienbatthuong = 1
    #    VÀ phukhoa_phanloai = 1, còn lại giữ nguyên giá trị mặc định của file Excel.
    sankhoa_tuchoi_c, sankhoa_check_c, sankhoa_sobo_c, sankhoa_xacdinh_c, sankhoa_phanloai_c = \
        SPECIALTY_BLOCKS_5FIELD["sankhoa"]
    phukhoa_tuchoi_c, phukhoa_check_c, phukhoa_sobo_c, phukhoa_xacdinh_c, phukhoa_phanloai_c = \
        SPECIALTY_BLOCKS_5FIELD["phukhoa"]

    if gt == "2":
        sankhoa_tuchoi1 = (not blank(sankhoa_tuchoi_c)) and val(sankhoa_tuchoi_c) == "1"
        if sankhoa_tuchoi1:
            fixes[sankhoa_check_c] = None
            fixes[sankhoa_sobo_c] = None
            fixes[sankhoa_xacdinh_c] = None
            fixes[sankhoa_phanloai_c] = None
        elif blank(sankhoa_check_c) and blank(sankhoa_sobo_c) and blank(sankhoa_xacdinh_c):
            fixes[sankhoa_phanloai_c] = 1

        phukhoa_tuchoi1 = (not blank(phukhoa_tuchoi_c)) and val(phukhoa_tuchoi_c) == "1"
        if phukhoa_tuchoi1:
            fixes[phukhoa_check_c] = None
            fixes[phukhoa_sobo_c] = None
            fixes[phukhoa_xacdinh_c] = None
            fixes[phukhoa_phanloai_c] = None
        elif blank(phukhoa_sobo_c) and blank(phukhoa_xacdinh_c):
            fixes[phukhoa_check_c] = 1
            fixes[phukhoa_phanloai_c] = 1

    if blank("nghenghiep_code") or blank("noi_cong_tac"):
        fixes["doi_tuong_kham"] = 3

    # Quy tắc mới (Jo bổ sung): Đối tượng khám (cột C) = 3 → điền cố định Hình thức chi trả khám
    # sức khỏe (cột V) + Hình thức nhà nước hỗ trợ (cột W). Dùng giá trị SAU KHI áp quy tắc ngay
    # trên, để vẫn tính cả trường hợp cột C gốc đang trống nhưng vừa được tự điền = 3.
    effective_doi_tuong = fixes.get("doi_tuong_kham", raw_by_code.get("doi_tuong_kham"))
    dt_codes = [p.strip() for p in clean_ws(effective_doi_tuong).split(",") if p.strip()]
    if "3" in dt_codes:
        fixes["hinh_thuc_chi_tra_khamsk"] = "Ngân sách thành phố hỗ trợ"
        fixes["hinh_thuc_chi_tra_khamsk_chi_tiet"] = "Khám Theo Hợp Đồng"

    # Quy tắc mới của Jo: 'sdt' phải đủ đúng 10 chữ số — không đúng 10 số (kể cả để trống, thiếu/
    # thừa số, lẫn ký tự khác) thì tự điền giá trị mặc định (đúng nguyên văn theo yêu cầu của Jo).
    if not SDT_10_RE.match(val("sdt")):
        fixes["sdt"] = SDT_DEFAULT_VALUE

    return fixes


# ============================================================
# HÀM TIỆN ÍCH
# ============================================================

def clean_ws(value):
    """Dọn khoảng trắng ẩn (NBSP, zero-width...) ở bất kỳ vị trí nào, gộp khoảng trắng thừa."""
    if value is None:
        return ""
    s = str(value)
    for ch in NBSP_CHARS:
        s = s.replace(ch, " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def has_hidden_ws(value):
    if value is None:
        return False
    s = str(value)
    return any(ch in s for ch in NBSP_CHARS)


def norm_key(value):
    return clean_ws(value).lower()


def is_blank(value):
    if value is None:
        return True
    if isinstance(value, str) and clean_ws(value) == "":
        return True
    return False


def parse_date_cell(value):
    """Trả về (date, chuỗi_lỗi). Chấp nhận datetime/date của Excel hoặc chuỗi dd/MM/yyyy (và vài biến thể)."""
    if value is None or (isinstance(value, str) and clean_ws(value) == ""):
        return None, None
    if isinstance(value, datetime):
        return value.date(), None
    if isinstance(value, date):
        return value, None
    s = clean_ws(value)
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date(), None
        except ValueError:
            continue
    return None, f"'{value}' không đúng định dạng ngày (cần dd/MM/yyyy)"


def decimal_format_issue(raw_value):
    """Medinet dùng dấu PHẨY cho phần thập phân (chỉ số cận lâm sàng/sinh tồn); dấu CHẤM chỉ hợp lệ
    khi là phân cách hàng nghìn của số nguyên. Trả về mô tả lỗi nếu chuỗi có vẻ dùng dấu chấm làm
    thập phân, ngược lại None. Ô đã là số thật trong Excel (không phải chuỗi) thì bỏ qua — không có
    khái niệm dấu phẩy/chấm ở tầng lưu trữ, chỉ chuỗi hiển thị mới có vấn đề này."""
    if not isinstance(raw_value, str):
        return None
    s = clean_ws(raw_value)
    if "." not in s or "," in s:
        return None
    if DECIMAL_DOT_THOUSANDS_RE.match(s):
        return None  # số nguyên có dấu chấm phân cách nghìn — chấp nhận
    return f"'{s}' dùng dấu chấm (.) — số thập phân trên Medinet phải dùng dấu phẩy (,), ví dụ '4,1' thay vì '4.1'"


def to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = clean_ws(value)
    if s == "":
        return None
    s2 = s.replace(".", "").replace(",", ".") if ("," in s and s.count(",") == 1) else s
    try:
        return float(s2)
    except ValueError:
        try:
            return float(s)
        except ValueError:
            return None


def _is_literal_zero(raw_value):
    """True nếu ô chứa đúng giá trị 0 (số 0 thật, hoặc chuỗi '0'/'0.0'/'0,0') — dùng cho quy tắc
    xoá trắng các ô ICD (*_chandoansobo_icd / *_chandoanxacdinh_icd) khi lỡ điền 0."""
    if raw_value is None:
        return False
    if isinstance(raw_value, (int, float)):
        return raw_value == 0
    s = clean_ws(raw_value)
    return s in ("0", "0.0", "0,0")


# ============================================================
# NHẬN DIỆN CẤU TRÚC FILE (dòng nhãn / dòng mã field / dòng bắt đầu dữ liệu)
# ============================================================

def detect_layout_rows(ws):
    """Tự nhận diện dòng 'mã field' (keyword) trong các dòng đầu của sheet, không phụ thuộc cố định
    vào dòng 4 — theo yêu cầu Jo: file upload chỉ cần CÓ dòng keyword giống file mẫu, bất kể nằm ở
    dòng số mấy, dữ liệu vẫn được mapping đúng theo keyword.

    Heuristic: dòng mã field là dòng có nhiều ô dạng snake_case (chữ thường/số/gạch dưới) nhất, và
    phải chứa ít nhất 2 trong số các mã 'mỏ neo' quen thuộc (ho_ten, dinh_danh_ca_nhan, ngay_kham,
    gioi_tinh) để tránh nhận nhầm một dòng dữ liệu thành dòng mã.

    Trả về (label_row, code_row, data_start_row, note) — note khác None nếu phải dùng mặc định
    (dòng 2/4/5) do không tự nhận diện được, hoặc để thông báo đã nhận diện ở dòng khác dòng 4."""
    best_row, best_score = None, 0
    for r in range(1, min(LAYOUT_SEARCH_ROWS, ws.max_row) + 1):
        codes_in_row = set()
        score = 0
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if not v:
                continue
            s = str(v).strip()
            if CODE_PATTERN_RE.match(s):
                score += 1
                codes_in_row.add(s)
        anchor_hits = len(CODE_ROW_ANCHOR_FIELDS & codes_in_row)
        if anchor_hits >= 2 and score > best_score:
            best_score = score
            best_row = r

    if best_row is None:
        return LABEL_ROW, CODE_ROW, DATA_START_ROW, (
            f"Không tự nhận diện được dòng mã field (keyword) trong {LAYOUT_SEARCH_ROWS} dòng đầu — "
            f"dùng mặc định dòng {CODE_ROW}. Nếu file có cấu trúc khác, kết quả kiểm tra có thể không chính xác."
        )

    code_row = best_row
    label_row = code_row - 2 if code_row - 2 >= 1 else LABEL_ROW
    data_start_row = code_row + 1
    note = None
    if code_row != CODE_ROW:
        note = f"Đã tự nhận diện dòng mã field (keyword) ở dòng {code_row} (khác dòng {CODE_ROW} mặc định)."
    return label_row, code_row, data_start_row, note


def _build_code_col_map(ws, code_row):
    """Quét 1 lần dòng mã field, trả về dict {mã field: số thứ tự cột}."""
    m = {}
    for c in range(1, ws.max_column + 1):
        code = clean_ws(ws.cell(row=code_row, column=c).value)
        if code and code not in m:
            m[code] = c
    return m


# ============================================================
# ĐỌC CẤU TRÚC FILE + CÁC SHEET DANH MỤC
# ============================================================

def load_reference_sheets(wb):
    refs = {}
    for sheet_name, cfg in REF_SHEETS.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
        try:
            name_idx = headers.index(cfg["name_col"]) + 1
        except ValueError:
            continue
        id_idx = None
        if cfg.get("id_col") and cfg["id_col"] in headers:
            id_idx = headers.index(cfg["id_col"]) + 1
        parent_idx = None
        if cfg.get("parent_col") and cfg["parent_col"] in headers:
            parent_idx = headers.index(cfg["parent_col"]) + 1

        by_name, by_id = {}, {}
        for r in range(2, ws.max_row + 1):
            name_val = ws.cell(row=r, column=name_idx).value
            if name_val is None:
                continue
            id_val = ws.cell(row=r, column=id_idx).value if id_idx else None
            parent_val = ws.cell(row=r, column=parent_idx).value if parent_idx else None
            entry = {"name": name_val, "id": id_val, "parent": parent_val}
            by_name[norm_key(name_val)] = entry
            if id_val is not None:
                by_id[str(id_val).strip()] = entry
        refs[sheet_name] = {"by_name": by_name, "by_id": by_id}
    return refs


def read_column_defs(ws, label_row=LABEL_ROW, code_row=CODE_ROW):
    """Đọc dòng nhãn + dòng mã, trả về list dict {col, label, code, required}."""
    cols = []
    seen_codes = {}
    dup_codes = []
    for c in range(2, ws.max_column + 1):  # bỏ cột 1 (STT)
        label = ws.cell(row=label_row, column=c).value
        code = ws.cell(row=code_row, column=c).value
        label_s = clean_ws(label) if label else ""
        code_s = clean_ws(code) if code else ""
        if not label_s and not code_s:
            continue
        required = label_s.endswith("*")
        cols.append({
            "col": c,
            "col_letter": get_column_letter(c),
            "label": label_s.rstrip("*").strip() if label_s else "",
            "code": code_s,
            "required": required,
        })
        if code_s:
            if code_s in seen_codes:
                dup_codes.append((seen_codes[code_s], c, code_s))
            else:
                seen_codes[code_s] = c
    return cols, dup_codes


def find_last_data_row(ws, key_cols, data_start_row=DATA_START_ROW):
    """Tìm dòng cuối cùng có dữ liệu ở cột Họ tên / CCCD, để không quét hàng nghìn dòng trống."""
    last = data_start_row - 1
    for r in range(data_start_row, ws.max_row + 1):
        if any(not is_blank(ws.cell(row=r, column=c).value) for c in key_cols):
            last = r
    return last


# ============================================================
# KIỂM TRA TỪNG Ô
# ============================================================

def check_cell(code, raw_value, col_defs_by_code, refs, row_ctx):
    """Trả về list các issue: mỗi issue là dict {level, message}."""
    issues = []
    blank = is_blank(raw_value)

    col_def = col_defs_by_code.get(code)
    required = col_def["required"] if col_def else False

    # Sản khoa/Phụ khoa đã chọn 'Từ chối khám' (=1) cho đúng khối này → theo yêu cầu mới của Jo,
    # KHÔNG cảnh báo/báo lỗi gì cho ô '*_phanloai' của khối đó nữa (kể cả để trống hay ngoài khoảng
    # 1-5) — hệ thống sẽ tự chuyển null trong file tải về. Xem chỗ tính '_skip_phanloai_range_codes'
    # trong validate_workbook().
    skip_phanloai_range = code in (row_ctx.get("_skip_phanloai_range_codes") or ())

    if blank:
        if code in BINARY_STRICT_FIELDS:
            # Yêu cầu mới của Jo: 20 câu tiền sử bệnh AB-AU (+ benh_5nam) BẮT BUỘC phải có dữ liệu
            # (0 hoặc 1) — để trống thì Cảnh báo (không chặn Lưu) và hệ thống tự điền =0 trong file
            # tải về (xem BINARY_STRICT_FIELDS trong compute_data_fixes_for_row).
            issues.append({
                "level": "Cảnh báo",
                "message": "Đang để trống — cần có dữ liệu 0 (Không)/1 (Có); hệ thống sẽ tự điền = 0 trong file tải về",
            })
            return issues
        if code == "sdt":
            # Quy tắc mới của Jo: sdt phải đủ 10 số — để trống cũng tính là "không đúng 10 số",
            # hệ thống sẽ tự điền giá trị mặc định trong file tải về (xem compute_data_fixes_for_row).
            issues.append({
                "level": "Cảnh báo",
                "message": f"Đang để trống — cần đủ 10 số; hệ thống sẽ tự điền mặc định '{SDT_DEFAULT_VALUE}' trong file tải về",
            })
            return issues
        if skip_phanloai_range:
            return issues
        if required:
            issues.append({"level": "Lỗi", "message": "Bắt buộc nhập nhưng đang để trống"})
        return issues

    # ---- cận lâm sàng (xét nghiệm máu / sinh hóa máu / xét nghiệm nước tiểu) ----
    # Có nhập thì CHỈ kiểm tra lỗi dấu chấm (.) thay cho dấu phẩy (,) ở phần thập phân;
    # không kiểm tra/cảnh báo gì khác cho các ô này (kể cả khoảng trắng ẩn hay "không phải số hợp lệ").
    if code in CANLAMSANG_NUMERIC_FIELDS:
        fmt_err = decimal_format_issue(raw_value)
        if fmt_err:
            issues.append({"level": "Lỗi", "message": fmt_err})
        return issues

    # khoảng trắng ẩn — cảnh báo chung cho mọi ô có chữ
    if isinstance(raw_value, str) and has_hidden_ws(raw_value):
        issues.append({"level": "Cảnh báo", "message": "Chuỗi dính khoảng trắng ẩn (NBSP/zero-width) — nên dọn lại trước khi dùng"})

    text = clean_ws(raw_value)

    # ---- ngày tháng ----
    if code in DATE_FIELDS:
        d, err = parse_date_cell(raw_value)
        if err:
            issues.append({"level": "Lỗi", "message": err})
        else:
            row_ctx[code] = d
        return issues

    # ---- CCCD ----
    if code == "dinh_danh_ca_nhan":
        if not CCCD_RE.match(text):
            issues.append({"level": "Lỗi", "message": f"'{text}' không phải CCCD hợp lệ (cần đúng 12 chữ số)"})
        return issues

    # ---- SĐT ----
    # Quy tắc mới của Jo: phải đủ ĐÚNG 10 chữ số — không đúng 10 số (thiếu/thừa số, lẫn ký tự khác)
    # thì Cảnh báo (không chặn Lưu) và hệ thống tự điền giá trị mặc định trong file tải về.
    if code == "sdt":
        if not SDT_10_RE.match(text):
            issues.append({
                "level": "Cảnh báo",
                "message": f"'{text}' phải đủ 10 chữ số — hệ thống sẽ tự điền mặc định '{SDT_DEFAULT_VALUE}' trong file tải về",
            })
        return issues

    # ---- chọn 1 trong danh sách cố định ----
    if code in CHOICE_ONE_OF:
        allowed = CHOICE_ONE_OF[code]
        if text not in allowed:
            issues.append({"level": "Lỗi", "message": f"Giá trị '{text}' không thuộc {allowed}"})
        return issues

    if code in BINARY_01_FIELDS:
        if text not in ("0", "1"):
            issues.append({"level": "Lỗi", "message": f"Giá trị '{text}' phải là 0 (Không) hoặc 1 (Có)"})
        return issues

    if code.endswith("_chuaphathienbatthuong"):
        if text not in ("0", "1"):
            issues.append({"level": "Lỗi", "message": f"Giá trị '{text}' phải là 0 hoặc 1"})
        return issues

    if code.endswith("_phanloai"):
        if skip_phanloai_range:
            # Đã chọn 'Từ chối khám' cho đúng khối Sản khoa/Phụ khoa này (yêu cầu mới của Jo) —
            # không cảnh báo lỗi phân loại từ 1-5 nữa, hệ thống sẽ tự chuyển null trong file tải về.
            return issues
        if text not in ("1", "2", "3", "4", "5"):
            issues.append({"level": "Lỗi", "message": f"Giá trị '{text}' phải từ 1 đến 5"})
        return issues

    # ---- ICD ----
    # mat_chandoansobo_icd/mat_chandoanxacdinh_icd đã đúng hậu tố "_icd" ngay từ đầu (tên thật
    # trong file), nên chỉ cần điều kiện chung, không cần liệt kê riêng.
    if code.endswith("_icd"):
        for part in [p.strip() for p in text.split(",") if p.strip()]:
            if not ICD_RE.match(part):
                issues.append({"level": "Cảnh báo", "message": f"Mã '{part}' không giống định dạng ICD-10 thông thường"})
        return issues

    # ---- số (không thuộc cận lâm sàng) ----
    if code in NUMERIC_FIELDS:
        fmt_err = decimal_format_issue(raw_value)
        if fmt_err:
            issues.append({"level": "Lỗi", "message": fmt_err})
        else:
            n = to_number(raw_value)
            if n is None:
                issues.append({"level": "Lỗi", "message": f"'{text}' không phải là số hợp lệ"})
            elif code in VITAL_SIGN_RANGES:
                lo, hi = VITAL_SIGN_RANGES[code]
                if not (lo <= n <= hi):
                    issues.append({"level": "Cảnh báo", "message": f"Giá trị '{text}' nằm ngoài khoảng cho phép ({lo}-{hi})"})
        return issues

    # ---- Kết luận: chọn đúng 1 trong 5 giá trị ----
    if code == "danh_muc_de_nghi":
        if not any(norm_key(choice) == norm_key(text) for choice in DANH_MUC_DE_NGHI_CHOICES):
            issues.append({
                "level": "Lỗi",
                "message": f"'{text}' không thuộc 5 giá trị hợp lệ: {' | '.join(DANH_MUC_DE_NGHI_CHOICES)}",
            })
        return issues

    # ---- danh mục ID (nhiều mã cách nhau dấu phẩy) ----
    if code in STRICT_ID_CATEGORY:
        sheet_name, id_col, multi = STRICT_ID_CATEGORY[code]
        ref = refs.get(sheet_name)
        if ref:
            parts = [p.strip() for p in text.split(",") if p.strip()] if multi else [text]
            for p in parts:
                if p not in ref["by_id"]:
                    issues.append({"level": "Lỗi", "message": f"Mã '{p}' không có trong danh mục sheet '{sheet_name}'"})
        return issues

    # ---- danh mục theo TÊN (nghề nghiệp, nơi công tác) ----
    if code in STRICT_TEXT_CATEGORY:
        # Khi Đối tượng khám = 1 hoặc 2 (Sinh viên/học viên hoặc Người lao động chính thức),
        # noi_cong_tac chỉ cần CÓ dữ liệu, KHÔNG đối chiếu danh mục NoiLamViec nữa (yêu cầu của Jo).
        # Phần bắt buộc-không-trống được kiểm ở check_doi_tuong_rules().
        if code == "noi_cong_tac" and row_ctx.get("_skip_noicongtac_catalog"):
            return issues
        sheet_name, _ = STRICT_TEXT_CATEGORY[code]
        ref = refs.get(sheet_name)
        if ref:
            key = norm_key(text)
            if key not in ref["by_name"]:
                issues.append({"level": "Lỗi", "message": f"'{text}' không khớp đúng tên nào trong danh mục sheet '{sheet_name}' — cần copy nguyên văn từ sheet đó"})
            elif has_hidden_ws(raw_value):
                pass  # đã cảnh báo khoảng trắng ẩn ở trên
        return issues

    # ---- Tỉnh / Phường xã ----
    if code == "city_id":
        row_ctx[code] = text
        ref = refs.get("Tinh")
        if ref:
            key = norm_key(text)
            if key not in ref["by_name"] and text not in ref["by_id"]:
                issues.append({"level": "Lỗi", "message": f"'{text}' không khớp Tỉnh/Thành nào trong sheet 'Tinh'"})
        return issues

    if code == "ward_id":
        ref = refs.get("PhuongXa")
        if ref:
            key = norm_key(text)
            entry = ref["by_name"].get(key) or ref["by_id"].get(text)
            if entry is None:
                issues.append({"level": "Lỗi", "message": f"'{text}' không khớp Phường/Xã nào trong sheet 'PhuongXa'"})
            else:
                city_val = row_ctx.get("city_id")
                city_ref = refs.get("Tinh")
                if city_val is not None and city_ref:
                    city_entry = city_ref["by_name"].get(norm_key(city_val)) or city_ref["by_id"].get(clean_ws(city_val))
                    if city_entry and entry.get("parent") is not None and str(entry["parent"]) != str(city_entry["id"]):
                        issues.append({"level": "Cảnh báo", "message": f"Phường/Xã '{text}' có vẻ không thuộc Tỉnh/Thành đã nhập ở cột 'city_id'"})
        return issues

    # mặc định: không có quy tắc riêng — chỉ đã kiểm tra bắt buộc + khoảng trắng ẩn ở trên
    return issues


# ============================================================
# KIỂM TRA CHÉO GIỮA CÁC Ô (khối chuyên khoa + loại trừ theo giới tính)
# ============================================================

def check_specialty_blocks(raw_by_code, gioi_tinh_text):
    """Áp dụng quy tắc nghiệp vụ Jo bổ sung:
    - 1 khối chuyên khoa (4 ô): chuaphathienbatthuong=1 thì KHÔNG được có ICD, và phanloai phải =1;
      có ICD (sơ bộ hoặc xác định) thì phanloai phải >=2; phanloai KHÔNG được để trống.
    - Sản khoa/Phụ khoa (5 ô, có thêm 'từ chối khám' ở đầu): nếu từ chối khám=1 thì phanloai không
      bắt buộc; nếu không từ chối khám thì áp dụng đúng quy tắc 4 ô ở trên.
    - Nếu gioi_tinh=1 (Nam): các ô thai sản + toàn bộ khối Sản khoa/Phụ khoa phải để trống.
    Trả về list issue, mỗi issue có thêm khoá 'code' để gắn đúng vào cột/nhãn khi báo cáo.
    """
    issues = []

    def blank(code):
        return is_blank(raw_by_code.get(code))

    def val(code):
        return clean_ws(raw_by_code.get(code))

    def four_field_rules(check_c, sobo_c, xacdinh_c, phanloai_c, allow_blank_phanloai, mismatch_note=None):
        out = []
        has_check1 = (not blank(check_c)) and val(check_c) == "1"
        has_sobo = not blank(sobo_c)
        has_xacdinh = not blank(xacdinh_c)
        phanloai_blank = blank(phanloai_c)

        if has_check1 and (has_sobo or has_xacdinh):
            bad_code = sobo_c if has_sobo else xacdinh_c
            out.append((bad_code, "Đã chọn 'Chưa phát hiện bất thường' (=1) nên không được điền chẩn đoán ICD ở ô này — mâu thuẫn dữ liệu"))

        if phanloai_blank:
            if not allow_blank_phanloai:
                out.append((phanloai_c, "Không được để trống — phải chọn phân loại (1-5) tương ứng"))
        else:
            p = val(phanloai_c)
            if has_check1 and p != "1":
                out.append((phanloai_c, f"Đã chọn 'Chưa phát hiện bất thường' nên phải là Loại 1, đang là '{p}'"))
            elif (has_sobo or has_xacdinh) and p not in ("2", "3", "4", "5"):
                msg = f"Đã có chẩn đoán ICD (sơ bộ/xác định) nên phải chọn từ Loại 2 trở lên, đang là '{p}'"
                if mismatch_note:
                    # Yêu cầu mới của Jo: có ICD (sơ bộ/xác định) nhưng phân loại vẫn = 1 → thêm câu
                    # nhắc ngắn gọn đúng nguyên văn Jo yêu cầu, bên cạnh chi tiết đã có sẵn.
                    msg = f"{mismatch_note} — {msg}"
                out.append((phanloai_c, msg))
        return out

    for check_c, sobo_c, xacdinh_c, phanloai_c in SPECIALTY_BLOCKS_4FIELD.values():
        for code, msg in four_field_rules(check_c, sobo_c, xacdinh_c, phanloai_c, allow_blank_phanloai=False):
            issues.append({"code": code, "level": "Lỗi", "message": msg})

    is_male = (not blank("gioi_tinh")) and gioi_tinh_text == "1"

    # Câu nhắc riêng theo đúng nguyên văn Jo yêu cầu cho từng khối, khi có ICD (sơ bộ/xác định)
    # nhưng phân loại vẫn đang = 1 (mâu thuẫn dữ liệu).
    PHANLOAI_MISMATCH_NOTE = {"sankhoa": "xem lại Phân loại", "phukhoa": "xem lại phân loại"}

    for block_name, (tuchoi_c, check_c, sobo_c, xacdinh_c, phanloai_c) in SPECIALTY_BLOCKS_5FIELD.items():
        if is_male:
            continue  # nam giới: xử lý loại trừ riêng ở dưới, không áp quy tắc 4-ô nữa
        tuchoi1 = (not blank(tuchoi_c)) and val(tuchoi_c) == "1"
        if tuchoi1:
            for c in (check_c, sobo_c, xacdinh_c, phanloai_c):
                if not blank(c):
                    issues.append({"code": c, "level": "Cảnh báo",
                                   "message": "Đã chọn 'Từ chối khám' nhưng ô này vẫn có giá trị — kiểm tra lại"})
        else:
            note = PHANLOAI_MISMATCH_NOTE.get(block_name)
            for code, msg in four_field_rules(check_c, sobo_c, xacdinh_c, phanloai_c,
                                               allow_blank_phanloai=False, mismatch_note=note):
                issues.append({"code": code, "level": "Lỗi", "message": msg})

    if is_male:
        for code in MALE_EXCLUDED_FIELDS:
            if not blank(code):
                issues.append({"code": code, "level": "Lỗi",
                               "message": "Giới tính Nam nhưng ô này chỉ áp dụng cho nữ — cần bỏ trống hoặc kiểm tra lại giới tính"})

    return issues


def check_doi_tuong_rules(raw_by_code):
    """Quy tắc theo Đối tượng khám (Jo bổ sung):
    Nếu doi_tuong_kham = 1 (Sinh viên, học viên) HOẶC = 2 (Người lao động chính thức theo pháp luật
    ATVSLĐ) thì 3 ô bắt buộc không được để trống: nghenghiep_code, noi_cong_tac, noi_cong_tac_xa_phuong.
    (doi_tuong_kham có thể chứa nhiều mã cách nhau dấu phẩy — chỉ cần có mã '1' hoặc '2' trong đó.)
    """
    issues = []
    dt_codes = [p.strip() for p in clean_ws(raw_by_code.get("doi_tuong_kham")).split(",") if p.strip()]
    if "1" in dt_codes or "2" in dt_codes:
        for code in ("nghenghiep_code", "noi_cong_tac", "noi_cong_tac_xa_phuong"):
            if is_blank(raw_by_code.get(code)):
                issues.append({
                    "code": code,
                    "level": "Lỗi",
                    "message": "Bắt buộc nhập khi Đối tượng khám là 'Sinh viên, học viên' (mã 1) hoặc 'Người lao động chính thức' (mã 2)",
                })
    return issues


# check_cccd_consistency đã được THAY THẾ bằng check_gioitinh_ngaysinh_rules() ở trên (theo file
# quy tắc Jo gửi 24/09/2026) — giờ đối chiếu CCCD với giới tính/năm sinh không còn chặn Lưu (Lỗi)
# nữa mà TỰ SỬA trong file tải về + báo Cảnh báo, và có thêm ưu tiên suy giới tính từ họ tên
# ("Thị"/"Văn") trước khi đối chiếu CCCD. Xem compute_gioitinh_ngaysinh_fixes() ở trên.


def check_eye_pairs(raw_by_code):
    """3 cặp đo thị lực (quy tắc Jo bổ sung):
    - Điền theo từng CẶP: trong 1 cặp, mắt phải + mắt trái phải cùng có dữ liệu (hoặc cùng để trống);
      lệch một ô → báo lỗi.
    - Cặp 1 ('không kính') LOẠI TRỪ với cặp 2 ('kính lỗ'): đã điền cặp 1 thì không được điền cặp 2,
      và ngược lại.
    - ĐÃ BỎ loại trừ giữa cặp 1 ('không kính') và cặp 3 ('có kính') theo yêu cầu mới của Jo — 2 cặp
      này giờ có thể điền CÙNG NHAU bình thường (trước đây bị chặn nhầm là "Lỗi").
    """
    issues = []

    def blank(c):
        return is_blank(raw_by_code.get(c))

    pair_filled = {}
    for name, (a, b) in EYE_PAIRS:
        ba, bb = blank(a), blank(b)
        if ba != bb:
            empty_code = a if ba else b
            other_code = b if ba else a
            issues.append({
                "code": empty_code,
                "level": "Lỗi",
                "message": f"Phải điền cả cặp — ô mắt kia ('{other_code}') đã có dữ liệu nhưng ô này để trống",
            })
        pair_filled[name] = (not ba) or (not bb)

    if pair_filled.get("khongkinh") and pair_filled.get("kinhlo"):
        pairs = dict(EYE_PAIRS)
        for c in pairs["kinhlo"]:
            if not blank(c):
                issues.append({
                    "code": c,
                    "level": "Lỗi",
                    "message": f"Đã điền cặp '{EYE_PAIR_LABEL['khongkinh']}' thì không điền cặp '{EYE_PAIR_LABEL['kinhlo']}' (2 cặp này loại trừ nhau)",
                })
    return issues


# ============================================================
# KIỂM TRA TOÀN BỘ FILE
# ============================================================

def validate_workbook(file_bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    if SHEET_MAIN not in wb.sheetnames:
        raise ValueError(f"Không tìm thấy sheet '{SHEET_MAIN}' trong file — đây có phải đúng file Mẫu 03 không?")

    ws = wb[SHEET_MAIN]
    label_row, code_row, data_start_row, layout_note = detect_layout_rows(ws)
    col_defs, dup_codes = read_column_defs(ws, label_row, code_row)
    col_defs_by_code = {c["code"]: c for c in col_defs if c["code"]}
    refs = load_reference_sheets(wb)

    key_codes = ["ho_ten", "dinh_danh_ca_nhan"]
    key_cols = [c["col"] for c in col_defs if c["code"] in key_codes]
    last_row = find_last_data_row(ws, key_cols, data_start_row) if key_cols else ws.max_row

    ho_ten_col = next((c["col"] for c in col_defs if c["code"] == "ho_ten"), None)
    cccd_col = next((c["col"] for c in col_defs if c["code"] == "dinh_danh_ca_nhan"), None)

    issues = []
    cccd_seen = {}
    theluc_by_row = {}          # dòng Excel -> loại thể lực đã tính (1-5)
    danhmucdenghi_by_row = {}   # dòng Excel -> giá trị đề xuất cho ô Kết luận (khi đang trống)
    denghi_by_row = {}          # dòng Excel -> giá trị đề xuất cho ô 'de_nghi' (khi đang trống)
    data_fixes_by_row = {}      # dòng Excel -> {mã field: giá trị mới} theo các quy tắc tự sửa dữ liệu

    for r in range(data_start_row, last_row + 1):
        if all(is_blank(ws.cell(row=r, column=c["col"]).value) for c in col_defs):
            continue  # dòng trống hoàn toàn giữa các dòng có dữ liệu — bỏ qua

        ho_ten = clean_ws(ws.cell(row=r, column=ho_ten_col).value) if ho_ten_col else ""
        cccd = clean_ws(ws.cell(row=r, column=cccd_col).value) if cccd_col else ""

        row_ctx = {}
        raw_by_code = {}
        # Đọc trước Đối tượng khám để check_cell biết có phải mã 1/2 không (ảnh hưởng cách xử lý noi_cong_tac)
        dt_def = col_defs_by_code.get("doi_tuong_kham")
        if dt_def:
            dt_codes = [p.strip() for p in clean_ws(ws.cell(row=r, column=dt_def["col"]).value).split(",") if p.strip()]
            row_ctx["_skip_noicongtac_catalog"] = ("1" in dt_codes or "2" in dt_codes)

        # Đọc trước sankhoa_tuchoikham/phukhoa_tuchoikham (yêu cầu mới của Jo): khối nào đã chọn
        # 'Từ chối khám' (=1) thì check_cell() bỏ qua hẳn kiểm tra '_phanloai' (1-5) của đúng khối
        # đó — xem 'skip_phanloai_range' trong check_cell().
        skip_phanloai_codes = set()
        for tuchoi_code, phanloai_code in (("sankhoa_tuchoikham", "sankhoa_phanloai"),
                                            ("phukhoa_tuchoikham", "phukhoa_phanloai")):
            tuchoi_def = col_defs_by_code.get(tuchoi_code)
            if tuchoi_def:
                tv = clean_ws(ws.cell(row=r, column=tuchoi_def["col"]).value)
                if tv == "1":
                    skip_phanloai_codes.add(phanloai_code)
        row_ctx["_skip_phanloai_range_codes"] = skip_phanloai_codes

        for cdef in col_defs:
            code = cdef["code"]
            if not code:
                continue
            raw_value = ws.cell(row=r, column=cdef["col"]).value
            raw_by_code[code] = raw_value
            cell_issues = check_cell(code, raw_value, col_defs_by_code, refs, row_ctx)
            for iss in cell_issues:
                level = iss["level"]
                if code in WARN_ONLY_FIELDS and level == "Lỗi":
                    level = "Cảnh báo"  # hạ mức, không chặn nhập liệu
                issues.append({
                    "Dòng Excel": r,
                    "Họ tên": ho_ten,
                    "CCCD": cccd,
                    "Cột": cdef["col_letter"],
                    "Nhãn": cdef["label"],
                    "Mã field": code,
                    "Giá trị": ws.cell(row=r, column=cdef["col"]).value,
                    "Mức độ": level,
                    "Chi tiết": iss["message"],
                })

        gioi_tinh_text = clean_ws(raw_by_code.get("gioi_tinh"))
        cross_issues = (check_specialty_blocks(raw_by_code, gioi_tinh_text)
                        + check_doi_tuong_rules(raw_by_code)
                        + check_gioitinh_ngaysinh_rules(raw_by_code)
                        + check_eye_pairs(raw_by_code)
                        + check_mat_rules(raw_by_code))
        for iss in cross_issues:
            cdef = col_defs_by_code.get(iss["code"])
            issues.append({
                "Dòng Excel": r,
                "Họ tên": ho_ten,
                "CCCD": cccd,
                "Cột": cdef["col_letter"] if cdef else "",
                "Nhãn": cdef["label"] if cdef else iss["code"],
                "Mã field": iss["code"],
                "Giá trị": raw_by_code.get(iss["code"]),
                "Mức độ": iss["level"],
                "Chi tiết": iss["message"],
            })

        if cccd:
            cccd_seen.setdefault(cccd, []).append(r)

        the_luc = compute_the_luc_for_row(raw_by_code)
        if the_luc is not None:
            theluc_by_row[r] = the_luc

        if is_blank(raw_by_code.get("danh_muc_de_nghi")):
            suggested_ketluan = compute_danh_muc_de_nghi_for_row(raw_by_code)
            if suggested_ketluan is not None:
                danhmucdenghi_by_row[r] = suggested_ketluan

        if is_blank(raw_by_code.get("de_nghi")):
            denghi_by_row[r] = DE_NGHI_DEFAULT_VALUE

        fixes = compute_data_fixes_for_row(raw_by_code)
        if fixes:
            data_fixes_by_row[r] = fixes

    for cccd, rows in cccd_seen.items():
        if len(rows) > 1:
            for r in rows:
                issues.append({
                    "Dòng Excel": r,
                    "Họ tên": clean_ws(ws.cell(row=r, column=ho_ten_col).value) if ho_ten_col else "",
                    "CCCD": cccd,
                    "Cột": cccd_col and get_column_letter(cccd_col),
                    "Nhãn": "Số CMND/CCCD",
                    "Mã field": "dinh_danh_ca_nhan",
                    "Giá trị": cccd,
                    "Mức độ": "Lỗi",
                    "Chi tiết": f"CCCD trùng với {len(rows)-1} dòng khác trong cùng file (dòng {', '.join(str(x) for x in rows if x != r)}) — Medinet sẽ từ chối lưu nếu CCCD đã tồn tại",
                })

    structural_notes = []
    if layout_note:
        structural_notes.append(layout_note)
    for col1, col2, code in dup_codes:
        if code == "mat_docau_mt":
            # Theo yêu cầu của Jo: bỏ qua cảnh báo này — đây là đặc điểm bình thường của file gốc,
            # không phải lỗi cần kiểm tra.
            continue
        structural_notes.append(
            f"Mã field '{code}' xuất hiện ở CẢ cột {get_column_letter(col1)} lẫn {get_column_letter(col2)} "
            f"— có thể là lỗi trong file gốc, cần kiểm tra kỹ trước khi dùng để tránh nhầm dữ liệu."
        )

    issues_df = pd.DataFrame(issues)
    n_rows_checked = max(0, last_row - data_start_row + 1)
    return (issues_df, structural_notes, n_rows_checked, col_defs, theluc_by_row,
            danhmucdenghi_by_row, denghi_by_row, data_fixes_by_row)


def _unlock_workbook(wb):
    """Tắt khoá bảo vệ (Protect Sheet) mang theo từ file mẫu gốc trên MỌI sheet — nếu không tắt,
    Excel sẽ tự ẩn/xám bớt nhiều nút ở Home và Filter khi mở file tải về, khiến không gõ sửa
    được dữ liệu. Việc này chỉ đổi thuộc tính bảo vệ hiển thị của Excel, KHÔNG đụng tới cấu trúc
    cột/mã field/dữ liệu nên không ảnh hưởng gì tới chuẩn import Medinet. Dùng chung cho mọi hàm
    xuất file Excel tải về."""
    for sh in wb.worksheets:
        sh.protection.sheet = False
    try:
        wb.security = None   # tắt luôn khoá cấu trúc workbook (ẩn/hiện/xoá sheet) nếu file mẫu có đặt
    except Exception:
        pass


def _mark_issue_cells(ws, issues_df):
    """Tô màu (đỏ = lỗi, vàng = cảnh báo) + ghi chú vào từng ô có vấn đề, theo đúng issues_df đã
    tính từ validate_workbook(). Gom các issue theo ô (row, col_letter) để 1 ô có thể có nhiều ghi
    chú. Dùng chung cho annotate_workbook() (bản có tự sửa dữ liệu) và annotate_workbook_raw()
    (bản giữ nguyên dữ liệu gốc) để không có 2 nơi định nghĩa khác nhau cách đánh dấu ô lỗi."""
    if issues_df is None or issues_df.empty:
        return
    by_cell = {}
    for _, row in issues_df.iterrows():
        col_letter = row.get("Cột")
        r = row.get("Dòng Excel")
        if not col_letter or r is None:
            continue
        key = (int(r), str(col_letter))
        by_cell.setdefault(key, {"msgs": [], "has_error": False})
        by_cell[key]["msgs"].append(f"[{row['Mức độ']}] {row['Chi tiết']}")
        if row["Mức độ"] == "Lỗi":
            by_cell[key]["has_error"] = True

    for (r, col_letter), info in by_cell.items():
        cell = ws[f"{col_letter}{r}"]
        cell.fill = FILL_ERROR if info["has_error"] else FILL_WARN
        text = "\n".join(info["msgs"])
        cm = Comment(text, "Công cụ kiểm tra")
        cm.width = 320
        cm.height = max(60, 18 * (len(info["msgs"]) + 1))
        cell.comment = cm


def annotate_workbook_raw(file_bytes, issues_df):
    """Tạo file Excel để tải về GIỮ NGUYÊN Y HỆT dữ liệu gốc đã tải lên — KHÔNG áp bất kỳ quy tắc
    'tự sửa dữ liệu' nào (không điền Phân Loại thể lực, không đề xuất Kết luận, không điền mặc
    định 'de_nghi', không áp các quy tắc tự sửa giới tính/tiền sử bệnh/ICD/loại khám...), KHÔNG
    canh giữa dữ liệu. CHỈ làm 2 việc: (1) tắt khoá bảo vệ sheet như annotate_workbook() để vẫn
    gõ sửa tay được; (2) tô màu + ghi chú vào các ô đang sai quy tắc (đỏ = lỗi, vàng = cảnh báo)
    để người dùng tự xem và tự sửa theo đúng dữ liệu gốc của mình.
    Trả về bytes của file .xlsx."""
    wb = openpyxl.load_workbook(BytesIO(file_bytes))   # giữ nguyên, KHÔNG data_only (để lưu lại được)
    _unlock_workbook(wb)
    ws = wb[SHEET_MAIN]
    _mark_issue_cells(ws, issues_df)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def annotate_workbook(file_bytes, issues_df, theluc_by_row, danhmucdenghi_by_row=None,
                       denghi_by_row=None, data_fixes_by_row=None):
    """Tạo file Excel để tải về:
    (1) điền loại thể lực đã tính vào cột 'phanloai' (Phân Loại thể lực);
    (2) điền đề xuất cho ô Kết luận ('danh_muc_de_nghi') và ô 'de_nghi' khi đang để trống;
    (3) áp các quy tắc tự sửa dữ liệu khác (giới tính, tiền sử bệnh 0/1, ICD=0, loại khám, hồng cầu...);
    (3b) ép cột 'ngay_kham'/'ngay_sinh' thành CHỮ (Text) đúng khuôn dd/mm/yyyy — Excel hay tự hiển
    thị ngày theo định dạng khác (vd yyyy-mm-dd) tuỳ locale máy dù giá trị bên trong vẫn là ngày
    hợp lệ, nên ép hẳn về chuỗi text để hiển thị nhất quán, đúng chuẩn Jo cần cho Tampermonkey;
    (4) canh giữa dữ liệu trong toàn bộ vùng dữ liệu;
    (5) tô màu + ghi chú vào từng ô lỗi/cảnh báo để người dùng biết chỗ cần sửa.
    File kết quả KHÔNG bị khoá/bảo vệ — vẫn filter, xoá, copy, paste bình thường.
    Trả về bytes của file .xlsx."""
    danhmucdenghi_by_row = danhmucdenghi_by_row or {}
    denghi_by_row = denghi_by_row or {}
    data_fixes_by_row = data_fixes_by_row or {}

    wb = openpyxl.load_workbook(BytesIO(file_bytes))   # giữ nguyên, KHÔNG data_only (để lưu lại được)
    _unlock_workbook(wb)
    ws = wb[SHEET_MAIN]

    label_row, code_row, data_start_row, _note = detect_layout_rows(ws)
    col_map = _build_code_col_map(ws, code_row)

    # (1) Tìm cột 'phanloai' (Phân Loại thể lực) để điền kết quả
    phanloai_col = col_map.get("phanloai")
    if phanloai_col:
        for r, val in theluc_by_row.items():
            ws.cell(row=r, column=phanloai_col).value = val

    # (2) Điền đề xuất cho ô Kết luận ('danh_muc_de_nghi') khi đang để trống
    danh_muc_de_nghi_col = col_map.get("danh_muc_de_nghi")
    if danh_muc_de_nghi_col:
        for r, val in danhmucdenghi_by_row.items():
            ws.cell(row=r, column=danh_muc_de_nghi_col).value = val

    # Điền giá trị mặc định cho ô 'de_nghi' (Đề nghị, ghi rõ) khi đang để trống
    de_nghi_col = col_map.get("de_nghi")
    if de_nghi_col:
        for r, val in denghi_by_row.items():
            ws.cell(row=r, column=de_nghi_col).value = val

    # (3) Các quy tắc tự sửa dữ liệu khác — giới tính (AY/AZ/CV-DE), tiền sử bệnh 0/1 (AA-AU),
    # ICD=0 -> null (BI-DB), loại khám (ED) mặc định 2, số lượng hồng cầu (EE) mặc định 0
    for r, fixes in data_fixes_by_row.items():
        for code, new_val in fixes.items():
            col = col_map.get(code)
            if col:
                ws.cell(row=r, column=col).value = new_val

    key_cols = [c for code, c in col_map.items() if code in ("ho_ten", "dinh_danh_ca_nhan")]
    last_row = find_last_data_row(ws, key_cols, data_start_row) if key_cols else ws.max_row

    # (3b) Ép 'ngay_kham'/'ngay_sinh' thành CHỮ (Text) đúng khuôn dd/mm/yyyy — chạy SAU bước (3) để
    # áp lại cả cho ngay_sinh vừa được tự sửa theo CCCD ở trên (giá trị lúc đó là date object, chưa
    # ép định dạng). Excel hay tự hiển thị ngày theo định dạng khác (vd yyyy-mm-dd) tuỳ locale máy
    # dù giá trị bên trong vẫn là ngày hợp lệ — ép hẳn về chuỗi text để hiển thị nhất quán. Ô nào
    # không đọc được thành ngày hợp lệ thì giữ nguyên, không đụng vào (đã có 'Lỗi' báo riêng ở
    # check_cell()).
    for date_code in DATE_FIELDS:  # ngay_kham, ngay_sinh
        date_col = col_map.get(date_code)
        if not date_col:
            continue
        for r in range(data_start_row, last_row + 1):
            cell = ws.cell(row=r, column=date_col)
            d, _err = parse_date_cell(cell.value)
            if d is not None:
                cell.value = d.strftime("%d/%m/%Y")
                cell.number_format = "@"  # Text — Excel không tự đổi hiển thị theo locale nữa

    # (4) Canh giữa dữ liệu trong toàn bộ vùng dữ liệu đã map được mã field
    last_col = max(col_map.values()) if col_map else ws.max_column
    center_align = Alignment(horizontal="center", vertical="center")
    for r in range(data_start_row, last_row + 1):
        for c in range(1, last_col + 1):
            ws.cell(row=r, column=c).alignment = center_align

    # (5) Tô màu + ghi chú vào từng ô lỗi/cảnh báo
    _mark_issue_cells(ws, issues_df)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
