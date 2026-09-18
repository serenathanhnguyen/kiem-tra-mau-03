import re
from datetime import datetime, date, timedelta
from io import BytesIO

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter

# ============================================================
# CẤU HÌNH CẤU TRÚC FILE MẪU 03
# ============================================================
SHEET_MAIN = "ThongTinHanhChinh"
LABEL_ROW = 2          # dòng nhãn (có dấu * cho cột bắt buộc)
CODE_ROW = 4            # dòng mã field (name) — dùng để map, giống script Tampermonkey
DATA_START_ROW = 5      # dữ liệu bệnh nhân bắt đầu từ dòng này

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
BINARY_01_FIELDS = [
    "benh_5nam", "benh_than_kinh", "benh_mat", "benh_tai", "benh_tim", "pt_tim_mach",
    "tang_ha", "kho_tho", "benh_phoi", "benh_than", "nghien_ruou_bia", "dai_thao_duong",
    "benh_tam_than", "mat_y_thuc", "ngat_chong_mat", "benh_tieu_hoa", "roi_loan_giac_ngu",
    "tai_bien_mach_mau_nao", "cot_song", "su_dung_ruou_bia", "su_dung_ma_tuy",
    "dieu_tri_benh_co_khong", "thai_san_co_khong",
]

# Các trường số (kết quả cận lâm sàng, chỉ số sinh tồn...) — phải là số
NUMERIC_FIELDS = [
    "chieucao", "cannang", "nhiptho", "mach", "huyetaptamthu", "huyetaptamtruong",
    "kskdk_xnm_slhc", "kskdk_xnm_huyetsacto", "kskdk_xnm_hematocrit", "kskdk_xnm_mcv",
    "kskdk_xnm_mch", "kskdk_xnm_mchc", "kskdk_xnm_rdw", "kskdk_xnm_slbc",
    "kskdk_xnm_slbc_trungtinh", "kskdk_xnm_slbc_lympho", "kskdk_xnm_slbc_donnhan",
    "kskdk_xnm_slbc_aitoan", "kskdk_xnm_slbc_aikiem", "kskdk_xnm_sltc",
    "kskdk_shm_duongmau", "kskdk_shm_ure", "kskdk_shm_creatinin", "kskdk_shm_asat_got",
    "kskdk_shm_alat_gpt", "kskdk_xnnt_titrong", "kskdk_xnnt_ph", "kskdk_xnnt_bachcau",
    "kskdk_xnnt_hongcau", "kskdk_xnnt_protein", "kskdk_xnnt_glucose", "kskdk_xnnt_cetonic",
    "kskdk_xnnt_bilirubin", "kskdk_xnnt_urobilinogen",
    "mat_khongkinh_mp", "mat_khongkinh_mt", "mat_kinhlo_mp", "mat_kinhlo_mt",
    "mat_cokinh_mp", "mat_cokinh_mt", "mat_docau_mp", "mat_docau_mt",
    "mat_dotru_mp", "mat_dotru_mt", "mat_truc_mt",
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
NBSP_CHARS = ["\xa0", "\u200b", "\ufeff", "\u2007", "\u202f"]

# 5 giá trị hợp lệ của ô Kết luận (đúng theo danh sách Jo cung cấp)
DANH_MUC_DE_NGHI_CHOICES = [
    "Bình thường, hẹn khám định kỳ lần sau",
    "Có yếu tố nguy cơ, cần theo dõi thêm",
    "Đã có bệnh mạn tính, tiếp tục điều trị theo phác đồ/toa cũ",
    "Chuyển tuyến, khám chuyên khoa",
    "Khác",
]

# 13 khối chuyên khoa dạng 4 ô: _chuaphathienbatthuong / _chandoansobo_icd / _chandoanxacdinh_icd / _phanloai
# Mã lấy ĐÚNG theo file gốc — thankinh có lỗi chính tả sẵn trong file ("chuandoansobo" thay vì "chandoansobo")
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
    "mat": ("mat_chuaphathienbatthuong", "mat_chandoansobo_icd", "mat_chandoanxacdinh_icd", "mat_phanloai"),
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
EYE_PAIRS = [
    ("khongkinh", ("mat_khongkinh_mp", "mat_khongkinh_mt")),  # cặp 1: không kính
    ("kinhlo", ("mat_kinhlo_mp", "mat_kinhlo_mt")),           # cặp 2: kính lỗ
    ("cokinh", ("mat_cokinh_mp", "mat_cokinh_mt")),           # cặp 3: có kính
]
EYE_PAIR_LABEL = {"khongkinh": "không kính", "kinhlo": "kính lỗ", "cokinh": "có kính"}


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


def read_column_defs(ws):
    """Đọc dòng nhãn + dòng mã, trả về list dict {col, label, code, required}."""
    cols = []
    seen_codes = {}
    dup_codes = []
    for c in range(2, ws.max_column + 1):  # bỏ cột 1 (STT)
        label = ws.cell(row=LABEL_ROW, column=c).value
        code = ws.cell(row=CODE_ROW, column=c).value
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


def find_last_data_row(ws, key_cols):
    """Tìm dòng cuối cùng có dữ liệu ở cột Họ tên / CCCD, để không quét hàng nghìn dòng trống."""
    last = DATA_START_ROW - 1
    for r in range(DATA_START_ROW, ws.max_row + 1):
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

    if blank:
        if required:
            issues.append({"level": "Lỗi", "message": "Bắt buộc nhập nhưng đang để trống"})
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
    if code == "sdt":
        if not PHONE_RE.match(text):
            issues.append({"level": "Cảnh báo", "message": f"'{text}' có định dạng số điện thoại lạ (thường là 10 số, bắt đầu bằng 0)"})
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
        if text not in ("1", "2", "3", "4", "5"):
            issues.append({"level": "Lỗi", "message": f"Giá trị '{text}' phải từ 1 đến 5"})
        return issues

    # ---- ICD ----
    if code.endswith("_icd"):
        for part in [p.strip() for p in text.split(",") if p.strip()]:
            if not ICD_RE.match(part):
                issues.append({"level": "Cảnh báo", "message": f"Mã '{part}' không giống định dạng ICD-10 thông thường"})
        return issues

    # ---- số ----
    if code in NUMERIC_FIELDS:
        fmt_err = decimal_format_issue(raw_value)
        if fmt_err:
            issues.append({"level": "Lỗi", "message": fmt_err})
        else:
            n = to_number(raw_value)
            if n is None:
                issues.append({"level": "Lỗi", "message": f"'{text}' không phải là số hợp lệ"})
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

    def four_field_rules(check_c, sobo_c, xacdinh_c, phanloai_c, allow_blank_phanloai):
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
                out.append((phanloai_c, f"Đã có chẩn đoán ICD (sơ bộ/xác định) nên phải chọn từ Loại 2 trở lên, đang là '{p}'"))
        return out

    for check_c, sobo_c, xacdinh_c, phanloai_c in SPECIALTY_BLOCKS_4FIELD.values():
        for code, msg in four_field_rules(check_c, sobo_c, xacdinh_c, phanloai_c, allow_blank_phanloai=False):
            issues.append({"code": code, "level": "Lỗi", "message": msg})

    is_male = (not blank("gioi_tinh")) and gioi_tinh_text == "1"

    for tuchoi_c, check_c, sobo_c, xacdinh_c, phanloai_c in SPECIALTY_BLOCKS_5FIELD.values():
        if is_male:
            continue  # nam giới: xử lý loại trừ riêng ở dưới, không áp quy tắc 4-ô nữa
        tuchoi1 = (not blank(tuchoi_c)) and val(tuchoi_c) == "1"
        if tuchoi1:
            for c in (check_c, sobo_c, xacdinh_c, phanloai_c):
                if not blank(c):
                    issues.append({"code": c, "level": "Cảnh báo",
                                   "message": "Đã chọn 'Từ chối khám' nhưng ô này vẫn có giá trị — kiểm tra lại"})
        else:
            for code, msg in four_field_rules(check_c, sobo_c, xacdinh_c, phanloai_c, allow_blank_phanloai=False):
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


def check_cccd_consistency(raw_by_code):
    """Suy thông tin từ CCCD 12 số và đối chiếu với ô đã nhập (quy tắc Jo bổ sung):
    - Ký tự thứ 4: CHẴN (0,2,4,6,8) → Nam (gioi_tinh=1); LẺ (1,3,5,7,9) → Nữ (gioi_tinh=2).
    - Ký tự thứ 4 cũng cho biết thế kỷ (0-1:19xx, 2-3:20xx, 4-5:21xx...); ký tự 5-6 = 2 số cuối
      năm sinh. Ghép lại ra năm sinh đầy đủ, đối chiếu với năm của ô ngay_sinh.
    Ví dụ 079171301583: ký tự 4='1' (lẻ→Nữ, thế kỷ 19xx), ký tự 5-6='71' → năm sinh 1971.
    Chỉ chạy khi CCCD đủ 12 số (CCCD sai đã được báo ở chỗ khác).
    """
    issues = []
    cccd = clean_ws(raw_by_code.get("dinh_danh_ca_nhan"))
    if not CCCD_RE.match(cccd):
        return issues

    d4 = int(cccd[3])          # ký tự thứ 4
    yy = int(cccd[4:6])        # ký tự 5-6

    # 1) Giới tính
    gt = clean_ws(raw_by_code.get("gioi_tinh"))
    if gt in ("1", "2"):
        expected = "1" if d4 % 2 == 0 else "2"
        if gt != expected:
            ten = "Nam" if expected == "1" else "Nữ"
            issues.append({
                "code": "gioi_tinh",
                "level": "Lỗi",
                "message": f"Giới tính không khớp CCCD: ký tự thứ 4 ('{cccd[3]}') cho biết là {ten}, nhưng ô giới tính đang là '{gt}'",
            })

    # 2) Năm sinh
    century = 1900 + (d4 // 2) * 100   # 0,1→1900; 2,3→2000; 4,5→2100; ...
    year_cccd = century + yy
    d, err = parse_date_cell(raw_by_code.get("ngay_sinh"))
    if d is not None and d.year != year_cccd:
        issues.append({
            "code": "ngay_sinh",
            "level": "Lỗi",
            "message": f"Năm sinh không khớp CCCD: CCCD cho biết năm sinh {year_cccd} (ký tự 4-6 = '{cccd[3:6]}'), nhưng ngày sinh đang là năm {d.year}",
        })
    return issues


def check_eye_pairs(raw_by_code):
    """3 cặp đo thị lực (quy tắc Jo bổ sung):
    - Điền theo từng CẶP: trong 1 cặp, mắt phải + mắt trái phải cùng có dữ liệu (hoặc cùng để trống);
      lệch một ô → báo lỗi.
    - Cặp 1 ('không kính') LOẠI TRỪ với cặp 2 ('kính lỗ') và cặp 3 ('có kính'): nếu đã điền cặp 1
      thì không được điền cặp 2/3, và ngược lại. (Cặp 2 và 3 vẫn có thể điền cùng nhau.)
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

    if pair_filled.get("khongkinh"):
        pairs = dict(EYE_PAIRS)
        for name in ("kinhlo", "cokinh"):
            if pair_filled.get(name):
                for c in pairs[name]:
                    if not blank(c):
                        issues.append({
                            "code": c,
                            "level": "Lỗi",
                            "message": f"Đã điền cặp '{EYE_PAIR_LABEL['khongkinh']}' thì không điền cặp '{EYE_PAIR_LABEL[name]}' (cặp 1 loại trừ cặp 2 và 3)",
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
    col_defs, dup_codes = read_column_defs(ws)
    col_defs_by_code = {c["code"]: c for c in col_defs if c["code"]}
    refs = load_reference_sheets(wb)

    key_codes = ["ho_ten", "dinh_danh_ca_nhan"]
    key_cols = [c["col"] for c in col_defs if c["code"] in key_codes]
    last_row = find_last_data_row(ws, key_cols) if key_cols else ws.max_row

    ho_ten_col = next((c["col"] for c in col_defs if c["code"] == "ho_ten"), None)
    cccd_col = next((c["col"] for c in col_defs if c["code"] == "dinh_danh_ca_nhan"), None)

    issues = []
    cccd_seen = {}

    for r in range(DATA_START_ROW, last_row + 1):
        if all(is_blank(ws.cell(row=r, column=c["col"]).value) for c in col_defs):
            continue  # dòng trống hoàn toàn giữa các dòng có dữ liệu — bỏ qua

        ho_ten = clean_ws(ws.cell(row=r, column=ho_ten_col).value) if ho_ten_col else ""
        cccd = clean_ws(ws.cell(row=r, column=cccd_col).value) if cccd_col else ""

        row_ctx = {}
        raw_by_code = {}
        # Đọc trước Đối tượng khám để check_cell biết có phải mã 2 không (ảnh hưởng cách xử lý noi_cong_tac)
        dt_def = col_defs_by_code.get("doi_tuong_kham")
        if dt_def:
            dt_codes = [p.strip() for p in clean_ws(ws.cell(row=r, column=dt_def["col"]).value).split(",") if p.strip()]
            row_ctx["_skip_noicongtac_catalog"] = ("1" in dt_codes or "2" in dt_codes)
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
                        + check_cccd_consistency(raw_by_code)
                        + check_eye_pairs(raw_by_code))
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
    for col1, col2, code in dup_codes:
        structural_notes.append(
            f"Mã field '{code}' xuất hiện ở CẢ cột {get_column_letter(col1)} lẫn {get_column_letter(col2)} "
            f"— có thể là lỗi trong file gốc, cần kiểm tra kỹ trước khi dùng để tránh nhầm dữ liệu."
        )

    issues_df = pd.DataFrame(issues)
    n_rows_checked = max(0, last_row - DATA_START_ROW + 1)
    return issues_df, structural_notes, n_rows_checked, col_defs
