"""
merger.py — Ghép nhiều file Excel Mẫu 03 (mỗi file 1 khóm/tổ/phường) thành 1 file duy nhất,
giữ nguyên cấu trúc mẫu Medinet, để chạy 1 lần bằng script Tampermonkey thay vì chạy tay từng
file nhỏ.

Cố tình TÁI DÙNG logic đọc cấu trúc file (dò dòng nhãn/dòng mã field/dòng dữ liệu, dọn khoảng
trắng ẩn, tìm dòng dữ liệu cuối...) từ validators.py — để không có 2 nơi định nghĩa khác nhau
"thế nào là 1 file Mẫu 03 hợp lệ". Đọc theo MÃ FIELD (không theo số cột), nên vẫn ghép đúng dù
1 file nguồn có cột lệch vị trí đôi chút so với file nền.
"""
import os
import re
from copy import copy
from io import BytesIO

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formula import Tokenizer
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange

from validators import (
    SHEET_MAIN,
    CODE_ROW_ANCHOR_FIELDS,
    detect_layout_rows,
    _build_code_col_map,
    clean_ws,
    is_blank,
    find_last_data_row,
    FILL_WARN,
    format_medinet_columns,
    is_type_annotation_row,
)

NEW_FIELDS = (
    ("ma_phieu", "ngay_kham", "Mã phiếu", None),
    ("tai_noithuong_tt", "loai_kham", "Kết quả khám thính lực", "Tai trái (Nói thường)"),
    ("tai_noithuong_tp", "loai_kham", None, "Tai phải (Nói Thường)"),
    ("tai_noitham_tt", "loai_kham", None, "Tai trái (Nói thầm)"),
    ("tai_noitham_tp", "loai_kham", None, "Tai phải (Nói thầm)"),
)


def _shift_ref(text, idx, context_sheet):
    """Dịch địa chỉ khi chèn cột; không dịch tham chiếu sheet danh mục hoặc chuỗi literal."""
    if "!" in text:
        prefix, addr = text.rsplit("!", 1)
        if prefix.strip("'").replace("''", "'") != SHEET_MAIN:
            return text
        lead = prefix + "!"
    else:
        if context_sheet != SHEET_MAIN:
            return text
        lead, addr = "", text
    if not re.fullmatch(r"\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}", addr):
        return text
    def repl(m):
        col = column_index_from_string(m.group(2))
        return m.group(1) + get_column_letter(col + (col >= idx))
    return lead + re.sub(r"(\$?)([A-Z]{1,3})", repl, addr)


def _shift_formula(formula, idx, context_sheet):
    if not formula or not isinstance(formula, str):
        return formula
    has_equal = formula.startswith("=")
    tokens = Tokenizer(formula if has_equal else "=" + formula).items
    result = "".join(_shift_ref(t.value, idx, context_sheet)
                     if t.type == "OPERAND" and t.subtype == "RANGE" else t.value for t in tokens)
    return ("=" if has_equal else "") + result


def _shift_range(ref, idx):
    r = CellRange(str(ref))
    if r.min_col >= idx:
        r.shift(col_shift=1)
    elif r.max_col >= idx:
        r.max_col += 1
    return str(r)


def _insert_template_column(wb, ws, idx):
    """openpyxl không tự dịch công thức/merge/validation khi insert_cols, nên cập nhật rõ ràng."""
    merged = [str(r) for r in ws.merged_cells.ranges]
    for ref in merged:
        ws.unmerge_cells(ref)
    dims = [(key, copy(dim)) for key, dim in ws.column_dimensions.items()]
    ws.insert_cols(idx)
    ws.column_dimensions.clear()
    for key, dim in dims:
        lo = dim.min or column_index_from_string(key)
        hi = dim.max or lo
        dim.min = lo + (lo >= idx)
        dim.max = hi + (hi >= idx)
        dim.index = get_column_letter(dim.min)
        ws.column_dimensions[dim.index] = dim
    for ref in merged:
        ws.merge_cells(_shift_range(ref, idx))
    for sheet in wb.worksheets:
        for row in sheet:
            for cell in row:
                if cell.data_type == "f":
                    cell.value = _shift_formula(cell.value, idx, sheet.title)
        for dv in sheet.data_validations.dataValidation:
            if sheet == ws:
                dv.sqref = MultiCellRange(" ".join(_shift_range(r, idx) for r in dv.sqref.ranges))
            dv.formula1 = _shift_formula(dv.formula1, idx, sheet.title)
            dv.formula2 = _shift_formula(dv.formula2, idx, sheet.title)
    for dn in wb.defined_names.values():
        dn.attr_text = _shift_formula(dn.attr_text, idx, None)
    if ws.auto_filter.ref:
        ws.auto_filter.ref = _shift_range(ws.auto_filter.ref, idx)
        for fc in ws.auto_filter.filterColumn:
            if fc.colId + CellRange(ws.auto_filter.ref).min_col >= idx:
                fc.colId += 1
    for table in ws.tables.values():
        # Không tự mở rộng Excel Table khi thay đổi schema: tránh tạo header/column metadata lệch.
        if CellRange(table.ref).min_col < idx <= CellRange(table.ref).max_col:
            raise ValueError("File nền có Excel Table tại vị trí cần thêm cột; hãy dùng mẫu Medinet thường.")
        table.ref = _shift_range(table.ref, idx)


def ensure_new_fields(wb, ws, code_row):
    added = []
    for i, (code, anchor, label, sublabel) in enumerate(NEW_FIELDS):
        cmap = _build_code_col_map(ws, code_row)
        if code in cmap:
            continue
        following = [x[0] for x in NEW_FIELDS[i + 1:] if x[1] == anchor and x[0] in cmap]
        anchor_code = following[0] if following else anchor
        if anchor_code not in cmap:
            raise ValueError(f"Không có keyword '{anchor_code}' để xác định vị trí cột '{code}'.")
        idx = cmap[anchor_code]
        _insert_template_column(wb, ws, idx)
        for r in range(1, ws.max_row + 1):
            ws.cell(r, idx)._style = copy(ws.cell(r, idx + 1)._style)
        ws.cell(code_row, idx).value = code
        if code_row >= 3:
            ws.cell(code_row - 2, idx).value = label
            ws.cell(code_row - 1, idx).value = sublabel
        if code == "ma_phieu":
            if code_row >= 4:
                ws.cell(code_row - 3, idx).value = (
                    "Chỉ nhập mã phiếu với trường hợp cập nhật\n"
                    "Nhập mã phiếu để xác định phiếu cần cập nhật.\n"
                    "Không nhập mã phiếu, lấy phiếu chưa xóa có ngày khám mới nhất để cập nhật"
                )
            if code_row >= 3:
                ws.merge_cells(start_row=code_row-2, end_row=code_row-1, start_column=idx, end_column=idx)
        added.append(code)
    cmap = _build_code_col_map(ws, code_row)
    hearing = [cmap[x[0]] for x in NEW_FIELDS[1:]]
    if added and code_row >= 3 and hearing == list(range(hearing[0], hearing[0] + 4)):
        for ref in list(ws.merged_cells.ranges):
            if ref.min_row == ref.max_row == code_row-2 and ref.min_col >= hearing[0] and ref.max_col <= hearing[-1]:
                ws.unmerge_cells(str(ref))
        ws.cell(code_row-2, hearing[0]).value = "Kết quả khám thính lực"
        ws.merge_cells(start_row=code_row-2, end_row=code_row-2, start_column=hearing[0], end_column=hearing[-1])
    return added

# Thiếu 1 trong các mã "mỏ neo" này thì coi file KHÔNG đúng cấu trúc Mẫu 03 (ví dụ lỡ tải nhầm
# file Mẫu 01/02/04) — dừng hẳn file đó, không đoán mò để tránh ghép sai dữ liệu.
REQUIRED_ANCHOR_FIELDS = CODE_ROW_ANCHOR_FIELDS  # {"ho_ten", "dinh_danh_ca_nhan", "ngay_kham", "gioi_tinh"}


def default_unit_label(filename):
    """Đoán tên đơn vị mặc định từ tên file, để người dùng sửa lại trên giao diện nếu cần."""
    return clean_ws(os.path.splitext(filename)[0])


def read_source_file(file_bytes, filename, unit_label):
    """Đọc 1 file nguồn, trả về dict mô tả cấu trúc + toàn bộ dòng bệnh nhân (map theo MÃ FIELD).
    Raise ValueError với thông báo rõ ràng nếu file không nhận diện được là đúng cấu trúc Mẫu 03."""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    if SHEET_MAIN not in wb.sheetnames:
        raise ValueError(f"Không có sheet '{SHEET_MAIN}' trong file — có thể đây không phải file Mẫu 03.")
    ws = wb[SHEET_MAIN]

    label_row, code_row, data_start_row, layout_note = detect_layout_rows(ws)
    col_map = _build_code_col_map(ws, code_row)

    missing_anchor = REQUIRED_ANCHOR_FIELDS - set(col_map)
    if missing_anchor:
        raise ValueError(
            f"Không nhận diện được các mã field bắt buộc {sorted(missing_anchor)} trong dòng mã "
            f"field (dòng {code_row}) — có thể file này không đúng cấu trúc Mẫu 03, hoặc là mẫu khác."
        )

    ho_ten_col = col_map.get("ho_ten")
    cccd_col = col_map.get("dinh_danh_ca_nhan")
    key_cols = [c for c in (ho_ten_col, cccd_col) if c]
    last_row = find_last_data_row(ws, key_cols, data_start_row)

    rows = []
    for r in range(data_start_row, last_row + 1):
        if is_type_annotation_row(ws, r, col_map):
            continue
        if all(is_blank(ws.cell(row=r, column=c).value) for c in key_cols):
            continue  # dòng trống xen giữa — bỏ qua, không tính là 1 bệnh nhân
        values = {code: ws.cell(row=r, column=c).value for code, c in col_map.items()}
        rows.append({
            "ho_ten": clean_ws(values.get("ho_ten")),
            "cccd": clean_ws(values.get("dinh_danh_ca_nhan")),
            "values": values,
        })

    return {
        "filename": filename,
        "unit_label": unit_label,
        "file_bytes": file_bytes,
        "layout_note": layout_note,
        "code_row": code_row,
        "data_start_row": data_start_row,
        "col_map": col_map,
        "rows": rows,
    }


def merge_mau03_files(sources):
    """sources: list các dict trả về từ read_source_file(), THEO ĐÚNG THỨ TỰ đã tải lên.
    File ĐẦU TIÊN trong danh sách được dùng làm NỀN — giữ nguyên định dạng, các sheet danh mục
    (NgheNghiep, NoiLamViec, DoiTuongKham...) và 4 dòng đầu; các sheet danh mục của các file khác
    KHÔNG được dùng tới, vì đó là danh mục dùng chung của Medinet chứ không phải dữ liệu riêng của
    từng đơn vị.

    Trả về (merged_bytes, summary):
      summary = {
        "total_rows": tổng số bệnh nhân đã ghép,
        "per_file": [(filename, unit_label, số dòng lấy được), ...],
        "duplicates": [{"cccd": ..., "rows": [(dòng Excel sau ghép, đơn vị, họ tên), ...]}, ...],
        "field_gaps": [(filename, (mã field thiếu so với file nền, ...)), ...],
        "layout_notes": [(filename, ghi chú cấu trúc), ...],
      }
    """
    if len(sources) < 2:
        raise ValueError("Cần ít nhất 2 file để ghép.")

    base = sources[0]
    base_col_map = base["col_map"]
    base_data_start_row = base["data_start_row"]

    wb = openpyxl.load_workbook(BytesIO(base["file_bytes"]))  # không data_only — giữ định dạng để lưu lại
    ws = wb[SHEET_MAIN]
    added_fields = ensure_new_fields(wb, ws, base["code_row"])
    base_col_map = _build_code_col_map(ws, base["code_row"])

    # Tắt khoá bảo vệ (Protect Sheet) mang theo từ file nền — cùng lỗi và cùng cách vá như
    # annotate_workbook() trong validators.py: không tắt thì Excel tự ẩn/xám bớt nút Home/Filter
    # ở file ghép tải về. Chỉ đổi thuộc tính bảo vệ hiển thị, không đụng cấu trúc/dữ liệu nên
    # không ảnh hưởng chuẩn import Medinet.
    for sh in wb.worksheets:
        sh.protection.sheet = False
    try:
        wb.security = None
    except Exception:
        pass

    # Xoá sạch vùng dữ liệu cũ (kể cả của chính file nền) — sẽ ghi lại TOÀN BỘ theo đúng thứ tự
    # ghép bên dưới, tránh sót dòng thừa nếu tổng số dòng ghép ít hơn số dòng gốc của file nền.
    old_last_row = find_last_data_row(
        ws, [c for c in (base_col_map.get("ho_ten"), base_col_map.get("dinh_danh_ca_nhan")) if c],
        base_data_start_row,
    )
    for r in range(base_data_start_row, old_last_row + 1):
        for c in base_col_map.values():
            ws.cell(row=r, column=c).value = None

    field_gaps = []
    all_rows = []  # (unit_label, filename, ho_ten, cccd, values)
    for src in sources:
        gaps = tuple(sorted(set(base_col_map) - set(src["col_map"])))
        if gaps:
            field_gaps.append((src["filename"], gaps))
        for row in src["rows"]:
            all_rows.append((src["unit_label"], src["filename"], row["ho_ten"], row["cccd"], row["values"]))

    # Ghi dữ liệu đã ghép vào sheet chính, đánh lại STT liên tục — đọc/ghi theo MÃ FIELD nên vẫn
    # đúng cột dù vị trí cột ở file nguồn có lệch đôi chút so với file nền.
    cccd_positions = {}
    for idx, (unit_label, filename, ho_ten, cccd, values) in enumerate(all_rows):
        r = base_data_start_row + idx
        ws.cell(row=r, column=1).value = idx + 1  # STT
        for code, col in base_col_map.items():
            cell = ws.cell(row=r, column=col)
            cell.value = values.get(code)
            if isinstance(cell.value, str):
                cell.data_type = "s"
        if cccd:
            cccd_positions.setdefault(cccd, []).append((r, unit_label, ho_ten))

    # Đánh dấu (KHÔNG chặn) các CCCD trùng — tô vàng + ghi chú, cùng kiểu FILL_WARN app kiểm tra
    # đang dùng, để nhất quán giao diện giữa 2 tính năng.
    cccd_col = base_col_map.get("dinh_danh_ca_nhan")
    duplicates = []
    if cccd_col:
        for cccd, occurrences in cccd_positions.items():
            if len(occurrences) > 1:
                duplicates.append({"cccd": cccd, "rows": occurrences})
                for r, unit_label, ho_ten in occurrences:
                    cell = ws.cell(row=r, column=cccd_col)
                    cell.fill = FILL_WARN
                    others = [f"{u} (dòng {rr})" for rr, u, _ in occurrences if rr != r]
                    cm = Comment(
                        "CCCD trùng với: " + "; ".join(others) + " — vẫn giữ lại cả hai, tự kiểm tra "
                        "xem có phải cùng 1 người bị nhập 2 lần hay trùng ngẫu nhiên trước khi chạy "
                        "script điền Medinet (CCCD trùng thật sẽ bị Medinet từ chối lưu).",
                        "Công cụ ghép file",
                    )
                    cm.width, cm.height = 320, 90
                    cell.comment = cm

    # Thêm 1 sheet phụ ghi rõ nguồn gốc từng dòng — không đụng vào cấu trúc sheet chính, chỉ để
    # tra cứu/đối chiếu sau này (ví dụ biết dòng lỗi nào cần báo lại cho đúng khóm/tổ).
    if "Nguon_Ghep" in wb.sheetnames:
        del wb["Nguon_Ghep"]
    src_ws = wb.create_sheet("Nguon_Ghep")
    src_ws.append(["STT (sau khi ghép)", "Họ tên", "CCCD", "Đơn vị nguồn", "Tên file gốc", "CCCD trùng?"])
    dup_cccds = {d["cccd"] for d in duplicates}
    for idx, (unit_label, filename, ho_ten, cccd, values) in enumerate(all_rows):
        src_ws.append([idx + 1, ho_ten, cccd, unit_label, filename, "Có" if cccd in dup_cccds else ""])

    format_medinet_columns(wb, ws, base["code_row"], base_data_start_row)
    format_medinet_columns(wb, src_ws, 1, 2)
    out = BytesIO()
    wb.save(out)

    summary = {
        "added_fields": added_fields,
        "total_rows": len(all_rows),
        "per_file": [(s["filename"], s["unit_label"], len(s["rows"])) for s in sources],
        "duplicates": duplicates,
        "field_gaps": field_gaps,
        "layout_notes": [(s["filename"], s["layout_note"]) for s in sources if s["layout_note"]],
    }
    return out.getvalue(), summary
