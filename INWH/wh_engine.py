"""
SAP IN WH Portable Engine (v1.0)
Xu ly loc trang thua va dong so thu tu don kho Warehouse Order tu PDF SAP
Dua tren bang doi chieu Excel (Mapping Table) hoac sinh tu dong.
"""

import sys
import os
import re
import csv
import argparse

try:
    import pymupdf as fitz
except ImportError:
    import fitz


def extract_wh_number(page_text: str) -> str:
    """
    Trich xuat ma Warehouse Order (10 chu so, thuong bat dau bang 60...) tu trang PDF SAP.
    """
    patterns = [
        r"(?:Warehouse\s*Order|WH\s*Order|WO|Số\s*WH|Order\s*No\.?)[\s:]*([0-9]{10})",
        r"\b(60[0-9]{8})\b",
        r"\b([0-9]{10})\b"
    ]
    for pattern in patterns:
        match = re.search(pattern, page_text, re.IGNORECASE)
        if match:
            return match.group(1)
    return "UNKNOWN_WH"


def has_material_items(page_text: str) -> bool:
    """
    Kiem tra trang co chua dong vat tu thuc te hay chi la trang thua (chu ky, ngay gio in).
    Trang chi chua header/footer ma khong co dong mat hang vat tu thi tra ve False.
    """
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]

    # Cac cum tu footer/header chu ky dac trung cua SAP can bo qua
    ignore_keywords = (
        "signature", "ký nhận", "người lập", "người giao", "người nhận",
        "printed by", "print date", "ngày in", "thời gian in", "page ", "trang ",
        "xác nhận", "warehouse order", "wh order"
    )

    # Nhan dien ma vat tu SAP:
    # 1. Alphanumeric 8-18 ky tu co ca chu va so (vi du: BIW32026240, V04010...)
    material_code_pattern = re.compile(r"\b(?=[A-Z0-9\-]{8,18}\b)(?=.*[A-Z])(?=.*[0-9])[A-Z0-9\-]+\b")
    # 2. Don vi tinh thong dung di kem so luong (PC, PCS, EA, SET, KG, M, L, BOX, ROL, CAR, UN...)
    uom_pattern = re.compile(r"\b(PC|PCS|EA|SET|KG|M|L|BOX|ROL|CAR|UN|THÙNG|BỘ|CHIẾC|CÁI)\b", re.IGNORECASE)

    for line in lines:
        low_line = line.lower()
        if any(kw in low_line for kw in ignore_keywords):
            continue
        # Bo qua ma WH 10 so thuan tuy
        if re.fullmatch(r"60[0-9]{8}", line) or re.fullmatch(r"[0-9]{10}", line):
            continue
        if material_code_pattern.search(line) or uom_pattern.search(line):
            return True

    return False


def load_mapping_file(mapping_csv_path: str):
    """
    Doc file mapping do Excel xuat ra:
    Mã WH -> Số đơn gán (vd: 6028720086 -> 06-01)
    """
    mapping_order = []
    mapping_dict = {}
    if not mapping_csv_path or not os.path.exists(mapping_csv_path):
        return mapping_order, mapping_dict

    # Ho tro ca UTF-8 BOM va UTF-8
    with open(mapping_csv_path, mode="r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row or len(row) < 2:
                continue
            wh = row[0].strip()
            label = row[1].strip()
            # Bo qua tieu de
            if wh and wh.lower() not in ("mã wh", "wh_no", "warehouse_order", "ma_wh"):
                if wh not in mapping_dict:
                    mapping_order.append(wh)
                mapping_dict[wh] = label

    return mapping_order, mapping_dict


def process_wh_pdf(input_pdf: str, mapping_csv: str, day_prefix: str, start_num: int, out_pdf: str, out_csv: str):
    if not os.path.exists(input_pdf):
        raise FileNotFoundError(f"Khong tim thay file PDF dau vao: {input_pdf}")

    doc = fitz.open(input_pdf)
    out_doc = fitz.open()

    mapping_order, mapping_dict = load_mapping_file(mapping_csv)
    use_mapping = len(mapping_dict) > 0

    # 1. Quet toan bo trang trong PDF va phan tich theo tung ma WH
    pdf_wh_data = {}  # wh_no -> {"pages": [idx...], "dropped": 0, "total": 0}
    pdf_wh_order = []
    last_wh = None

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        text = page.get_text("text")
        wh_no = extract_wh_number(text)

        if wh_no == "UNKNOWN_WH" and last_wh is not None:
            wh_no = last_wh
        last_wh = wh_no

        if wh_no not in pdf_wh_data:
            pdf_wh_data[wh_no] = {"pages": [], "dropped": 0, "total": 0}
            pdf_wh_order.append(wh_no)

        pdf_wh_data[wh_no]["total"] += 1
        if has_material_items(text):
            pdf_wh_data[wh_no]["pages"].append(page_idx)
        else:
            pdf_wh_data[wh_no]["dropped"] += 1

    # 2. Xac dinh thu tu WH cho file PDF dau ra
    final_wh_sequence = []
    if use_mapping:
        # Uu tien theo dung thu tu danh sach Excel
        for wh in mapping_order:
            if wh not in final_wh_sequence:
                final_wh_sequence.append(wh)
        # Neu trong PDF co WH ngoai danh sach Excel thi them vao cuoi
        for wh in pdf_wh_order:
            if wh not in final_wh_sequence:
                final_wh_sequence.append(wh)
    else:
        # Neu khong co mapping, giu nguyen thu tu trong PDF va tu sinh so theo day_prefix
        final_wh_sequence = list(pdf_wh_order)
        cur_seq = start_num
        for wh in final_wh_sequence:
            mapping_dict[wh] = f"{day_prefix}-{cur_seq:02d}"
            cur_seq += 1

    # 3. Loc trang, sap xep lai va dong so goc duoi ben phai
    csv_records = []

    for idx, wh_no in enumerate(final_wh_sequence, start=1):
        label = mapping_dict.get(wh_no, "")
        in_excel = wh_no in mapping_dict if use_mapping else True
        in_pdf = wh_no in pdf_wh_data

        if in_excel and in_pdf:
            st = pdf_wh_data[wh_no]
            total_p = st["total"]
            kept_p = len(st["pages"])
            dropped_p = st["dropped"]
            status = "Đã khớp & đóng số" + (f" (Lọc {dropped_p} trang rác)" if dropped_p > 0 else "")

            for page_idx in st["pages"]:
                out_doc.insert_pdf(doc, from_page=page_idx, to_page=page_idx)
                new_page = out_doc[-1]
                rect = new_page.rect

                if label:
                    # Tọa độ góc dưới bên phải
                    stamp_rect = fitz.Rect(rect.width - 145, rect.height - 48, rect.width - 20, rect.height - 15)
                    # Khung vien trang chong lem chu
                    new_page.draw_rect(stamp_rect, color=(0, 0, 0), fill=(1, 1, 1), width=1)
                    new_page.insert_textbox(
                        stamp_rect,
                        label,
                        fontsize=15,
                        fontname="helv",
                        color=(0, 0, 0),
                        align=fitz.TEXT_ALIGN_CENTER,
                    )

        elif in_excel and not in_pdf:
            total_p = 0
            kept_p = 0
            dropped_p = 0
            status = "❌ THIẾU: Có trong Excel nhưng không thấy trong PDF"

        else:  # in_pdf and not in_excel
            st = pdf_wh_data[wh_no]
            total_p = st["total"]
            kept_p = len(st["pages"])
            dropped_p = st["dropped"]
            status = "⚠️ THỪA: Có trong PDF nhưng chưa khai báo số trong Excel"

            for page_idx in st["pages"]:
                out_doc.insert_pdf(doc, from_page=page_idx, to_page=page_idx)

        csv_records.append([idx, label, wh_no, total_p, kept_p, dropped_p, status])

    out_doc.save(out_pdf, garbage=4, deflate=True)
    out_doc.close()
    doc.close()

    # 4. Xuat CSV doi chieu chuan UTF-8 co BOM
    with open(out_csv, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["STT", "Số đánh dấu", "Mã Warehouse Order (WH)", "Trang gốc", "Trang giữ", "Trang bỏ", "Trạng thái đối soát"])
        writer.writerows(csv_records)

    print(f"Xử lý hoàn tất! Đã xuất PDF: {out_pdf}")
    print(f"Đã xuất đối chiếu CSV: {out_csv}")


def main():
    parser = argparse.ArgumentParser(description="SAP IN WH Portable Engine v1.0")
    parser.add_argument("--input", required=True, help="File PDF goc xuat tu SAP")
    parser.add_argument("--mapping", default="", help="File CSV mapping (WH, Label) xuat tu Excel")
    parser.add_argument("--day", default="06", help="Tien to ngay neu tu sinh (vd: 06)")
    parser.add_argument("--start", type=int, default=1, help="So bat dau neu tu sinh (vd: 1)")
    parser.add_argument("--out-pdf", required=True, help="File PDF dau ra da dong so")
    parser.add_argument("--out-csv", required=True, help="File CSV bao cao doi chieu")
    args = parser.parse_args()

    process_wh_pdf(
        input_pdf=args.input,
        mapping_csv=args.mapping,
        day_prefix=args.day.strip(),
        start_num=args.start,
        out_pdf=args.out_pdf,
        out_csv=args.out_csv,
    )


if __name__ == "__main__":
    main()
