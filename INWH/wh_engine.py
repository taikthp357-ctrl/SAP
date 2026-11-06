"""
SAP IN WH Portable Engine (v1.1 - Slim & Smart)
Tu dong nhan dien WH tu PDF hoac doi soat theo bang mapping Excel.
Loc trang thua khong co vat tu, bao toan 100% barcode va dong so thu tu chuan.
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
    Trich xuat ma Warehouse Order tu trang PDF SAP thong minh.
    Ho tro nhieu kieu trinh bay: Label, 10 chu so (60..., 80...), 7-9 chu so.
    """
    # 1. Kiem tra theo nhan dinh danh ro rang
    m = re.search(
        r"(?:Warehouse\s*Order|WH\s*Order|WO|Số\s*WH|Order\s*No\.?|Lệnh\s*kho|Số\s*lệnh)[\s:#]*([0-9]{6,12})",
        page_text,
        re.IGNORECASE
    )
    if m:
        return m.group(1)

    # 2. Kiem tra ma 10 so tieu chuan SAP EWM / ERP (bat dau bang 60..., 80... hoac 10 so)
    m = re.search(r"\b(60[0-9]{8}|80[0-9]{8}|[0-9]{10})\b", page_text)
    if m:
        return m.group(1)

    # 3. Kiem tra ma so 7 - 9 chu so
    m = re.search(r"\b([0-9]{7,9})\b", page_text)
    if m:
        return m.group(1)

    return "UNKNOWN_WH"


def has_material_items(page_text: str) -> bool:
    """
    Kiem tra trang co chua dong vat tu thuc te hay chi la trang thua (chu ky, ngay gio in).
    """
    lines = [line.strip() for line in page_text.splitlines() if line.strip()]

    ignore_keywords = (
        "signature", "ký nhận", "người lập", "người giao", "người nhận",
        "printed by", "print date", "ngày in", "thời gian in", "page ", "trang ",
        "xác nhận", "warehouse order", "wh order"
    )

    material_code_pattern = re.compile(r"\b(?=[A-Z0-9\-]{8,18}\b)(?=.*[A-Z])(?=.*[0-9])[A-Z0-9\-]+\b")
    uom_pattern = re.compile(r"\b(PC|PCS|EA|SET|KG|M|L|BOX|ROL|CAR|UN|THÙNG|BỘ|CHIẾC|CÁI)\b", re.IGNORECASE)

    for line in lines:
        low_line = line.lower()
        if any(kw in low_line for kw in ignore_keywords):
            continue
        if re.fullmatch(r"[0-9]{6,12}", line):
            continue
        if material_code_pattern.search(line) or uom_pattern.search(line):
            return True

    return False


def load_mapping_file(mapping_csv_path: str):
    """
    Doc file mapping do Excel xuat ra.
    Tu dong loai bo cac ky tu rac hoac dong trong (vd chi co chu 'W' hoac space).
    """
    mapping_order = []
    mapping_dict = {}
    if not mapping_csv_path or not os.path.exists(mapping_csv_path):
        return mapping_order, mapping_dict

    with open(mapping_csv_path, mode="r", encoding="utf-8-sig", errors="ignore") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row:
                continue
            wh = row[0].strip()
            label = row[1].strip() if len(row) > 1 else ""
            
            # Loai bo tieu de va cac gia tri khong hop le (vd: chi co chu 'w', ngan hon 4 ky tu)
            if wh and wh.lower() not in ("mã wh", "wh_no", "warehouse_order", "ma_wh", "w"):
                if len(wh) >= 4:
                    if wh not in mapping_dict:
                        mapping_order.append(wh)
                    mapping_dict[wh] = label

    return mapping_order, mapping_dict


def process_wh_pdf(input_pdf: str, mapping_csv: str, day_prefix: str, start_num: int, out_pdf: str, out_csv: str):
    if not os.path.exists(input_pdf):
        raise FileNotFoundError(f"Khong tim thay file PDF: {input_pdf}")

    doc = fitz.open(input_pdf)
    out_doc = fitz.open()

    mapping_order, mapping_dict = load_mapping_file(mapping_csv)
    use_mapping = len(mapping_order) > 0

    # 1. Quet toan bo trang trong PDF va gom nhom theo ma WH
    pdf_wh_data = {}  # wh_no -> {"pages": [idx...], "dropped": 0, "total": 0}
    pdf_wh_order = []
    last_wh = None

    for page_idx in range(len(doc)):
        page = doc[page_idx]
        text = page.get_text("text")
        wh_no = extract_wh_number(text)

        # Neu trang sau khong co WH, ke thua WH cua trang lien truoc
        if wh_no == "UNKNOWN_WH" and last_wh is not None:
            wh_no = last_wh
        elif wh_no == "UNKNOWN_WH" and last_wh is None:
            wh_no = f"DOC_{page_idx+1}"

        last_wh = wh_no

        if wh_no not in pdf_wh_data:
            pdf_wh_data[wh_no] = {"pages": [], "dropped": 0, "total": 0}
            pdf_wh_order.append(wh_no)

        pdf_wh_data[wh_no]["total"] += 1
        if has_material_items(text):
            pdf_wh_data[wh_no]["pages"].append(page_idx)
        else:
            pdf_wh_data[wh_no]["dropped"] += 1

    # 2. Xac dinh danh sach WH va gan so thu tu
    final_wh_sequence = []
    if use_mapping:
        # Che do Map theo danh sach Excel
        for wh in mapping_order:
            if wh not in final_wh_sequence:
                final_wh_sequence.append(wh)
        for wh in pdf_wh_order:
            if wh not in final_wh_sequence:
                final_wh_sequence.append(wh)
    else:
        # CHE DO TU DONG QUET 100%: Lay truc tiep cac WH doc duoc tu PDF!
        final_wh_sequence = list(pdf_wh_order)
        cur_seq = start_num
        for wh in final_wh_sequence:
            mapping_dict[wh] = f"{day_prefix}{cur_seq:02d}"
            cur_seq += 1

    # 3. Loc trang, sap xep va dong so goc duoi ben phai
    csv_records = []

    for idx, wh_no in enumerate(final_wh_sequence, start=1):
        label = mapping_dict.get(wh_no, "")
        in_excel = (wh_no in mapping_dict) if use_mapping else True
        in_pdf = wh_no in pdf_wh_data

        if in_excel and in_pdf:
            st = pdf_wh_data[wh_no]
            total_p = st["total"]
            # Neu tat ca cac trang deu bi coi la khong co vat tu, giu lai it nhat trang 1 de khong bi mat phieu
            if len(st["pages"]) == 0 and total_p > 0:
                st["pages"].append(0)
                st["dropped"] = max(0, total_p - 1)

            kept_p = len(st["pages"])
            dropped_p = st["dropped"]

            if not use_mapping:
                status = "Đã nhận diện từ PDF & đóng số" + (f" (Lọc {dropped_p} trang rác)" if dropped_p > 0 else "")
            else:
                status = "Đã khớp & đóng số" + (f" (Lọc {dropped_p} trang rác)" if dropped_p > 0 else "")

            for page_idx in st["pages"]:
                out_doc.insert_pdf(doc, from_page=page_idx, to_page=page_idx)
                new_page = out_doc[-1]
                rect = new_page.rect

                if label:
                    stamp_rect = fitz.Rect(rect.width - 145, rect.height - 48, rect.width - 20, rect.height - 15)
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

    # 4. Xuat file CSV doi chieu chuan UTF-8 BOM
    with open(out_csv, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["STT", "Số đánh dấu", "Mã Warehouse Order (WH)", "Trang gốc", "Trang giữ", "Trang bỏ", "Trạng thái đối soát"])
        writer.writerows(csv_records)

    print(f"Xử lý thành công! Mode: {'Mapping' if use_mapping else 'Auto-Detect'}")


def main():
    parser = argparse.ArgumentParser(description="SAP IN WH Portable Engine v1.1")
    parser.add_argument("--input", required=True, help="File PDF goc xuat tu SAP")
    parser.add_argument("--mapping", default="", help="File CSV mapping tu Excel")
    parser.add_argument("--day", default="06-", help="Tien to ngay (vd: 06-)")
    parser.add_argument("--start", type=int, default=1, help="So bat dau (vd: 1)")
    parser.add_argument("--out-pdf", required=True, help="File PDF dau ra")
    parser.add_argument("--out-csv", required=True, help="File CSV doi chieu")
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
