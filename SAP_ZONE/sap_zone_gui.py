# -*- coding: utf-8 -*-
"""
SAP ZONE GUI Application (v1.6)
Giao diện trực quan 2 Tab tối giản, tự động hóa quy trình:
- Tab 1: Vận hành chính (1-Click Run, Nhận diện dữ liệu, Xem & Sao chép 3 cột kết quả)
- Tab 2: Cấu hình hệ thống (SAP, EWM, Gom nhóm, In ấn & Zone Mapping)
"""

import os
import sys
import json
import time
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from sap_zone_core import (
    ConfigManager,
    ZoneManager,
    SapConnector,
    SapZonePipeline,
    ExcelResultExporter,
    parse_paste_data
)

SAMPLE_INPUT = """BIW32011852HA\t12\tOP10
BIW32011853HA\t24\tOP20
BIW32011854HA\t6\tZone 3.1
BIW32011855HA\t15\tOP40
BIW32011856HA\t30\tOP50"""


class SapZoneApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("SAP ZONE — Ra Đơn & In Phiếu Tự Động (v1.6)")
        self.root.geometry("1020x720")
        self.root.minsize(880, 600)

        # Cấu hình & Dữ liệu
        self.config = ConfigManager.load_config()
        self.zone_mgr = ZoneManager()
        self.pipeline: Optional[SapZonePipeline] = None
        self.worker_thread: Optional[threading.Thread] = None
        self.is_running = False

        # Kết quả chạy gần nhất
        self.current_results = []
        self.current_pdf_folder = self.config.get("pdf_output_folder", os.path.abspath("./OUTPUT_PDF"))

        # Thiết lập giao diện
        self._setup_styles()
        self._build_ui()
        self._check_sap_status_async()

    def _setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # Màu sắc chủ đạo
        bg_main = "#f8fafc"
        card_bg = "#ffffff"
        text_dark = "#0f172a"
        accent_blue = "#1e40af"

        self.root.configure(bg=bg_main)

        style.configure(".", font=("Segoe UI", 10), background=bg_main, foreground=text_dark)
        style.configure("TNotebook", background=bg_main, tabmargins=[6, 6, 6, 0])
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[16, 8], background="#e2e8f0", foreground="#334155")
        style.map("TNotebook.Tab",
                  background=[("selected", "#ffffff")],
                  foreground=[("selected", accent_blue)])

        style.configure("Card.TLabelframe", background=card_bg, relief="solid", borderwidth=1)
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"), background=card_bg, foreground=accent_blue)

        style.configure("Action.TButton", font=("Segoe UI", 11, "bold"), background=accent_blue, foreground="#ffffff", padding=[12, 8])
        style.map("Action.TButton",
                  background=[("active", "#1d4ed8"), ("pressed", "#1e3a8a")])

        style.configure("Success.TButton", font=("Segoe UI", 9, "bold"), background="#15803d", foreground="#ffffff", padding=[8, 5])
        style.map("Success.TButton",
                  background=[("active", "#16a34a"), ("pressed", "#14532d")])

        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#f1f5f9", foreground="#1e293b", relief="flat")
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=26, background="#ffffff", fieldbackground="#ffffff")

    def _build_ui(self):
        # 1. Header Bar
        header_frame = tk.Frame(self.root, bg="#1e3a8a", height=60)
        header_frame.pack(fill="x", side="top")

        lbl_title = tk.Label(
            header_frame,
            text="⚡ SAP ZONE — TỰ ĐỘNG RA ĐƠN & IN PHIẾU",
            font=("Segoe UI", 14, "bold"),
            fg="#ffffff",
            bg="#1e3a8a"
        )
        lbl_title.pack(side="left", padx=16, pady=10)

        lbl_sub = tk.Label(
            header_frame,
            text="EWM APFL • FIFO/FEFO • MIGO 311 • WHO F00112",
            font=("Segoe UI", 9),
            fg="#93c5fd",
            bg="#1e3a8a"
        )
        lbl_sub.pack(side="left", padx=4, pady=12)

        # SAP Status Badge trên Header
        self.sap_badge_var = tk.StringVar(value="⏳ Đang kiểm tra SAP...")
        self.lbl_sap_badge = tk.Label(
            header_frame,
            textvariable=self.sap_badge_var,
            font=("Segoe UI", 9, "bold"),
            fg="#ffffff",
            bg="#475569",
            padx=10,
            pady=4,
            relief="flat"
        )
        self.lbl_sap_badge.pack(side="right", padx=16, pady=12)

        # 2. Main Notebook (2 Tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=10)

        self.tab_main = ttk.Frame(self.notebook)
        self.tab_config = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_main, text="  🚀 VẬN HÀNH CHÍNH  ")
        self.notebook.add(self.tab_config, text="  ⚙️ CẤU HÌNH HỆ THỐNG  ")

        self._build_tab_main()
        self._build_tab_config()

    # ==========================================================================
    # TAB 1: VẬN HÀNH CHÍNH (1-CLICK RUN & BẢNG KẾT QUẢ 3 CỘT)
    # ==========================================================================
    def _build_tab_main(self):
        # Chia layout 2 phần: Trên (Dán dữ liệu & Nút Chạy), Dưới (Bảng kết quả 3 cột)
        paned = ttk.PanedWindow(self.tab_main, orient="vertical")
        paned.pack(fill="both", expand=True, padx=8, pady=8)

        # --- KHUNG TRÊN: NHẬP LIỆU & ĐIỀU KHIỂN ---
        top_frame = ttk.LabelFrame(paned, text=" 1. Dán Dữ Liệu Yêu Cầu Ra Đơn (Từ Excel) ", style="Card.TLabelframe")
        paned.add(top_frame, weight=2)

        # Khung nút thao tác nhanh trên ô nhập
        input_toolbar = tk.Frame(top_frame, bg="#ffffff")
        input_toolbar.pack(fill="x", padx=10, pady=(6, 4))

        lbl_hint = tk.Label(
            input_toolbar,
            text="Định dạng cột: [Mã hàng]  [Số lượng]  [Trạm] (Ngăn cách bởi Tab hoặc Khoảng trắng)",
            font=("Segoe UI", 9, "italic"),
            fg="#64748b",
            bg="#ffffff"
        )
        lbl_hint.pack(side="left")

        btn_clear = ttk.Button(input_toolbar, text="🗑 Xóa", command=self._clear_input, width=7)
        btn_clear.pack(side="right", padx=2)

        btn_sample = ttk.Button(input_toolbar, text="📝 Nạp mẫu thử", command=self._load_sample, width=13)
        btn_sample.pack(side="right", padx=2)

        btn_paste = ttk.Button(input_toolbar, text="📋 Dán từ Clipboard", command=self._paste_clipboard, width=16)
        btn_paste.pack(side="right", padx=2)

        # Ô Text nhập dữ liệu
        text_frame = tk.Frame(top_frame, bg="#ffffff")
        text_frame.pack(fill="both", expand=True, padx=10, pady=4)

        self.txt_input = tk.Text(
            text_frame,
            height=6,
            font=("Consolas", 10),
            wrap="none",
            relief="solid",
            borderwidth=1,
            bg="#f8fafc"
        )
        scroll_y = ttk.Scrollbar(text_frame, orient="vertical", command=self.txt_input.yview)
        scroll_x = ttk.Scrollbar(text_frame, orient="horizontal", command=self.txt_input.xview)
        self.txt_input.configure(xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.txt_input.pack(side="left", fill="both", expand=True)

        # Khung điều khiển: Nút Chạy to & Tiến trình
        ctrl_frame = tk.Frame(top_frame, bg="#ffffff")
        ctrl_frame.pack(fill="x", padx=10, pady=(6, 8))

        # Checkbox Chế độ Mô phỏng (Mock mode)
        self.mock_mode_var = tk.BooleanVar(value=False)
        chk_mock = tk.Checkbutton(
            ctrl_frame,
            text="Chạy mô phỏng (Test không cần mở SAP)",
            variable=self.mock_mode_var,
            font=("Segoe UI", 9),
            bg="#ffffff",
            activebackground="#ffffff"
        )
        chk_mock.pack(side="left", padx=4)

        # Nút Chạy chính to nổi bật
        self.btn_run = tk.Button(
            ctrl_frame,
            text="▶  BẮT ĐẦU CHẠY: RA ĐƠN & IN PHIẾU",
            font=("Segoe UI", 11, "bold"),
            bg="#1e40af",
            fg="#ffffff",
            activebackground="#1d4ed8",
            activeforeground="#ffffff",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=self._on_btn_run_clicked
        )
        self.btn_run.pack(side="right", padx=6)

        # Thanh tiến trình & Trạng thái
        status_bar_frame = tk.Frame(top_frame, bg="#ffffff")
        status_bar_frame.pack(fill="x", padx=10, pady=(0, 8))

        self.progress_bar = ttk.Progressbar(status_bar_frame, mode="determinate")
        self.progress_bar.pack(fill="x", side="top", pady=2)

        self.status_text_var = tk.StringVar(value="Sẵn sàng nhận lệnh.")
        self.lbl_status = tk.Label(
            status_bar_frame,
            textvariable=self.status_text_var,
            font=("Segoe UI", 9, "bold"),
            fg="#0369a1",
            bg="#ffffff",
            anchor="w"
        )
        self.lbl_status.pack(fill="x", side="left", pady=2)

        # --- KHUNG DƯỚI: BẢNG KẾT QUẢ 3 CỘT CHUẨN ---
        bottom_frame = ttk.LabelFrame(paned, text=" 2. Bảng Kết Quả Theo Dõi (3 Cột Chuẩn: Số OB • Mã Hàng • Số Đơn) ", style="Card.TLabelframe")
        paned.add(bottom_frame, weight=3)

        # Thanh công cụ bảng kết quả
        table_toolbar = tk.Frame(bottom_frame, bg="#ffffff")
        table_toolbar.pack(fill="x", padx=10, pady=(6, 4))

        self.lbl_row_count = tk.Label(
            table_toolbar,
            text="Tổng cộng: 0 dòng",
            font=("Segoe UI", 9, "bold"),
            fg="#475569",
            bg="#ffffff"
        )
        self.lbl_row_count.pack(side="left")

        btn_open_pdf = ttk.Button(table_toolbar, text="📁 Mở Thư Mục PDF", command=self._open_pdf_folder)
        btn_open_pdf.pack(side="right", padx=3)

        btn_export_excel = ttk.Button(table_toolbar, text="📊 Xuất Excel (KET_QUA)", command=self._export_excel)
        btn_export_excel.pack(side="right", padx=3)

        btn_copy_3col = tk.Button(
            table_toolbar,
            text="📋 Sao Chép 3 Cột (Dán Vào File Theo Dõi)",
            font=("Segoe UI", 9, "bold"),
            bg="#15803d",
            fg="#ffffff",
            activebackground="#16a34a",
            activeforeground="#ffffff",
            relief="flat",
            padx=10,
            pady=4,
            cursor="hand2",
            command=self._copy_3_columns_to_clipboard
        )
        btn_copy_3col.pack(side="right", padx=3)

        # Bảng Treeview 3 cột
        tree_container = tk.Frame(bottom_frame, bg="#ffffff")
        tree_container.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        columns = ("ob", "material", "who")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode="extended")

        self.tree.heading("ob", text="Số OB")
        self.tree.heading("material", text="Mã Hàng (Material)")
        self.tree.heading("who", text="Số Đơn (WHO)")

        self.tree.column("ob", width=200, anchor="center")
        self.tree.column("material", width=250, anchor="w")
        self.tree.column("who", width=380, anchor="w")

        tree_scroll_y = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set)

        tree_scroll_y.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

        # Đổi màu dòng chẵn lẻ
        self.tree.tag_configure("oddrow", background="#f8fafc")
        self.tree.tag_configure("evenrow", background="#ffffff")

    # ==========================================================================
    # TAB 2: CẤU HÌNH HỆ THỐNG
    # ==========================================================================
    def _build_tab_config(self):
        # Scrollable container cho Tab cấu hình
        canvas = tk.Canvas(self.tab_config, bg="#f8fafc", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.tab_config, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f8fafc")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        scrollbar.pack(side="right", fill="y")

        # 1. Khung Kết nối SAP
        grp_sap = ttk.LabelFrame(scrollable_frame, text=" 1. Kết Nối SAP GUI ", style="Card.TLabelframe")
        grp_sap.pack(fill="x", padx=12, pady=6)

        row1 = tk.Frame(grp_sap, bg="#ffffff")
        row1.pack(fill="x", padx=10, pady=6)

        tk.Label(row1, text="Hệ thống (System):", bg="#ffffff", width=18, anchor="w").pack(side="left")
        self.cfg_sap_sys = ttk.Entry(row1, width=12)
        self.cfg_sap_sys.insert(0, str(self.config.get("sap_system", "PRD")))
        self.cfg_sap_sys.pack(side="left", padx=4)

        tk.Label(row1, text="Client:", bg="#ffffff", width=8, anchor="w").pack(side="left", padx=(16, 0))
        self.cfg_sap_client = ttk.Entry(row1, width=10)
        self.cfg_sap_client.insert(0, str(self.config.get("sap_client", "300")))
        self.cfg_sap_client.pack(side="left", padx=4)

        tk.Label(row1, text="Session Index:", bg="#ffffff", width=14, anchor="w").pack(side="left", padx=(16, 0))
        self.cfg_sap_idx = ttk.Spinbox(row1, from_=0, to=5, width=6)
        self.cfg_sap_idx.set(str(self.config.get("sap_session_idx", 0)))
        self.cfg_sap_idx.pack(side="left", padx=4)

        btn_test_sap = ttk.Button(row1, text="🔌 Thử Kết Nối", command=self._test_sap_connection)
        btn_test_sap.pack(side="right", padx=4)

        # 2. Khung Tham số EWM & Kho
        grp_ewm = ttk.LabelFrame(scrollable_frame, text=" 2. Tham Số Kho EWM & Ra Đơn ", style="Card.TLabelframe")
        grp_ewm.pack(fill="x", padx=12, pady=6)

        row2 = tk.Frame(grp_ewm, bg="#ffffff")
        row2.pack(fill="x", padx=10, pady=6)

        tk.Label(row2, text="Mã kho (Warehouse):", bg="#ffffff", width=18, anchor="w").pack(side="left")
        self.cfg_whse = ttk.Entry(row2, width=12)
        self.cfg_whse.insert(0, str(self.config.get("ewm_warehouse", "VF02")))
        self.cfg_whse.pack(side="left", padx=4)

        tk.Label(row2, text="Storage Type:", bg="#ffffff", width=14, anchor="w").pack(side="left", padx=(16, 0))
        self.cfg_storage = ttk.Entry(row2, width=12)
        self.cfg_storage.insert(0, str(self.config.get("ewm_storage_type", "APFL")))
        self.cfg_storage.pack(side="left", padx=4)

        tk.Label(row2, text="Chủ hàng (Owner):", bg="#ffffff", width=16, anchor="w").pack(side="left", padx=(16, 0))
        self.cfg_owner = ttk.Entry(row2, width=12)
        self.cfg_owner.insert(0, str(self.config.get("ewm_owner", "3093")))
        self.cfg_owner.pack(side="left", padx=4)

        row2b = tk.Frame(grp_ewm, bg="#ffffff")
        row2b.pack(fill="x", padx=10, pady=(0, 6))

        tk.Label(row2b, text="Plant MIGO:", bg="#ffffff", width=18, anchor="w").pack(side="left")
        self.cfg_plant = ttk.Entry(row2b, width=12)
        self.cfg_plant.insert(0, str(self.config.get("plant", "3093")))
        self.cfg_plant.pack(side="left", padx=4)

        tk.Label(row2b, text="SLoc Xuất:", bg="#ffffff", width=10, anchor="w").pack(side="left", padx=(12, 0))
        self.cfg_sloc_from = ttk.Entry(row2b, width=8)
        self.cfg_sloc_from.insert(0, str(self.config.get("sloc_from", "1005")))
        self.cfg_sloc_from.pack(side="left", padx=4)

        tk.Label(row2b, text="SLoc Nhận:", bg="#ffffff", width=10, anchor="w").pack(side="left", padx=(12, 0))
        self.cfg_sloc_to = ttk.Entry(row2b, width=8)
        self.cfg_sloc_to.insert(0, str(self.config.get("sloc_to", "3001")))
        self.cfg_sloc_to.pack(side="left", padx=4)

        # 3. Khung Giới hạn Gom nhóm & Phân bổ FIFO
        grp_group = ttk.LabelFrame(scrollable_frame, text=" 3. Quy Tắc Gom Nhóm Đơn & Phân Bổ FIFO ", style="Card.TLabelframe")
        grp_group.pack(fill="x", padx=12, pady=6)

        row3 = tk.Frame(grp_group, bg="#ffffff")
        row3.pack(fill="x", padx=10, pady=6)

        tk.Label(row3, text="Số mã tối đa / OB:", bg="#ffffff", width=18, anchor="w").pack(side="left")
        self.cfg_max_mats = ttk.Spinbox(row3, from_=1, to=20, width=6)
        self.cfg_max_mats.set(str(self.config.get("max_materials_per_group", 5)))
        self.cfg_max_mats.pack(side="left", padx=4)

        tk.Label(row3, text="Số dòng tối đa / OB:", bg="#ffffff", width=18, anchor="w").pack(side="left", padx=(16, 0))
        self.cfg_max_lines = ttk.Spinbox(row3, from_=1, to=50, width=6)
        self.cfg_max_lines.set(str(self.config.get("max_lines_per_group", 20)))
        self.cfg_max_lines.pack(side="left", padx=4)

        tk.Label(row3, text="Chiến lược:", bg="#ffffff", width=12, anchor="w").pack(side="left", padx=(16, 0))
        self.cfg_strategy = ttk.Combobox(row3, values=["FIFO", "FEFO"], width=8, state="readonly")
        self.cfg_strategy.set(str(self.config.get("allocation_strategy", "FIFO")))
        self.cfg_strategy.pack(side="left", padx=4)

        # 4. Khung In ấn & Lưu PDF
        grp_print = ttk.LabelFrame(scrollable_frame, text=" 4. Cấu Hình In Phiếu SAP (F00112) & Xuất PDF ", style="Card.TLabelframe")
        grp_print.pack(fill="x", padx=12, pady=6)

        row4 = tk.Frame(grp_print, bg="#ffffff")
        row4.pack(fill="x", padx=10, pady=6)

        tk.Label(row4, text="Số bản in SAP (Copies):", bg="#ffffff", width=20, anchor="w").pack(side="left")
        self.cfg_copies = ttk.Spinbox(row4, from_=1, to=5, width=6)
        self.cfg_copies.set(str(self.config.get("sap_print_copies", 2)))
        self.cfg_copies.pack(side="left", padx=4)

        self.cfg_auto_print = tk.BooleanVar(value=bool(self.config.get("auto_print_pdf", True)))
        chk_auto_print = tk.Checkbutton(
            row4,
            text="Tự động gọi in và lưu PDF khi hoàn thành",
            variable=self.cfg_auto_print,
            font=("Segoe UI", 9),
            bg="#ffffff"
        )
        chk_auto_print.pack(side="left", padx=16)

        row4b = tk.Frame(grp_print, bg="#ffffff")
        row4b.pack(fill="x", padx=10, pady=(0, 6))

        tk.Label(row4b, text="Thư mục lưu PDF:", bg="#ffffff", width=20, anchor="w").pack(side="left")
        self.cfg_pdf_dir = ttk.Entry(row4b, width=45)
        self.cfg_pdf_dir.insert(0, str(self.config.get("pdf_output_folder", os.path.abspath("./OUTPUT_PDF"))))
        self.cfg_pdf_dir.pack(side="left", padx=4, fill="x", expand=True)

        btn_browse_pdf = ttk.Button(row4b, text="📂 Chọn...", width=9, command=self._browse_pdf_dir)
        btn_browse_pdf.pack(side="left", padx=4)

        # 5. Khung Zone Mapping
        grp_zone = ttk.LabelFrame(scrollable_frame, text=" 5. Dữ Liệu Ánh Xạ Trạm & Zone (zone_map.json) ", style="Card.TLabelframe")
        grp_zone.pack(fill="x", padx=12, pady=6)

        row5 = tk.Frame(grp_zone, bg="#ffffff")
        row5.pack(fill="x", padx=10, pady=6)

        zone_count = len(self.zone_mgr.mappings)
        self.lbl_zone_status = tk.Label(
            row5,
            text=f"Hiện có: {zone_count} quy tắc ánh xạ Mã hàng / Trạm -> Zone -> Tổ trưởng.",
            font=("Segoe UI", 9),
            fg="#0f172a",
            bg="#ffffff"
        )
        self.lbl_zone_status.pack(side="left")

        btn_reload_zone = ttk.Button(row5, text="🔄 Tải Lại", command=self._reload_zone_map)
        btn_reload_zone.pack(side="right", padx=4)

        btn_open_zone = ttk.Button(row5, text="📝 Xem File JSON", command=self._open_zone_file)
        btn_open_zone.pack(side="right", padx=4)

        # Nút Lưu Cấu Hình
        save_frame = tk.Frame(scrollable_frame, bg="#f8fafc")
        save_frame.pack(fill="x", padx=12, pady=16)

        btn_save_cfg = tk.Button(
            save_frame,
            text="💾  LƯU CẤU HÌNH HỆ THỐNG",
            font=("Segoe UI", 11, "bold"),
            bg="#15803d",
            fg="#ffffff",
            activebackground="#16a34a",
            activeforeground="#ffffff",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=self._save_config
        )
        btn_save_cfg.pack(side="right")

    # ==========================================================================
    # LOGIC TAB 1: THAO TÁC NHẬP LIỆU & CHẠY TIẾN TRÌNH
    # ==========================================================================
    def _clear_input(self):
        self.txt_input.delete("1.0", tk.END)

    def _load_sample(self):
        self.txt_input.delete("1.0", tk.END)
        self.txt_input.insert("1.0", SAMPLE_INPUT)

    def _paste_clipboard(self):
        try:
            cb_text = self.root.clipboard_get()
            if cb_text:
                self.txt_input.delete("1.0", tk.END)
                self.txt_input.insert("1.0", cb_text)
        except Exception as e:
            messagebox.showwarning("Clipboard", f"Không đọc được nội dung bộ nhớ đệm: {e}")

    def _on_btn_run_clicked(self):
        if self.is_running:
            # Người dùng bấm dừng
            if self.pipeline:
                self.pipeline.request_stop()
            self.status_text_var.set("Đang yêu cầu dừng tiến trình...")
            return

        # Đọc và phân tích dữ liệu dán
        raw_text = self.txt_input.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng dán danh sách mã hàng và số lượng từ Excel trước khi bấm Chạy!")
            return

        requests = parse_paste_data(raw_text)
        if not requests:
            messagebox.showwarning("Dữ liệu không hợp lệ", "Không nhận diện được dòng dữ liệu hợp lệ nào!\nVui lòng kiểm tra định dạng: [Mã hàng] [Số lượng] [Trạm].")
            return

        # Xác nhận chạy
        mock_mode = self.mock_mode_var.get()
        mode_str = "MÔ PHỎNG (TEST)" if mock_mode else "TRỰC TIẾP TRÊN SAP"
        confirm_msg = f"Xác nhận thực hiện quy trình ({mode_str}) cho {len(requests)} dòng yêu cầu?\n\nHệ thống sẽ:\n1. Đọc tồn MON (APFL)\n2. Phân bổ FIFO & Gom nhóm theo Zone\n3. Tạo MIGO 311\n4. Tra cứu WHO cho từng Material\n5. In phiếu SAP 2 bản lưu PDF"
        if not messagebox.askyesno("Xác nhận chạy", confirm_msg):
            return

        # Bắt đầu luồng chạy
        self.is_running = True
        self.btn_run.configure(text="⏹  DỪNG LẠI", bg="#dc2626", activebackground="#b91c1c")
        self.progress_bar["value"] = 0
        self.status_text_var.set("Bắt đầu xử lý...")

        # Xóa bảng kết quả cũ
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.lbl_row_count.configure(text="Đang xử lý...")

        self.pipeline = SapZonePipeline(self.config, self.zone_mgr)
        self.worker_thread = threading.Thread(
            target=self._run_pipeline_worker,
            args=(requests, mock_mode),
            daemon=True
        )
        self.worker_thread.start()

    def _run_pipeline_worker(self, requests, mock_mode):
        try:
            SapConnector.init_thread_com()
            res = self.pipeline.run(
                requests=requests,
                progress_cb=self._update_progress_from_worker,
                status_cb=self._update_status_from_worker,
                is_mock=mock_mode
            )
            self.root.after(0, self._on_pipeline_finished, res)
        except Exception as ex:
            self.root.after(0, self._on_pipeline_error, str(ex))
        finally:
            SapConnector.uninit_thread_com()

    def _update_progress_from_worker(self, pct: int):
        self.root.after(0, lambda: self.progress_bar.configure(value=pct))

    def _update_status_from_worker(self, msg: str):
        self.root.after(0, lambda: self.status_text_var.set(msg))

    def _on_pipeline_finished(self, summary: dict):
        self.is_running = False
        self.btn_run.configure(text="▶  BẮT ĐẦU CHẠY: RA ĐƠN & IN PHIẾU", bg="#1e40af", activebackground="#1d4ed8")

        if summary.get("success"):
            records = summary.get("records", [])
            self.current_results = records

            # Đổ dữ liệu vào Treeview
            for idx, r in enumerate(records):
                tag = "evenrow" if idx % 2 == 0 else "oddrow"
                self.tree.insert("", "end", values=(r["ob"], r["material"], r["who"]), tags=(tag,))

            self.lbl_row_count.configure(text=f"Tổng cộng: {len(records)} dòng kết quả")
            self.status_text_var.set("Hoàn thành quy trình thành công!")

            # Thông báo hoàn tất
            pdf_count = len(summary.get("pdf_files", []))
            pdf_info = f"\nĐã xuất {pdf_count} file PDF phiếu in." if pdf_count > 0 else ""
            messagebox.showinfo("Hoàn Tất", f"Đã xử lý xong {len(records)} dòng kết quả!{pdf_info}\n\nBạn có thể bấm 'Sao Chép 3 Cột' để dán vào file Excel theo dõi.")
        else:
            self.status_text_var.set(f"Thất bại: {summary.get('message', 'Lỗi không xác định')}")
            messagebox.showerror("Thông Báo Lỗi", summary.get("message", "Đã có lỗi xảy ra trong quá trình thực hiện!"))

    def _on_pipeline_error(self, err_msg: str):
        self.is_running = False
        self.btn_run.configure(text="▶  BẮT ĐẦU CHẠY: RA ĐƠN & IN PHIẾU", bg="#1e40af", activebackground="#1d4ed8")
        self.status_text_var.set(f"Lỗi: {err_msg}")
        messagebox.showerror("Lỗi Nghiêm Trọng", f"Tiến trình bị dừng do ngoại lệ:\n{err_msg}")

    # ==========================================================================
    # CÁC NÚT THAO TÁC TRÊN BẢNG KẾT QUẢ
    # ==========================================================================
    def _copy_3_columns_to_clipboard(self):
        """Sao chép toàn bộ bảng kết quả 3 cột theo định dạng TAB (chuẩn dán Excel)."""
        if not self.current_results:
            messagebox.showinfo("Sao chép", "Chưa có kết quả để sao chép!")
            return

        lines = []
        for r in self.current_results:
            lines.append(f"{r['ob']}\t{r['material']}\t{r['who']}")
        tsv_text = "\r\n".join(lines)

        self.root.clipboard_clear()
        self.root.clipboard_append(tsv_text)
        self.root.update()

        messagebox.showinfo("Đã sao chép", f"Đã sao chép {len(self.current_results)} dòng (3 cột ngăn cách Tab) vào Clipboard!\nBạn chỉ cần mở file Excel theo dõi và nhấn Ctrl+V.")

    def _export_excel(self):
        if not self.current_results:
            messagebox.showinfo("Xuất Excel", "Chưa có kết quả để xuất Excel!")
            return

        default_filename = f"KET_QUA_RA_DON_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=default_filename
        )
        if not file_path:
            return

        ok = ExcelResultExporter.export(self.current_results, file_path)
        if ok:
            messagebox.showinfo("Xuất Excel", f"Đã xuất thành công kết quả ra file:\n{file_path}")
        else:
            messagebox.showerror("Lỗi", "Không thể xuất file Excel. Vui lòng kiểm tra quyền ghi hoặc thư viện openpyxl.")

    def _open_pdf_folder(self):
        pdf_folder = self.config.get("pdf_output_folder", os.path.abspath("./OUTPUT_PDF"))
        os.makedirs(pdf_folder, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(pdf_folder)
            else:
                subprocess.Popen(["xdg-open", pdf_folder])
        except Exception as ex:
            messagebox.showerror("Mở thư mục", f"Không thể mở thư mục:\n{ex}")

    # ==========================================================================
    # LOGIC TAB 2: CẤU HÌNH & KIỂM TRA SAP
    # ==========================================================================
    def _check_sap_status_async(self):
        """Kiểm tra nền trạng thái kết nối SAP khi vừa mở app."""
        def check():
            res = SapConnector.test_connection(
                session_idx=int(self.config.get("sap_session_idx", -1)),
                preferred_sys=self.config.get("sap_system", ""),
                preferred_client=self.config.get("sap_client", "")
            )
            if res.get("ok"):
                txt = f"🟢 SAP Sẵn Sàng ({res.get('system')}/{res.get('client')} - User: {res.get('user')})"
                bg_color = "#15803d"
            else:
                txt = "🔴 Chưa Kết Nối SAP"
                bg_color = "#b91c1c"
            self.root.after(0, lambda: self._update_sap_badge(txt, bg_color))

        threading.Thread(target=check, daemon=True).start()

    def _update_sap_badge(self, text: str, bg_color: str):
        self.sap_badge_var.set(text)
        self.lbl_sap_badge.configure(bg=bg_color)

    def _test_sap_connection(self):
        try:
            s_idx = int(self.cfg_sap_idx.get())
        except ValueError:
            s_idx = 0
        res = SapConnector.test_connection(
            session_idx=s_idx,
            preferred_sys=self.cfg_sap_sys.get().strip(),
            preferred_client=self.cfg_sap_client.get().strip()
        )
        if res.get("ok"):
            self._update_sap_badge(f"🟢 SAP Sẵn Sàng ({res.get('system')}/{res.get('client')})", "#15803d")
            messagebox.showinfo(
                "Kết Nối SAP",
                f"Kết nối SAP thành công!\n\n"
                f"• Hệ thống (System): {res.get('system')}\n"
                f"• Phân hệ (Client): {res.get('client')}\n"
                f"• Tài khoản (User): {res.get('user')}\n"
                f"• Giao dịch hiện tại: {res.get('tcode') or '(Easy Access)'}"
            )
        else:
            self._update_sap_badge("🔴 Chưa Kết Nối SAP", "#b91c1c")
            messagebox.showerror("Lỗi Kết Nối", f"Không thể kết nối phiên SAP:\n{res.get('error')}")

    def _browse_pdf_dir(self):
        d = filedialog.askdirectory(initialdir=self.cfg_pdf_dir.get())
        if d:
            self.cfg_pdf_dir.delete(0, tk.END)
            self.cfg_pdf_dir.insert(0, os.path.normpath(d))

    def _reload_zone_map(self):
        self.zone_mgr.load_mappings()
        count = len(self.zone_mgr.mappings)
        self.lbl_zone_status.configure(text=f"Hiện có: {count} quy tắc ánh xạ Mã hàng / Trạm -> Zone -> Tổ trưởng.")
        messagebox.showinfo("Zone Mapping", f"Đã nạp lại thành công {count} quy tắc ánh xạ!")

    def _open_zone_file(self):
        path = os.path.abspath(self.zone_mgr.MAPPING_FILE)
        if os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không mở được file:\n{e}")
        else:
            messagebox.showwarning("Không tìm thấy", f"File không tồn tại: {path}")

    def _save_config(self):
        try:
            new_cfg = {
                "sap_system": self.cfg_sap_sys.get().strip(),
                "sap_client": self.cfg_sap_client.get().strip(),
                "sap_session_idx": int(self.cfg_sap_idx.get()),
                "ewm_warehouse": self.cfg_whse.get().strip(),
                "ewm_storage_type": self.cfg_storage.get().strip(),
                "ewm_owner": self.cfg_owner.get().strip(),
                "plant": self.cfg_plant.get().strip(),
                "sloc_from": self.cfg_sloc_from.get().strip(),
                "sloc_to": self.cfg_sloc_to.get().strip(),
                "max_materials_per_group": int(self.cfg_max_mats.get()),
                "max_lines_per_group": int(self.cfg_max_lines.get()),
                "allocation_strategy": self.cfg_strategy.get().strip(),
                "sap_print_copies": int(self.cfg_copies.get()),
                "auto_print_pdf": self.cfg_auto_print.get(),
                "pdf_output_folder": self.cfg_pdf_dir.get().strip()
            }
            ok = ConfigManager.save_config(new_cfg)
            if ok:
                self.config = new_cfg
                messagebox.showinfo("Cấu Hình", "Đã lưu thành công cấu hình vào config.json!")
            else:
                messagebox.showerror("Lỗi", "Không thể ghi file cấu hình config.json!")
        except Exception as ex:
            messagebox.showerror("Lỗi Dữ Liệu", f"Dữ liệu cấu hình không hợp lệ:\n{ex}")


def main():
    root = tk.Tk()
    app = SapZoneApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
