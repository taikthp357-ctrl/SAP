# -*- coding: utf-8 -*-
"""
SAP ZONE GUI Application (v1.6 Pro - Revised)
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
from typing import Optional, List, Dict, Any

from sap_zone_core import (
    ConfigManager,
    ZoneManager,
    SapConnector,
    SapZonePipeline,
    ExcelResultExporter,
    parse_paste_data
)

SAMPLE_INPUT = """BIW32211725AA\t300\tHùng/Long
BIW32211727AA\t300\tHùng/Long
BIW32011854HA\t6\tZone 3.1
BIW32011855HA\t15\t1.2"""


class SapZoneApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("SAP ZONE — Ra Đơn & In Phiếu Tự Động (v1.6 Pro)")
        self.root.geometry("1060x740")
        self.root.minsize(900, 620)

        # Cấu hình & Dữ liệu
        self.config = ConfigManager.load_config()
        self.zone_mgr = ZoneManager()
        self.pipeline: Optional[SapZonePipeline] = None
        self.worker_thread: Optional[threading.Thread] = None
        self.is_running = False

        # Kết quả chạy gần nhất
        self.current_results: List[Dict[str, str]] = []
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
            font=("Segoe UI", 13, "bold"),
            fg="#ffffff",
            bg="#1e3a8a"
        )
        lbl_title.pack(side="left", padx=15, pady=12)

        lbl_sub = tk.Label(
            header_frame,
            text="EWM APFL • FIFO/FEFO • MIGO 311 • WHO F00112",
            font=("Segoe UI", 9),
            fg="#93c5fd",
            bg="#1e3a8a"
        )
        lbl_sub.pack(side="left", padx=(0, 15), pady=12)

        # Trạng thái SAP Badge
        self.sap_badge_var = tk.StringVar(value="⏳ Đang kiểm tra SAP...")
        self.lbl_sap_badge = tk.Label(
            header_frame,
            textvariable=self.sap_badge_var,
            font=("Segoe UI", 9, "bold"),
            fg="#ffffff",
            bg="#334155",
            padx=12,
            pady=4,
            relief="flat"
        )
        self.lbl_sap_badge.pack(side="right", padx=15, pady=12)

        # 2. Main Notebook (2 Tab)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        self.tab_main = tk.Frame(self.notebook, bg="#f8fafc")
        self.tab_config = tk.Frame(self.notebook, bg="#f8fafc")

        self.notebook.add(self.tab_main, text="  🚀 VẬN HÀNH CHÍNH  ")
        self.notebook.add(self.tab_config, text="  ⚙️ CẤU HÌNH HỆ THỐNG  ")

        self._build_tab_main()
        self._build_tab_config()

    # ==========================================================================
    # TAB 1: VẬN HÀNH CHÍNH
    # ==========================================================================
    def _build_tab_main(self):
        paned = ttk.PanedWindow(self.tab_main, orient="vertical")
        paned.pack(fill="both", expand=True, padx=6, pady=6)

        # --- KHUNG TRÊN: NHẬP DỮ LIỆU DÁN ---
        top_frame = ttk.LabelFrame(paned, text=" 1. Dán Dữ Liệu Yêu Cầu Ra Đơn (Từ Excel) ", style="Card.TLabelframe")
        paned.add(top_frame, weight=2)

        # Toolbar nhập liệu
        input_toolbar = tk.Frame(top_frame, bg="#ffffff")
        input_toolbar.pack(fill="x", padx=10, pady=(6, 4))

        lbl_hint = tk.Label(
            input_toolbar,
            text="Định dạng cột: [Mã hàng]  [Số lượng]  [Trạm]  [Zone tùy chọn] (Ngăn cách bởi Tab hoặc Khoảng trắng)\n• Ưu tiên: Zone nhập vào > zone_map.json > Gom nhóm tự động theo Trạm (vd: Hùng/Long)",
            font=("Segoe UI", 9),
            fg="#64748b",
            bg="#ffffff",
            justify="left"
        )
        lbl_hint.pack(side="left")

        btn_clear = ttk.Button(input_toolbar, text="🗑 Xóa", width=8, command=self._clear_input)
        btn_clear.pack(side="right", padx=3)

        btn_sample = ttk.Button(input_toolbar, text="📝 Nạp mẫu thử", width=14, command=self._load_sample)
        btn_sample.pack(side="right", padx=3)

        btn_paste = ttk.Button(input_toolbar, text="📋 Dán từ Clipboard", width=18, command=self._paste_clipboard)
        btn_paste.pack(side="right", padx=3)

        # Khung Textarea
        text_container = tk.Frame(top_frame, bg="#ffffff")
        text_container.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        self.txt_input = tk.Text(
            text_container,
            wrap="none",
            font=("Consolas", 10),
            bg="#f8fafc",
            fg="#0f172a",
            relief="solid",
            borderwidth=1,
            height=6
        )
        scroll_y = ttk.Scrollbar(text_container, orient="vertical", command=self.txt_input.yview)
        scroll_x = ttk.Scrollbar(text_container, orient="horizontal", command=self.txt_input.xview)
        self.txt_input.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.txt_input.pack(side="left", fill="both", expand=True)

        # Nạp dữ liệu mẫu ban đầu
        self.txt_input.insert("1.0", SAMPLE_INPUT)

        # Khung điều khiển chạy & Tiến trình
        ctrl_frame = tk.Frame(top_frame, bg="#ffffff")
        ctrl_frame.pack(fill="x", padx=10, pady=(4, 8))

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
            text="Tổng cộng: 0 dòng kết quả",
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
            padx=12,
            pady=5,
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
        canvas = tk.Canvas(self.tab_config, bg="#f8fafc", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.tab_config, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f8fafc")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content_pad = 16

        # KHỐI 1: Kết nối SAP GUI
        grp_sap = ttk.LabelFrame(scrollable_frame, text=" 1. Cấu Hình Kết Nối SAP GUI ", style="Card.TLabelframe")
        grp_sap.pack(fill="x", padx=content_pad, pady=(content_pad, 8))

        f_sap = tk.Frame(grp_sap, bg="#ffffff", padx=12, pady=10)
        f_sap.pack(fill="x")

        tk.Label(f_sap, text="System (Tên hệ thống):", bg="#ffffff").grid(row=0, column=0, sticky="w", pady=4)
        self.cfg_sap_sys = ttk.Entry(f_sap, width=15)
        self.cfg_sap_sys.insert(0, self.config.get("sap_system", "PSV"))
        self.cfg_sap_sys.grid(row=0, column=1, sticky="w", padx=6, pady=4)

        tk.Label(f_sap, text="Client (Phân hệ):", bg="#ffffff").grid(row=0, column=2, sticky="w", padx=(20, 0), pady=4)
        self.cfg_sap_client = ttk.Entry(f_sap, width=10)
        self.cfg_sap_client.insert(0, self.config.get("sap_client", "908"))
        self.cfg_sap_client.grid(row=0, column=3, sticky="w", padx=6, pady=4)

        tk.Label(f_sap, text="Session Index (-1 = tự động tìm):", bg="#ffffff").grid(row=0, column=4, sticky="w", padx=(20, 0), pady=4)
        self.cfg_sap_idx = ttk.Entry(f_sap, width=8)
        self.cfg_sap_idx.insert(0, str(self.config.get("sap_session_idx", -1)))
        self.cfg_sap_idx.grid(row=0, column=5, sticky="w", padx=6, pady=4)

        btn_test_sap = ttk.Button(f_sap, text="🔌 Thử Kết Nối", command=self._test_sap_connection)
        btn_test_sap.grid(row=0, column=6, sticky="e", padx=(20, 0), pady=4)

        # KHỐI 2: Tham số EWM & Kho MIGO
        grp_ewm = ttk.LabelFrame(scrollable_frame, text=" 2. Tham Số Kho EWM & MIGO 311 ", style="Card.TLabelframe")
        grp_ewm.pack(fill="x", padx=content_pad, pady=8)

        f_ewm = tk.Frame(grp_ewm, bg="#ffffff", padx=12, pady=10)
        f_ewm.pack(fill="x")

        tk.Label(f_ewm, text="Mã kho EWM (Warehouse):", bg="#ffffff").grid(row=0, column=0, sticky="w", pady=4)
        self.cfg_whse = ttk.Entry(f_ewm, width=12)
        self.cfg_whse.insert(0, self.config.get("ewm_warehouse", "VF02"))
        self.cfg_whse.grid(row=0, column=1, sticky="w", padx=6, pady=4)

        tk.Label(f_ewm, text="Storage Type tồn:", bg="#ffffff").grid(row=0, column=2, sticky="w", padx=(20, 0), pady=4)
        self.cfg_storage = ttk.Entry(f_ewm, width=12)
        self.cfg_storage.insert(0, self.config.get("ewm_storage_type", "APFL"))
        self.cfg_storage.grid(row=0, column=3, sticky="w", padx=6, pady=4)

        tk.Label(f_ewm, text="Chủ hàng (Owner):", bg="#ffffff").grid(row=0, column=4, sticky="w", padx=(20, 0), pady=4)
        self.cfg_owner = ttk.Entry(f_ewm, width=12)
        self.cfg_owner.insert(0, self.config.get("ewm_owner", "3093"))
        self.cfg_owner.grid(row=0, column=5, sticky="w", padx=6, pady=4)

        # Hàng 2: Plant & SLoc tách biệt
        tk.Label(f_ewm, text="Plant MIGO:", bg="#ffffff").grid(row=1, column=0, sticky="w", pady=6)
        self.cfg_plant = ttk.Entry(f_ewm, width=12)
        self.cfg_plant.insert(0, self.config.get("plant", "3093"))
        self.cfg_plant.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        tk.Label(f_ewm, text="Kho xuất (From SLoc):", bg="#ffffff").grid(row=1, column=2, sticky="w", padx=(20, 0), pady=6)
        self.cfg_sloc_from = ttk.Entry(f_ewm, width=12)
        self.cfg_sloc_from.insert(0, self.config.get("sloc_from", "1005"))
        self.cfg_sloc_from.grid(row=1, column=3, sticky="w", padx=6, pady=6)

        tk.Label(f_ewm, text="Kho nhận (To SLoc):", bg="#ffffff").grid(row=1, column=4, sticky="w", padx=(20, 0), pady=6)
        self.cfg_sloc_to = ttk.Entry(f_ewm, width=12)
        self.cfg_sloc_to.insert(0, self.config.get("sloc_to", "3001"))
        self.cfg_sloc_to.grid(row=1, column=5, sticky="w", padx=6, pady=6)

        # KHỐI 3: Quy tắc gom nhóm & Phân bổ
        grp_alloc = ttk.LabelFrame(scrollable_frame, text=" 3. Quy Tắc Gom Nhóm & Chiến Lược Phân Bổ ", style="Card.TLabelframe")
        grp_alloc.pack(fill="x", padx=content_pad, pady=8)

        f_alloc = tk.Frame(grp_alloc, bg="#ffffff", padx=12, pady=10)
        f_alloc.pack(fill="x")

        tk.Label(f_alloc, text="Số mã tối đa / OB:", bg="#ffffff").grid(row=0, column=0, sticky="w", pady=4)
        self.cfg_max_mats = ttk.Entry(f_alloc, width=8)
        self.cfg_max_mats.insert(0, str(self.config.get("max_materials_per_group", 5)))
        self.cfg_max_mats.grid(row=0, column=1, sticky="w", padx=6, pady=4)

        tk.Label(f_alloc, text="Số dòng tối đa / OB:", bg="#ffffff").grid(row=0, column=2, sticky="w", padx=(20, 0), pady=4)
        self.cfg_max_lines = ttk.Entry(f_alloc, width=8)
        self.cfg_max_lines.insert(0, str(self.config.get("max_lines_per_group", 20)))
        self.cfg_max_lines.grid(row=0, column=3, sticky="w", padx=6, pady=4)

        tk.Label(f_alloc, text="Chiến lược xuất:", bg="#ffffff").grid(row=0, column=4, sticky="w", padx=(20, 0), pady=4)
        self.cfg_strategy = ttk.Combobox(f_alloc, values=["FIFO", "FEFO"], width=8, state="readonly")
        self.cfg_strategy.set(self.config.get("allocation_strategy", "FIFO"))
        self.cfg_strategy.grid(row=0, column=5, sticky="w", padx=6, pady=4)

        # KHỐI 4: In ấn phiếu SAP F00112 & Lưu PDF
        grp_print = ttk.LabelFrame(scrollable_frame, text=" 4. In Phiếu SAP (F00112) & Xuất File PDF ", style="Card.TLabelframe")
        grp_print.pack(fill="x", padx=content_pad, pady=8)

        f_print = tk.Frame(grp_print, bg="#ffffff", padx=12, pady=10)
        f_print.pack(fill="x")

        tk.Label(f_print, text="Số bản in SAP:", bg="#ffffff").grid(row=0, column=0, sticky="w", pady=4)
        self.cfg_copies = ttk.Entry(f_print, width=8)
        self.cfg_copies.insert(0, str(self.config.get("sap_print_copies", 2)))
        self.cfg_copies.grid(row=0, column=1, sticky="w", padx=6, pady=4)

        self.cfg_auto_print = tk.BooleanVar(value=self.config.get("auto_print_pdf", True))
        chk_auto_print = tk.Checkbutton(
            f_print,
            text="Tự động gọi in và lưu PDF khi hoàn thành đơn",
            variable=self.cfg_auto_print,
            bg="#ffffff"
        )
        chk_auto_print.grid(row=0, column=2, columnspan=2, sticky="w", padx=(20, 0), pady=4)

        tk.Label(f_print, text="Thư mục lưu PDF:", bg="#ffffff").grid(row=1, column=0, sticky="w", pady=6)
        self.cfg_pdf_dir = ttk.Entry(f_print, width=40)
        self.cfg_pdf_dir.insert(0, self.config.get("pdf_output_folder", os.path.abspath("./OUTPUT_PDF")))
        self.cfg_pdf_dir.grid(row=1, column=1, columnspan=3, sticky="we", padx=6, pady=6)

        btn_browse_pdf = ttk.Button(f_print, text="📁 Chọn Thư Mục", command=self._browse_pdf_dir)
        btn_browse_pdf.grid(row=1, column=4, sticky="w", padx=6, pady=6)

        # KHỐI 5: Zone Mapping
        grp_zone = ttk.LabelFrame(scrollable_frame, text=" 5. Danh Mục Ánh Xạ Zone (zone_map.json) ", style="Card.TLabelframe")
        grp_zone.pack(fill="x", padx=content_pad, pady=8)

        f_zone = tk.Frame(grp_zone, bg="#ffffff", padx=12, pady=10)
        f_zone.pack(fill="x")

        self.lbl_zone_status = tk.Label(
            f_zone,
            text=f"Hiện có: {len(self.zone_mgr.mappings)} quy tắc ánh xạ Mã hàng / Trạm -> Zone -> Tổ trưởng.",
            bg="#ffffff",
            fg="#334155"
        )
        self.lbl_zone_status.pack(side="left")

        btn_reload_zone = ttk.Button(f_zone, text="🔄 Nạp Lại File", command=self._reload_zone_map)
        btn_reload_zone.pack(side="right", padx=3)

        btn_open_zone = ttk.Button(f_zone, text="📝 Mở File zone_map.json", command=self._open_zone_file)
        btn_open_zone.pack(side="right", padx=3)

        # NÚT LƯU CẤU HÌNH
        btn_save_cfg = tk.Button(
            scrollable_frame,
            text="💾  LƯU CẤU HÌNH HỆ THỐNG",
            font=("Segoe UI", 11, "bold"),
            bg="#15803d",
            fg="#ffffff",
            activebackground="#16a34a",
            activeforeground="#ffffff",
            relief="flat",
            padx=25,
            pady=10,
            cursor="hand2",
            command=self._save_config
        )
        btn_save_cfg.pack(pady=(12, 24))

    # ==========================================================================
    # CÁC HÀM XỬ LÝ SỰ KIỆN TAB 1
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
                self.txt_input.insert("1.0", cb_text.strip())
        except Exception as e:
            messagebox.showwarning("Clipboard", f"Không đọc được nội dung Clipboard:\n{e}")

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
        confirm_msg = f"Xác nhận thực hiện quy trình ({mode_str}) cho {len(requests)} dòng yêu cầu?\n\nHệ thống sẽ:\n1. Đọc tồn MON (APFL, Owner 3093)\n2. Phân bổ FIFO & Gom nhóm theo 4 tầng ưu tiên Zone\n3. Tạo MIGO 311 (1005 -> 3001)\n4. Tra cứu WHO cho từng Material\n5. In phiếu SAP 2 bản lưu PDF"
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
        messagebox.showerror("Thông Báo Lỗi", f"Tiến trình bị dừng do ngoại lệ:\n{err_msg}")

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
            SapConnector.init_thread_com()
            try:
                res = SapConnector.test_connection(
                    session_idx=int(self.config.get("sap_session_idx", -1)),
                    preferred_sys=self.config.get("sap_system", "PSV"),
                    preferred_client=self.config.get("sap_client", "908")
                )
                if res.get("ok"):
                    txt = f"🟢 SAP Sẵn Sàng ({res.get('system')}/{res.get('client')} - User: {res.get('user')})"
                    bg_color = "#15803d"
                else:
                    txt = "🔴 Chưa Kết Nối SAP"
                    bg_color = "#b91c1c"
                self.root.after(0, lambda: self._update_sap_badge(txt, bg_color))
            finally:
                SapConnector.uninit_thread_com()

        threading.Thread(target=check, daemon=True).start()

    def _update_sap_badge(self, text: str, bg_color: str):
        self.sap_badge_var.set(text)
        self.lbl_sap_badge.configure(bg=bg_color)

    def _test_sap_connection(self):
        SapConnector.init_thread_com()
        try:
            try:
                s_idx = int(self.cfg_sap_idx.get())
            except ValueError:
                s_idx = -1
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
        finally:
            SapConnector.uninit_thread_com()

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
