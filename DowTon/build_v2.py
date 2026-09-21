import win32com.client as win32
import os, sys, shutil

src_path = os.path.abspath(r"G:\SAP\DowTon\260920_DowTon_V1.xlsm")
dst_path = os.path.abspath(r"G:\SAP\DowTon\260920_DowTon_V2.xlsm")

print(f"Step 1: Copying {src_path} -> {dst_path}...")
shutil.copy2(src_path, dst_path)

xl = win32.Dispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False

try:
    wb = xl.Workbooks.Open(dst_path)
    print("Workbook opened successfully!")
    
    # ----------------------------------------------------
    # Helper to find or add sheet
    # ----------------------------------------------------
    def get_or_add_sheet(name, before_sheet=None, after_sheet=None):
        for s in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(s).Name.upper() == name.upper():
                return wb.Sheets(s)
        if before_sheet:
            sh = wb.Sheets.Add(Before=before_sheet)
        elif after_sheet:
            sh = wb.Sheets.Add(After=after_sheet)
        else:
            sh = wb.Sheets.Add()
        sh.Name = name
        return sh

    ws_dash = get_or_add_sheet("DASHBOARD", before_sheet=wb.Sheets(1))
    ws_full = get_or_add_sheet("Ton_Full", after_sheet=ws_dash)
    ws_do = get_or_add_sheet("KQ_DO", after_sheet=ws_full)
    ws_sto = get_or_add_sheet("KQ_STO", after_sheet=ws_do)
    
    # Ensure sheets visibility
    ws_dash.Visible = -1 # xlSheetVisible
    ws_full.Visible = -1
    ws_do.Visible = -1
    ws_sto.Visible = -1
    
    # ----------------------------------------------------
    # Fix Check tồn SUMIFS performance regression
    # ----------------------------------------------------
    print("Step 2: Optimizing Check tồn (replacing 686,868-row SUMIFS)...")
    sh_check = wb.Sheets("Check tồn")
    sh_mb = wb.Sheets("MB52")
    
    last_mb = sh_mb.Cells(sh_mb.Rows.Count, "A").End(-4162).Row
    mb_vals = sh_mb.Range(f"A2:I{last_mb}").Value
    
    dic_kho_body = {}
    dic_line_body = {}
    dic_transit = {}
    dic_kho_ma = {}
    dic_line_ma = {}
    dic_desc = {}
    
    for row in mb_vals:
        mat = str(row[0]).strip() if row[0] is not None else ""
        if not mat: continue
        
        qty = float(row[3]) if row[3] is not None else 0.0
        plant = str(row[5]).strip().replace(".0", "") if row[5] is not None else ""
        sloc = str(row[6]).strip().replace(".0", "") if row[6] is not None else ""
        transit = float(row[7]) if row[7] is not None else 0.0
        desc = str(row[8]).strip() if row[8] is not None else ""
        
        if mat not in dic_desc and desc: dic_desc[mat] = desc
        
        if plant == "3001" and sloc == "1003":
            dic_kho_body[mat] = dic_kho_body.get(mat, 0.0) + qty
        elif plant == "3001" and sloc == "3003":
            dic_line_body[mat] = dic_line_body.get(mat, 0.0) + qty
            
        dic_transit[mat] = dic_transit.get(mat, 0.0) + transit
        
        if plant == "3093" and sloc in ("3093", "3003"):
            dic_kho_ma[mat] = dic_kho_ma.get(mat, 0.0) + qty
        elif plant == "3093" and sloc in ("3000", "3001"):
            dic_line_ma[mat] = dic_line_ma.get(mat, 0.0) + qty

    last_check = sh_check.Cells(sh_check.Rows.Count, "A").End(-4162).Row
    check_mats = sh_check.Range(f"A2:B{last_check}").Value
    
    calc_rows = []
    for r in range(len(check_mats)):
        mat = str(check_mats[r][0]).strip() if check_mats[r][0] is not None else ""
        usage = float(check_mats[r][1]) if check_mats[r][1] is not None and float(check_mats[r][1] or 0) > 0 else 1.0
        
        tk_b = dic_kho_body.get(mat, 0.0)
        tl_b = dic_line_body.get(mat, 0.0)
        can_build = round((tk_b + tl_b) / usage, 1)
        it = dic_transit.get(mat, 0.0)
        tk_ma = dic_kho_ma.get(mat, 0.0)
        tl_ma = dic_line_ma.get(mat, 0.0)
        tong_ma = tk_ma + tl_ma
        
        calc_rows.append([tk_b, tl_b, can_build, it, tk_ma, tl_ma, tong_ma])
        
    sh_check.Range(f"C2:I{last_check}").Value = calc_rows
    print(f"Check tồn updated with {len(calc_rows)} rows of precomputed stock data!")

    # ----------------------------------------------------
    # Step 3: Format DASHBOARD sheet
    # ----------------------------------------------------
    print("Step 3: Formatting DASHBOARD sheet...")
    ws_dash.Cells.Clear()
    
    # Column Widths
    col_widths = {
        'A': 3, 'B': 6, 'C': 18, 'D': 28, 'E': 14,
        'F': 12, 'G': 14, 'H': 12, 'I': 12, 'J': 12,
        'K': 12, 'L': 14, 'M': 18, 'N': 24, 'O': 3
    }
    for col_letter, width in col_widths.items():
        ws_dash.Columns(col_letter).ColumnWidth = width
        
    # Row Heights
    ws_dash.Rows(1).RowHeight = 40
    ws_dash.Rows(2).RowHeight = 24
    for r in range(3, 8): ws_dash.Rows(r).RowHeight = 22
    ws_dash.Rows(8).RowHeight = 26
    ws_dash.Rows(9).RowHeight = 28

    # Title Banner (Row 1, B1:I1) and Slogan (J1:N1)
    title_rng = ws_dash.Range("B1:I1")
    title_rng.Merge()
    title_rng.Value = "MA LOGISTIC CONTROL - DASHBOARD CẢNH BÁO"
    title_rng.Interior.Color = 0x5D361B # Dark Navy BGR: 0x5D, 0x36, 0x1B -> RGB(27, 54, 93)
    title_rng.Font.Name = "Calibri"
    title_rng.Font.Size = 16
    title_rng.Font.Bold = True
    title_rng.Font.Color = 0xFFFFFF
    title_rng.HorizontalAlignment = -4131 # xlLeft
    title_rng.VerticalAlignment = -4108 # xlCenter
    ws_dash.Range("B1").IndentLevel = 1

    slogan_rng = ws_dash.Range("J1:N1")
    slogan_rng.Merge()
    slogan_rng.Value = "An toàn tồn kho - Chủ động sản xuất - Đáp ứng khách hàng"
    slogan_rng.Interior.Color = 0x5D361B
    slogan_rng.Font.Name = "Calibri"
    slogan_rng.Font.Size = 10
    slogan_rng.Font.Italic = True
    slogan_rng.Font.Color = 0xFFFFFF
    slogan_rng.HorizontalAlignment = -4152 # xlRight
    slogan_rng.VerticalAlignment = -4108 # xlCenter
    ws_dash.Range("J1").IndentLevel = 1

    # Block 1: THÔNG TIN ĐIỀU KHIỂN (B2:D7)
    hdr_ctrl = ws_dash.Range("B2:D2")
    hdr_ctrl.Merge()
    hdr_ctrl.Value = "THÔNG TIN ĐIỀU KHIỂN"
    hdr_ctrl.Interior.Color = 0xF2E1D9 # Light blue fill
    hdr_ctrl.Font.Name = "Calibri"
    hdr_ctrl.Font.Size = 11
    hdr_ctrl.Font.Bold = True
    hdr_ctrl.Font.Color = 0x794E1F # Dark blue text
    hdr_ctrl.HorizontalAlignment = -4108
    hdr_ctrl.VerticalAlignment = -4108
    
    ctrl_rows = [
        (3, "Model:", "Limo"),
        (4, "Mức tồn min:", 200),
        (5, "Lần tải MB52 gần nhất:", "18/03/2025 09:24:17"),
        (6, "Tổng mã cảnh báo:", 0),
        (7, "Trạng thái:", "Đã tải MB52")
    ]
    for r, lbl_text, val in ctrl_rows:
        ws_dash.Cells(r, "B").Value = lbl_text
        ws_dash.Cells(r, "B").Font.Bold = True
        ws_dash.Cells(r, "B").Font.Size = 10
        ws_dash.Cells(r, "B").HorizontalAlignment = -4152 # xlRight
        ws_dash.Cells(r, "B").VerticalAlignment = -4108
        
        val_rng = ws_dash.Range(f"C{r}:D{r}")
        val_rng.Merge()
        val_rng.Value = val
        val_rng.Font.Name = "Calibri"
        val_rng.Font.Size = 10
        val_rng.HorizontalAlignment = -4108
        val_rng.VerticalAlignment = -4108
        
        if r == 3: # Model
            val_rng.Font.Bold = True
            val_rng.Interior.Color = 0xF9F9F9
            val_rng.Font.Color = 0x794E1F
            val_rng.HorizontalAlignment = -4131 # xlLeft so text doesn't clash with button on right
            val_rng.IndentLevel = 1
        elif r == 4: # Min level
            val_rng.Font.Bold = True
            val_rng.Interior.Color = 0xF9F9F9
        elif r == 6: # Warning count
            val_rng.Font.Bold = True
            val_rng.Font.Color = 0x0000C0 # Red
        elif r == 7: # Status
            val_rng.Interior.Color = 0xDAEFE2 # Light Green
            val_rng.Font.Color = 0x235738 # Dark Green
            val_rng.Font.Bold = True

    # Border for Control Box
    for edge in [7, 8, 9, 10, 11, 12]: # xlEdgeLeft..
        try:
            ws_dash.Range("B2:D7").Borders(edge).LineStyle = 1
            ws_dash.Range("B2:D7").Borders(edge).Color = 0xD9D9D9
        except: pass
    ws_dash.Range("B2:D7").BorderAround(LineStyle=1, Weight=2, Color=0xBDD7EE)

    # Info Banner (E6:K7)
    info_rng = ws_dash.Range("E6:K7")
    info_rng.Merge()
    info_rng.Value = "ℹ️  Bảng cảnh báo tự động cập nhật khi thay đổi Model và Mức tồn min."
    info_rng.Interior.Color = 0xF7EBDD # Soft light blue
    info_rng.Font.Name = "Calibri"
    info_rng.Font.Size = 10
    info_rng.Font.Italic = True
    info_rng.Font.Color = 0x794E1F
    info_rng.HorizontalAlignment = -4108
    info_rng.VerticalAlignment = -4108
    info_rng.BorderAround(LineStyle=1, Weight=1, Color=0xBDD7EE)

    # Block 3: KPI Summary Box (L2:N7)
    kpi_hdr = ws_dash.Range("L2:N2")
    kpi_hdr.Merge()
    kpi_hdr.Value = "TỔNG HỢP CẢNH BÁO (Model: Limo)"
    kpi_hdr.Interior.Color = 0xF2E1D9
    kpi_hdr.Font.Name = "Calibri"
    kpi_hdr.Font.Size = 11
    kpi_hdr.Font.Bold = True
    kpi_hdr.Font.Color = 0x794E1F
    kpi_hdr.HorizontalAlignment = -4108
    kpi_hdr.VerticalAlignment = -4108

    kpi_items = [
        (3, "Mã dưới min", 0, 0xCCF2FF, 0x006080), # Yellow
        (4, "Ưu tiên kéo line", 0, 0xD6E4FC, 0x1159C6), # Light orange
        (5, "UT xuất", 0, 0xF7EBDD, 0x794E1F), # Light blue
        (6, "Hết tồn", 0, 0xADCBF8, 0x0000C0) # Coral / Light red
    ]
    for r, lbl_text, val, fill_color, text_color in kpi_items:
        lbl_rng = ws_dash.Range(f"L{r}:M{r}")
        lbl_rng.Merge()
        lbl_rng.Value = lbl_text
        lbl_rng.Font.Name = "Calibri"
        lbl_rng.Font.Size = 10
        lbl_rng.Font.Bold = True
        lbl_rng.Interior.Color = fill_color
        lbl_rng.Font.Color = text_color
        lbl_rng.HorizontalAlignment = -4131
        lbl_rng.VerticalAlignment = -4108
        lbl_rng.IndentLevel = 1
        
        v_cell = ws_dash.Cells(r, "N")
        v_cell.Value = val
        v_cell.Font.Name = "Calibri"
        v_cell.Font.Size = 11
        v_cell.Font.Bold = True
        v_cell.Interior.Color = fill_color
        v_cell.Font.Color = text_color
        v_cell.HorizontalAlignment = -4108
        v_cell.VerticalAlignment = -4108

    ws_dash.Range("L2:N6").BorderAround(LineStyle=1, Weight=2, Color=0xBDD7EE)

    # Section Title (Row 8, B8)
    sec_title = ws_dash.Range("B8:N8")
    sec_title.Value = ""
    ws_dash.Cells(8, "B").Value = "DANH SÁCH MÃ CẢNH BÁO"
    ws_dash.Cells(8, "B").Font.Name = "Calibri"
    ws_dash.Cells(8, "B").Font.Size = 12
    ws_dash.Cells(8, "B").Font.Bold = True
    ws_dash.Cells(8, "B").Font.Color = 0x794E1F
    ws_dash.Cells(8, "B").VerticalAlignment = -4108

    # Table Header (Row 9, B9:N9)
    headers = [
        ("STT", -4108),
        ("Material", -4108),
        ("Tên mã", -4131),
        ("Model", -4108),
        ("Tồn line", -4152),
        ("Tồn kho KH", -4152),
        ("Tồn MA", -4152),
        ("Intransit", -4152),
        ("DOH KH", -4152),
        ("Mức min", -4152),
        ("Mức độ", -4108),
        ("Cảnh báo", -4108),
        ("Hành động ưu tiên", -4108)
    ]
    for idx, (h_name, align) in enumerate(headers):
        col_idx = idx + 2 # B is 2
        c = ws_dash.Cells(9, col_idx)
        c.Value = h_name
        c.Interior.Color = 0x794E1F # Dark Blue RGB(31, 78, 121)
        c.Font.Name = "Calibri"
        c.Font.Size = 10
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
    ws_dash.Range("B9:N9").RowHeight = 28
    try: ws_dash.Range("B9:N9").AutoFilter()
    except: pass

    # Clear existing shapes on DASHBOARD
    print("Adding Action Buttons to DASHBOARD...")
    while ws_dash.Shapes.Count > 0:
        ws_dash.Shapes(1).Delete()

    # 1. Action Buttons in Rows 3-4 (Columns E to K)
    # Available space: from left of Col E to right of Col K (before Col L)
    left_start = ws_dash.Range("E3").Left + 2
    right_end = ws_dash.Range("L3").Left - 4
    total_width_avail = right_end - left_start
    btn_gap = 8
    btn_w = (total_width_avail - 3 * btn_gap) / 4
    btn_h = ws_dash.Range("E3:E4").Height - 6
    btn_top = ws_dash.Range("E3").Top + 3

    action_buttons = [
        {
            "name": "btnRefreshMB52",
            "text": "🔄 Refresh MB52",
            "macro": "Dashboard_Refresh_MB52",
            "fill": 0xB5752F, # Blue RGB(47, 117, 181)
        },
        {
            "name": "btnCheckDO",
            "text": "📄 Check DO",
            "macro": "Run_VL06O",
            "fill": 0x358254, # Green RGB(84, 130, 53)
        },
        {
            "name": "btnCheckSTO",
            "text": "🔄 Check STO",
            "macro": "Run_ME2N",
            "fill": 0x317DED, # Orange RGB(237, 125, 49)
        },
        {
            "name": "btnCopyCanhBao",
            "text": "📋 Copy Cảnh Báo",
            "macro": "Copy_CanhBao",
            "fill": 0xA03070, # Purple RGB(112, 48, 160)
        }
    ]

    for idx, b in enumerate(action_buttons):
        b_left = left_start + idx * (btn_w + btn_gap)
        shp = ws_dash.Shapes.AddShape(5, b_left, btn_top, btn_w, btn_h) # 5 = msoShapeRoundedRectangle
        shp.Name = b["name"]
        shp.OnAction = b["macro"]
        shp.Fill.Solid()
        shp.Fill.ForeColor.RGB = b["fill"]
        shp.Line.Visible = False
        shp.TextFrame2.TextRange.Text = b["text"]
        shp.TextFrame2.TextRange.Font.Name = "Calibri"
        shp.TextFrame2.TextRange.Font.Size = 11
        shp.TextFrame2.TextRange.Font.Bold = True
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
        shp.TextFrame2.VerticalAnchor = 3 # msoAnchorMiddle
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 # msoAlignCenter

    # 2. btnSelectModel inside Row 3 (right side of C3:D3)
    c3_rng = ws_dash.Range("C3:D3")
    sm_w = 82
    sm_h = c3_rng.Height - 4
    sm_top = c3_rng.Top + 2
    sm_left = c3_rng.Left + c3_rng.Width - sm_w - 2
    shp_sm = ws_dash.Shapes.AddShape(5, sm_left, sm_top, sm_w, sm_h)
    shp_sm.Name = "btnSelectModel"
    shp_sm.OnAction = "Show_Select_Models"
    shp_sm.Fill.Solid()
    shp_sm.Fill.ForeColor.RGB = 0x9C7141 # Steel Blue
    shp_sm.Line.Visible = False
    shp_sm.TextFrame2.TextRange.Text = "🎯 Chọn Model"
    shp_sm.TextFrame2.TextRange.Font.Name = "Calibri"
    shp_sm.TextFrame2.TextRange.Font.Size = 9.5
    shp_sm.TextFrame2.TextRange.Font.Bold = True
    shp_sm.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
    shp_sm.TextFrame2.VerticalAnchor = 3
    shp_sm.TextFrame2.TextRange.ParagraphFormat.Alignment = 2

    # ----------------------------------------------------
    # Step 4: Format Ton_Full sheet
    # ----------------------------------------------------
    print("Step 4: Formatting Ton_Full sheet...")
    ws_full.Cells.Clear()
    
    full_widths = {
        'A': 6, 'B': 16, 'C': 26, 'D': 12, 'E': 8,
        'F': 14, 'G': 14, 'H': 16, 'I': 12, 'J': 14,
        'K': 14, 'L': 14, 'M': 10, 'N': 10, 'O': 14,
        'P': 18, 'Q': 12, 'R': 16, 'S': 20
    }
    for col_l, w in full_widths.items():
        ws_full.Columns(col_l).ColumnWidth = w
        
    ws_full.Rows(1).RowHeight = 35
    ws_full.Rows(2).RowHeight = 28
    
    full_title = ws_full.Range("A1:S1")
    full_title.Merge()
    full_title.Value = "MA LOGISTIC CONTROL - BẢNG TẢI TỒN CHI TIẾT (FULL STOCK)"
    full_title.Interior.Color = 0x5D361B
    full_title.Font.Name = "Calibri"
    full_title.Font.Size = 15
    full_title.Font.Bold = True
    full_title.Font.Color = 0xFFFFFF
    full_title.HorizontalAlignment = -4131
    full_title.VerticalAlignment = -4108
    full_title.IndentLevel = 1

    full_headers = [
        "STT", "Material", "Tên mã", "Model", "Usage",
        "Tồn kho Body", "Tồn line Body", "Car can build Body", "In Transit",
        "Tồn kho MA", "Tồn line MA", "Tổng MA", "Rack", "Packaging",
        "Thứ tự vào chuyền", "TRẠM", "Mức độ", "Cảnh báo", "Hành động ưu tiên"
    ]
    for idx, h_txt in enumerate(full_headers):
        c = ws_full.Cells(2, idx + 1)
        c.Value = h_txt
        c.Interior.Color = 0x794E1F
        c.Font.Name = "Calibri"
        c.Font.Size = 10
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108
    try: ws_full.Range("A2:S2").AutoFilter()
    except: pass

    # ----------------------------------------------------
    # Step 5: Format KQ_DO sheet
    # ----------------------------------------------------
    print("Step 5: Formatting KQ_DO sheet...")
    ws_do.Cells.Clear()
    ws_do.Rows(1).RowHeight = 35
    ws_do.Rows(4).RowHeight = 26
    
    do_title = ws_do.Range("A1:L1")
    do_title.Merge()
    do_title.Value = "KẾT QUẢ KIỂM TRA DO (VL06O)"
    do_title.Interior.Color = 0x358254 # Green
    do_title.Font.Name = "Calibri"
    do_title.Font.Size = 15
    do_title.Font.Bold = True
    do_title.Font.Color = 0xFFFFFF
    do_title.HorizontalAlignment = -4131
    do_title.VerticalAlignment = -4108
    do_title.IndentLevel = 1

    ws_do.Cells(2, "A").Value = "Thời gian kiểm tra DO:"
    ws_do.Cells(2, "A").Font.Bold = True
    ws_do.Cells(2, "C").Value = "Chưa chạy kiểm tra DO"
    
    # Table 1 Headers (A4:G4)
    do_h1 = ["STT", "Product", "SLI", "Q'ty DO Open", "Q'ty Slog Filled", "Q'ty Batch+Slog Filled", "Status"]
    for idx, h in enumerate(do_h1):
        c = ws_do.Cells(4, idx + 1)
        c.Value = h
        c.Interior.Color = 0x358254
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108

    # Table 2 Headers (I4:L4)
    do_h2 = ["STT", "STO Number", "DO Qty Open", "Material"]
    for idx, h in enumerate(do_h2):
        c = ws_do.Cells(4, idx + 9)
        c.Value = h
        c.Interior.Color = 0x358254
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108

    ws_do.Columns("A").ColumnWidth = 6
    ws_do.Columns("B").ColumnWidth = 16
    ws_do.Columns("C").ColumnWidth = 12
    ws_do.Columns("D").ColumnWidth = 14
    ws_do.Columns("E").ColumnWidth = 14
    ws_do.Columns("F").ColumnWidth = 18
    ws_do.Columns("G").ColumnWidth = 16
    ws_do.Columns("H").ColumnWidth = 4
    ws_do.Columns("I").ColumnWidth = 6
    ws_do.Columns("J").ColumnWidth = 16
    ws_do.Columns("K").ColumnWidth = 14
    ws_do.Columns("L").ColumnWidth = 16

    # ----------------------------------------------------
    # Step 6: Format KQ_STO sheet
    # ----------------------------------------------------
    print("Step 6: Formatting KQ_STO sheet...")
    ws_sto.Cells.Clear()
    ws_sto.Rows(1).RowHeight = 35
    ws_sto.Rows(4).RowHeight = 26
    
    sto_title = ws_sto.Range("A1:T1")
    sto_title.Merge()
    sto_title.Value = "KẾT QUẢ KIỂM TRA STO (ME2N)"
    sto_title.Interior.Color = 0x317DED # Orange
    sto_title.Font.Name = "Calibri"
    sto_title.Font.Size = 15
    sto_title.Font.Bold = True
    sto_title.Font.Color = 0xFFFFFF
    sto_title.HorizontalAlignment = -4131
    sto_title.VerticalAlignment = -4108
    sto_title.IndentLevel = 1

    ws_sto.Cells(2, "A").Value = "Thời gian kiểm tra STO:"
    ws_sto.Cells(2, "A").Font.Bold = True
    ws_sto.Cells(2, "C").Value = "Chưa chạy kiểm tra STO"
    
    # Table 1 Headers (A4:I4)
    sto_h1 = ["STT", "Material", "SLI", "Q'ty DO Open", "Q'ty STO Open", "STO Number", "Qty Need Increase/Create", "Qty Need DO Create", "Status / Hành động"]
    for idx, h in enumerate(sto_h1):
        c = ws_sto.Cells(4, idx + 1)
        c.Value = h
        c.Interior.Color = 0x317DED
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108

    # Table 2 Headers (K4:N4)
    sto_h2 = ["STT", "Material", "STO Number", "Qty Need Push DO"]
    for idx, h in enumerate(sto_h2):
        c = ws_sto.Cells(4, idx + 11)
        c.Value = h
        c.Interior.Color = 0x317DED
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108

    # Table 3 Headers (P4:T4)
    sto_h3 = ["STT", "STO Number", "STO Open Qty", "Material", "BEDAT"]
    for idx, h in enumerate(sto_h3):
        c = ws_sto.Cells(4, idx + 16)
        c.Value = h
        c.Interior.Color = 0x317DED
        c.Font.Bold = True
        c.Font.Color = 0xFFFFFF
        c.HorizontalAlignment = -4108
        c.VerticalAlignment = -4108

    ws_sto.Columns("A").ColumnWidth = 6
    ws_sto.Columns("B").ColumnWidth = 16
    ws_sto.Columns("C").ColumnWidth = 12
    ws_sto.Columns("D").ColumnWidth = 14
    ws_sto.Columns("E").ColumnWidth = 14
    ws_sto.Columns("F").ColumnWidth = 16
    ws_sto.Columns("G").ColumnWidth = 22
    ws_sto.Columns("H").ColumnWidth = 18
    ws_sto.Columns("I").ColumnWidth = 22
    ws_sto.Columns("J").ColumnWidth = 4
    ws_sto.Columns("K").ColumnWidth = 6
    ws_sto.Columns("L").ColumnWidth = 16
    ws_sto.Columns("M").ColumnWidth = 16
    ws_sto.Columns("N").ColumnWidth = 16
    ws_sto.Columns("O").ColumnWidth = 4
    ws_sto.Columns("P").ColumnWidth = 6
    ws_sto.Columns("Q").ColumnWidth = 16
    ws_sto.Columns("R").ColumnWidth = 14
    ws_sto.Columns("S").ColumnWidth = 16
    ws_sto.Columns("T").ColumnWidth = 14

    # ----------------------------------------------------
    # Step 7: Inject VBA Modules & UserForm
    # ----------------------------------------------------
    print("Step 7: Injecting updated VBA code...")
    vb_proj = wb.VBProject
    
    # 7.1 Check if Module_Dashboard exists, add or update
    mod_dash = None
    for comp in vb_proj.VBComponents:
        if comp.Name == "Module_Dashboard":
            mod_dash = comp
            break
    if not mod_dash:
        mod_dash = vb_proj.VBComponents.Add(1) # vbext_ct_StdModule
        mod_dash.Name = "Module_Dashboard"

    mod_dash_path = os.path.join(os.path.dirname(__file__), "Module_Dashboard.bas")
    with open(mod_dash_path, "r", encoding="utf-8") as f:
        code_dashboard = f.read()

    mod_dash.CodeModule.DeleteLines(1, mod_dash.CodeModule.CountOfLines)
    mod_dash.CodeModule.AddFromString(code_dashboard)
    print("Module_Dashboard injected successfully from Module_Dashboard.bas!")

    # 7.2 Update Module1.bas (Export_MB52 calls Calculate_All_Stock)
    mod1 = None
    for comp in vb_proj.VBComponents:
        if comp.Name == "Module1":
            mod1 = comp
            break
    if mod1:
        code_mod1 = mod1.CodeModule.Lines(1, mod1.CodeModule.CountOfLines)
        if "Calculate_All_Stock" not in code_mod1:
            # Add Call Calculate_All_Stock(False) before MsgBox or CleanExit
            code_mod1 = code_mod1.replace(
                'ExportWB.Close False',
                'ExportWB.Close False\n\n    \' Auto update Dashboard and Ton_Full\n    Call Calculate_All_Stock(False)'
            )
            mod1.CodeModule.DeleteLines(1, mod1.CodeModule.CountOfLines)
            mod1.CodeModule.AddFromString(code_mod1)
            print("Module1 updated to trigger Calculate_All_Stock automatically!")

    # 7.3 Update Module2.bas (Run_VL06O outputs to KQ_DO)
    mod2 = None
    for comp in vb_proj.VBComponents:
        if comp.Name == "Module2":
            mod2 = comp
            break
    if mod2:
        code_mod2 = mod2.CodeModule.Lines(1, mod2.CodeModule.CountOfLines)
        # Check if KQ_DO is referenced
        if "KQ_DO" not in code_mod2:
            target_str = 'wsI.Range("H3").Resize(ReportIdx - 1, 6).Value = ArrReport'
            replacement_str = '''wsI.Range("H3").Resize(ReportIdx - 1, 6).Value = ArrReport

    ' =========================================
    ' XUAT KET QUA RA SHEET KQ_DO
    ' =========================================
    Dim wsDO As Worksheet
    Set wsDO = Sheets("KQ_DO")
    Dim oldLastDO As Long
    oldLastDO = wsDO.Cells(wsDO.Rows.Count, "A").End(xlUp).Row
    If oldLastDO >= 5 Then wsDO.Range("A5:L" & oldLastDO).ClearContents
    
    wsDO.Range("C2").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")
    
    If ReportIdx > 1 Then
        Dim arrDOOut() As Variant
        ReDim arrDOOut(1 To ReportIdx - 1, 1 To 7)
        Dim rDO As Long, cDO As Long
        For rDO = 1 To ReportIdx - 1
            arrDOOut(rDO, 1) = rDO
            For cDO = 1 To 6: arrDOOut(rDO, cDO + 1) = ArrReport(rDO, cDO): Next cDO
        Next rDO
        wsDO.Range("A5").Resize(ReportIdx - 1, 7).Value = arrDOOut
    End If
    
    If ArrIndex > 1 Then
        Dim arrDORawSTO() As Variant
        ReDim arrDORawSTO(1 To ArrIndex - 1, 1 To 4)
        For rDO = 1 To ArrIndex - 1
            arrDORawSTO(rDO, 1) = rDO
            For cDO = 1 To 3: arrDORawSTO(rDO, cDO + 1) = ArrSTO(rDO, cDO): Next cDO
        Next rDO
        wsDO.Range("I5").Resize(ArrIndex - 1, 4).Value = arrDORawSTO
    End If'''
            if target_str in code_mod2:
                code_mod2 = code_mod2.replace(target_str, replacement_str)
                mod2.CodeModule.DeleteLines(1, mod2.CodeModule.CountOfLines)
                mod2.CodeModule.AddFromString(code_mod2)
                print("Module2 updated to output to KQ_DO!")

    # 7.4 Update Module3.bas (Run_ME2N outputs to KQ_STO)
    mod3 = None
    for comp in vb_proj.VBComponents:
        if comp.Name == "Module3":
            mod3 = comp
            break
    if mod3:
        code_mod3 = mod3.CodeModule.Lines(1, mod3.CodeModule.CountOfLines)
        if "KQ_STO" not in code_mod3:
            target_str_me2n = 'MsgBox "ME2N Completed", vbInformation'
            replacement_str_me2n = '''    ' =========================================
    ' XUAT KET QUA RA SHEET KQ_STO
    ' =========================================
    Dim wsKqSTO As Worksheet
    Set wsKqSTO = Sheets("KQ_STO")
    Dim oldLastKqSTO As Long
    oldLastKqSTO = wsKqSTO.Cells(wsKqSTO.Rows.Count, "A").End(xlUp).Row
    If oldLastKqSTO >= 5 Then wsKqSTO.Range("A5:T" & oldLastKqSTO).ClearContents
    
    wsKqSTO.Range("C2").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")
    
    If LastReportRow >= 3 Then
        Dim arrKqSTO1() As Variant
        ReDim arrKqSTO1(1 To LastReportRow - 2, 1 To 9)
        Dim rSTO As Long
        For rSTO = 3 To LastReportRow
            arrKqSTO1(rSTO - 2, 1) = rSTO - 2
            arrKqSTO1(rSTO - 2, 2) = wsI.Cells(rSTO, "H").Value
            arrKqSTO1(rSTO - 2, 3) = wsI.Cells(rSTO, "I").Value
            arrKqSTO1(rSTO - 2, 4) = wsI.Cells(rSTO, "J").Value
            arrKqSTO1(rSTO - 2, 5) = wsI.Cells(rSTO, "N").Value
            arrKqSTO1(rSTO - 2, 6) = wsI.Cells(rSTO, "O").Value
            arrKqSTO1(rSTO - 2, 7) = wsI.Cells(rSTO, "P").Value
            arrKqSTO1(rSTO - 2, 8) = wsI.Cells(rSTO, "Q").Value
            arrKqSTO1(rSTO - 2, 9) = wsI.Cells(rSTO, "M").Value
        Next rSTO
        wsKqSTO.Range("A5").Resize(LastReportRow - 2, 9).Value = arrKqSTO1
    End If
    
    If PushIdx > 1 Then
        Dim arrKqPush() As Variant
        ReDim arrKqPush(1 To PushIdx - 1, 1 To 4)
        For rSTO = 1 To PushIdx - 1
            arrKqPush(rSTO, 1) = rSTO
            arrKqPush(rSTO, 2) = ArrPush(rSTO, 1)
            arrKqPush(rSTO, 3) = ArrPush(rSTO, 2)
            arrKqPush(rSTO, 4) = ArrPush(rSTO, 3)
        Next rSTO
        wsKqSTO.Range("K5").Resize(PushIdx - 1, 4).Value = arrKqPush
    End If
    
    If STOIdx > 1 Then
        Dim arrKqRawME2N() As Variant
        ReDim arrKqRawME2N(1 To STOIdx - 1, 1 To 5)
        For rSTO = 1 To STOIdx - 1
            arrKqRawME2N(rSTO, 1) = rSTO
            arrKqRawME2N(rSTO, 2) = ArrME2N(rSTO, 1)
            arrKqRawME2N(rSTO, 3) = ArrME2N(rSTO, 2)
            arrKqRawME2N(rSTO, 4) = ArrME2N(rSTO, 3)
            arrKqRawME2N(rSTO, 5) = ArrME2N(rSTO, 4)
        Next rSTO
        wsKqSTO.Range("P5").Resize(STOIdx - 1, 5).Value = arrKqRawME2N
        wsKqSTO.Range("T5:T" & (STOIdx + 3)).NumberFormat = "dd/mm/yyyy"
    End If

    MsgBox "ME2N Completed - K" & ChrW(7871) & "t qu" & ChrW(7843) & " " & ChrW(273) & ChrW(227) & " xu" & ChrW(7845) & "t ra sheet KQ_STO!", vbInformation'''
            if target_str_me2n in code_mod3:
                code_mod3 = code_mod3.replace(target_str_me2n, replacement_str_me2n)
                mod3.CodeModule.DeleteLines(1, mod3.CodeModule.CountOfLines)
                mod3.CodeModule.AddFromString(code_mod3)
                print("Module3 updated to output to KQ_STO!")

    # 7.5 Add frmSelectModel UserForm
    uf_exists = False
    for comp in vb_proj.VBComponents:
        if comp.Name == "frmSelectModel":
            uf_exists = True
            break
    if not uf_exists:
        print("Creating UserForm frmSelectModel...")
        uf = vb_proj.VBComponents.Add(3)
        uf.Name = "frmSelectModel"
        uf.Properties("Caption").Value = "Chọn Model (Multi-Select)"
        uf.Properties("Width").Value = 300
        uf.Properties("Height").Value = 360
        
        lbl = uf.Designer.Controls.Add("Forms.Label.1", "lblInfo")
        lbl.Caption = "Chọn một hoặc nhiều Model cần theo dõi:"
        lbl.Top = 10
        lbl.Left = 10
        lbl.Width = 270
        lbl.Height = 20
        
        lst = uf.Designer.Controls.Add("Forms.ListBox.1", "lstModels")
        lst.Top = 32
        lst.Left = 10
        lst.Width = 265
        lst.Height = 235
        lst.MultiSelect = 1 # fmMultiSelectMulti
        lst.ListStyle = 1   # fmListStyleOption (checkboxes!)
        
        btnAll = uf.Designer.Controls.Add("Forms.CommandButton.1", "btnSelectAll")
        btnAll.Caption = "Chọn Tất Cả"
        btnAll.Top = 275
        btnAll.Left = 10
        btnAll.Width = 80
        btnAll.Height = 26
        
        btnClear = uf.Designer.Controls.Add("Forms.CommandButton.1", "btnClearAll")
        btnClear.Caption = "Bỏ Chọn"
        btnClear.Top = 275
        btnClear.Left = 95
        btnClear.Width = 80
        btnClear.Height = 26
        
        btnApply = uf.Designer.Controls.Add("Forms.CommandButton.1", "btnApply")
        btnApply.Caption = "Áp Dụng"
        btnApply.Top = 275
        btnApply.Left = 185
        btnApply.Width = 90
        btnApply.Height = 26

        uf_code = """Option Explicit

Private Sub UserForm_Initialize()
    Dim wsCheck As Worksheet
    Dim lastRow As Long
    Dim dicModels As Object
    Dim r As Long, m As String
    Dim k As Variant
    Dim curVal As String
    Dim isAll As Boolean
    Dim i As Long
    Dim parts() As String, p As Variant
    Dim selDic As Object
    
    Me.Caption = "Ch" & ChrW(7885) & "n Model (Multi-Select)"
    Me.lblInfo.Caption = "Ch" & ChrW(7885) & "n m" & ChrW(7897) & "t ho" & ChrW(7863) & "c nhi" & ChrW(7873) & "u Model c" & ChrW(7847) & "n theo d" & ChrW(245) & "i:"
    Me.btnSelectAll.Caption = "Ch" & ChrW(7885) & "n T" & ChrW(7845) & "t C" & ChrW(7843)
    Me.btnClearAll.Caption = "B" & ChrW(7887) & " Ch" & ChrW(7885) & "n"
    Me.btnApply.Caption = ChrW(193) & "p D" & ChrW(7909) & "ng"
    
    Set wsCheck = GetCheckSheet(ThisWorkbook)
    lastRow = wsCheck.Cells(wsCheck.Rows.Count, "A").End(xlUp).Row
    
    Set dicModels = CreateObject("Scripting.Dictionary")
    For r = 2 To lastRow
        m = Trim(CStr(wsCheck.Cells(r, 10).Value))
        If Len(m) > 0 Then dicModels(m) = True
    Next r
    
    Me.lstModels.Clear
    For Each k In dicModels.Keys
        Me.lstModels.AddItem CStr(k)
    Next k
    
    ' Pre-select based on C3
    curVal = Trim(CStr(ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value))
    isAll = (Len(curVal) = 0 Or UCase(curVal) = "ALL" Or curVal = UniText("Tat_ca") Or curVal = "T" & ChrW(7845) & "t c" & ChrW(7843))
    
    If isAll Then
        For i = 0 To Me.lstModels.ListCount - 1
            Me.lstModels.Selected(i) = True
        Next i
    Else
        parts = Split(curVal, ",")
        Set selDic = CreateObject("Scripting.Dictionary")
        For Each p In parts
            If Len(Trim(CStr(p))) > 0 Then selDic(UCase(Trim(CStr(p)))) = True
        Next p
        
        For i = 0 To Me.lstModels.ListCount - 1
            If selDic.Exists(UCase(Me.lstModels.List(i))) Then
                Me.lstModels.Selected(i) = True
            Else
                Me.lstModels.Selected(i) = False
            End If
        Next i
    End If
End Sub

Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 0 To Me.lstModels.ListCount - 1
        Me.lstModels.Selected(i) = True
    Next i
End Sub

Private Sub btnClearAll_Click()
    Dim i As Long
    For i = 0 To Me.lstModels.ListCount - 1
        Me.lstModels.Selected(i) = False
    Next i
End Sub

Private Sub btnApply_Click()
    Dim selCount As Long, i As Long
    Dim res As String
    selCount = 0: res = ""
    
    For i = 0 To Me.lstModels.ListCount - 1
        If Me.lstModels.Selected(i) Then
            selCount = selCount + 1
            If Len(res) > 0 Then res = res & ", "
            res = res & Me.lstModels.List(i)
        End If
    Next i
    
    If selCount = 0 Or selCount = Me.lstModels.ListCount Then
        ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value = "ALL"
    Else
        ThisWorkbook.Sheets("DASHBOARD").Range("C3").Value = res
    End If
    
    Unload Me
    Call Calculate_All_Stock(False)
End Sub
"""
        uf.CodeModule.AddFromString(uf_code)
        print("frmSelectModel code injected successfully!")

    # 7.6 Inject Worksheet events for DASHBOARD
    dash_comp = None
    for comp in vb_proj.VBComponents:
        if comp.Type == 100: # Document / Worksheet
            try:
                if comp.Properties("Name").Value == ws_dash.CodeName:
                    dash_comp = comp
                    break
            except: pass
            
    if not dash_comp:
        # Fallback by matching name
        for comp in vb_proj.VBComponents:
            if comp.Type == 100 and (ws_dash.CodeName in comp.Name):
                dash_comp = comp
                break
                
    if dash_comp:
        dash_event_code = """Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    If Not Intersect(Target, Me.Range("C3:C4")) Is Nothing Then
        Application.EnableEvents = False
        Call Calculate_All_Stock(False)
        Application.EnableEvents = True
    End If
CleanExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Me.Range("C3")) Is Nothing Then
        Cancel = True
        Call Show_Select_Models
    End If
End Sub
"""
        dash_comp.CodeModule.DeleteLines(1, dash_comp.CodeModule.CountOfLines)
        dash_comp.CodeModule.AddFromString(dash_event_code)
        print(f"DASHBOARD events injected into {dash_comp.Name}!")

    # ----------------------------------------------------
    # Step 8: Run Calculate_All_Stock to populate live data
    # ----------------------------------------------------
    print("Step 8: Populating DASHBOARD and Ton_Full live data...")
    # Set default values
    ws_dash.Range("C3").Value = "Limo"
    ws_dash.Range("C4").Value = 200
    
    # Run macro via Excel COM
    xl.Run("Calculate_All_Stock", False)
    print("Calculate_All_Stock completed successfully!")

    # Verify rows in DASHBOARD and Ton_Full
    dash_rows = ws_dash.Cells(ws_dash.Rows.Count, "B").End(-4162).Row
    full_rows = ws_full.Cells(ws_full.Rows.Count, "A").End(-4162).Row
    print(f"Verification: DASHBOARD has {dash_rows - 9} warning rows. Ton_Full has {full_rows - 2} stock rows.")

    # ----------------------------------------------------
    # Step 9: Save Workbook
    # ----------------------------------------------------
    print(f"Step 9: Saving {dst_path}...")
    ws_dash.Activate()
    ws_dash.Range("A1").Select()
    wb.Save()
    print("Saved successfully!")

finally:
    wb.Close(SaveChanges=True)
    xl.Quit()

print("ALL STEPS COMPLETED SUCCESSFULLY!")
