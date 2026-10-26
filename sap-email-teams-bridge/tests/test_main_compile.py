import py_compile
import ast
from pathlib import Path


def test_main_syntax_and_compile():
    """Đảm bảo main.py luôn được compile sạch sẽ, không có lỗi SyntaxError hay IndentationError."""
    root = Path(__file__).resolve().parent.parent
    main_py = root / "main.py"
    assert main_py.exists(), f"File {main_py} does not exist"

    # 1. Compile check (doraise=True sẽ ném PyCompileError nếu có cú pháp sai)
    compiled = py_compile.compile(str(main_py), doraise=True)
    assert compiled is not None

    # 2. AST parse check
    content = main_py.read_text(encoding="utf-8")
    parsed_ast = ast.parse(content, filename=str(main_py))
    assert parsed_ast is not None

    bridge_py = root / "sap_email_teams_bridge.py"
    if bridge_py.exists():
        assert py_compile.compile(str(bridge_py), doraise=True) is not None


def test_main_no_undefined_cfg_and_valid_dispatch():
    """Kiểm tra không còn biến cfg chưa khai báo trong main.py (P1 bugfix)."""
    root = Path(__file__).resolve().parent.parent
    main_py = root / "main.py"
    content = main_py.read_text(encoding="utf-8")
    assert 'cfg.get(' not in content, "Found undefined cfg.get in main.py!"
    assert 'config.get("warehouse_slocs"' in content, "Expected config.get('warehouse_slocs' in main.py!"


def test_main_dispatch_contains_expansion_a_commands():
    """Kiểm tra main.py có đầy đủ 4 nhánh dispatch cho Mở rộng A (TIM_GR_DOC, CHECK_IB, TIM_IB, CHECK_PO)."""
    root = Path(__file__).resolve().parent.parent
    main_py = root / "main.py"
    content = main_py.read_text(encoding="utf-8")
    for cmd in ["TIM_GR_DOC", "CHECK_IB", "TIM_IB", "CHECK_PO"]:
        assert f'cmd in ["{cmd}"' in content or f'"{cmd}"' in content, f"Command {cmd} missing in main.py!"
    assert "sap.check_gr_doc_mb51" in content
    assert "sap.check_inbound_detail_vl06i" in content
    assert "sap.find_inbound_vl06i" in content
    assert "sap.check_po_me2n" in content


def test_main_dispatch_contains_warehouse_phases_commands():
    """Kiểm tra main.py có đầy đủ các nhánh dispatch cho toàn bộ 5 tool kho (BBBG, NVL, OB, ASN)."""
    root = Path(__file__).resolve().parent.parent
    main_py = root / "main.py"
    content = main_py.read_text(encoding="utf-8")
    for cmd in ["TAO_BBBG", "RA_DON_NVL", "DIEU_PHOI_OB", "TRA_CUU_ASN"]:
        assert cmd in content, f"Command {cmd} missing in main.py!"
    assert "BbbgService.reconcile" in content
    assert "NvlService.allocate_orders" in content
    assert "ObService.validate_whole_batch_gate" in content
    assert "ExtendedQueryService.track_asn_provenance" in content


def test_complete_mail_item_defensive_behavior():
    """Kiểm tra complete_mail_item không bao giờ làm sập chương trình khi UnRead hoặc Move lỗi."""
    import sys
    from unittest.mock import MagicMock
    root = Path(__file__).resolve().parent.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    import main

    # 1. Normal case
    mail_mock = MagicMock()
    bot_folder = MagicMock()
    bot_folder.Folders = {"Processed": MagicMock(), "Failed": MagicMock()}
    main.complete_mail_item(mail_mock, bot_folder, "processed")
    assert mail_mock.UnRead is False
    mail_mock.Save.assert_called_once()
    mail_mock.Move.assert_called_once()

    # 2. Case Item.UnRead lock / exception (Property 'Item.UnRead' can not be set)
    locked_mail = MagicMock()
    type(locked_mail).UnRead = property(fget=lambda self: True, fset=MagicMock(side_effect=Exception("Property 'Item.UnRead' can not be set.")))
    locked_mail.EntryID = "00000000TEST"
    ns_mock = MagicMock()
    fresh_mail = MagicMock()
    ns_mock.GetItemFromID.return_value = fresh_mail

    # Phải bắt lỗi an toàn và hoàn tất di chuyển
    main.complete_mail_item(locked_mail, bot_folder, "failed", outlook_ns=ns_mock)
    ns_mock.GetItemFromID.assert_called_once_with("00000000TEST")
    assert fresh_mail.UnRead is False
    fresh_mail.Save.assert_called_once()

    # 3. None mail or None folder
    main.complete_mail_item(None, bot_folder)
    main.complete_mail_item(locked_mail, None)


def test_start_bridge_cmd_watchdog_present():
    """Đảm bảo START_BRIDGE.cmd có cơ chế WATCHDOG_LOOP tự động phục hồi 24/7."""
    root = Path(__file__).resolve().parent.parent
    cmd_file = root / "START_BRIDGE.cmd"
    content = cmd_file.read_text(encoding="utf-8", errors="ignore")
    assert ":WATCHDOG_LOOP" in content
    assert "timeout /t 5" in content or "ping -n 6" in content
    assert "goto :WATCHDOG_LOOP" in content
