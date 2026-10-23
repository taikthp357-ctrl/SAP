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
