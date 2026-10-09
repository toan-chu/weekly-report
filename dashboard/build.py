#!/usr/bin/env python3
"""Đóng gói dashboard/src thành MỘT file HTML chạy offline (cùng cách với incal).

Chạy:  python dashboard/build.py
Kết quả: dashboard/Credit_Dashboard.html — tool chép file này vào thư mục gốc workspace.
"""
import base64
import mimetypes
import pathlib
import re
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

HERE = pathlib.Path(__file__).resolve().parent
SRC = HERE / "src"
OUT = HERE / "Credit_Dashboard.html"


def read(p: pathlib.Path) -> str:
    if not p.exists():
        sys.exit(f"Thiếu file: {p}")
    return p.read_text(encoding="utf-8")


def b64(p: pathlib.Path) -> str:
    return base64.b64encode(p.read_bytes()).decode("ascii")


def inline_fonts(css: str) -> str:
    def thay(m):
        p = (SRC / m.group(2)).resolve()
        mime = mimetypes.guess_type(p.name)[0] or "font/woff2"
        return f"url('data:{mime};base64,{b64(p)}')"
    css = re.sub(r"url\((['\"]?)(assets/fonts/[^)'\"]+)\1\)", thay, css)
    if "assets/fonts/" in css:
        sys.exit("Còn URL font chưa nhúng")
    return css


def main() -> None:
    html = read(SRC / "index.template.html")
    parts = {
        "__CSS__": inline_fonts(read(SRC / "fonts.css") + "\n" + read(SRC / "app.css")),
        "__XLSX__": read(SRC / "vendor" / "xlsx.full.min.js"),
        "__APP__": read(SRC / "app.js"),
        "__LOGO__": "data:image/png;base64," + b64(SRC / "assets" / "logo.png"),
    }
    for k, v in parts.items():
        token = "/*" + k + "*/"
        if token not in html:
            sys.exit(f"Khung HTML thiếu chỗ cắm {token}")
        html = html.replace(token, v)
    OUT.write_text(html, encoding="utf-8")
    print(f"Đã dựng {OUT.name} ({OUT.stat().st_size / 1024 / 1024:.2f} MB)")


if __name__ == "__main__":
    main()
