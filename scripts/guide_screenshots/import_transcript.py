"""Capture the REAL statement-import conversation for docs/USER-GUIDE.md.

Runs the actual prompt_imports() flow against a temp sample workbook and a
tiny synthetic password-protected CAS PDF (fictional folio), with the
keyboard scripted, and prints the exact transcript — so the guide's fenced
"what it looks like" block never drifts from the product. Re-run after any
prompt-wording change and paste the output into §6 of the guide.

    .venv/bin/python scripts/guide_screenshots/import_transcript.py
"""

import builtins
import contextlib
import getpass
import io
import sys
import tempfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

TODAY = date(2026, 7, 18)
PASSWORD = "demo"

CAS_LINES = [
    "Consolidated Account Statement",
    "CAMS + KFintech",
    "01-Jan-2024 To 18-Jul-2026",
    "",
    "AMIT KUMAR",
    "",
    "XYZ123 - Parag Parikh Flexi Cap Fund - Direct Plan - Growth"
    "  ISIN: INF879O01027",
    "Folio No: 12345678 / 90",
    "Opening Unit Balance: 0.000",
    "01-Jan-2024      Purchase - Systematic                 10,000.00"
    "      200.0000      50.0000       200.000",
    "05-Feb-2024      Purchase - Systematic                 10,000.00"
    "      196.0784      51.0000       396.078",
    "10-Mar-2024      Purchase - Systematic                 10,000.00"
    "      192.3077      52.0000       588.386",
    "Closing Unit Balance: 588.386",
]


def _mini_pdf(lines, path):
    stream = ("BT /F1 10 Tf 40 800 Td " + " ".join(
        f"({ln.replace('(', '[').replace(')', ']')}) Tj 0 -14 Td"
        for ln in lines) + " ET").encode()
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 842] "
        b"/Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n"
        + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out, offsets = b"%PDF-1.4\n", []
    for i, body in enumerate(objs, start=1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for off in offsets:
        out += f"{off:010d} 00000 n \n".encode()
    out += (f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\n"
            f"startxref\n{xref}\n%%EOF").encode()
    path.write_bytes(out)


def main() -> None:
    from pypdf import PdfReader, PdfWriter

    from networth.generate import build_workbook
    from networth.sample_data import sample_portfolio
    from networth.update import prompt_imports

    tmp = Path(tempfile.mkdtemp())
    wb = tmp / "Family_Portfolio_Tracker.xlsx"
    build_workbook(sample_portfolio(), str(wb), today=TODAY)

    plain = tmp / "plain.pdf"
    _mini_pdf(CAS_LINES, plain)
    writer = PdfWriter()
    for page in PdfReader(str(plain)).pages:
        writer.add_page(page)
    writer.encrypt(PASSWORD)
    with open(tmp / "CAS_2026.pdf", "wb") as fh:
        writer.write(fh)
    plain.unlink()

    answers = iter(["", "", ""])   # read it in / who owns / write confirm

    def scripted_input(prompt=""):
        ans = next(answers)
        print(prompt, end="")
        print(ans)
        return ans

    def scripted_getpass(prompt=""):
        print(prompt)               # nothing echoes for a real password
        return PASSWORD

    buf = io.StringIO()
    real_input, real_getpass = builtins.input, getpass.getpass
    builtins.input, getpass.getpass = scripted_input, scripted_getpass
    try:
        with contextlib.redirect_stdout(buf):
            prompt_imports(tmp, wb, ["Amit", "Priya", "Rahul"], [],
                           True, TODAY, workbook=wb)
    finally:
        builtins.input, getpass.getpass = real_input, real_getpass

    print(buf.getvalue().rstrip())


if __name__ == "__main__":
    main()
