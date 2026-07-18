"""Statement & broker-file importers (v1.7, SPEC §6.17).

The pipeline is parse → validate → reconcile → preview → write, and every
stage may refuse. Parsers (cams.py, brokers.py) turn a user-supplied file
into a normalized ImportBatch (common.py); merge.py is the single engine
that reconciles a batch against the workbook and mutates PortfolioData.
Nothing in this package touches the filesystem except pdftext.py and the
top-level detect_format(); nothing here prompts — update.py owns all
interaction, these modules stay headless and testable.
"""


def detect_format(path, with_payload: bool = False):
    """What kind of import file this is, by CONTENT (never the name):
    'cas' for a PDF (confirmed as a statement after decryption),
    'equity_csv' / 'equity_xlsx' for a recognisable tradebook/holdings
    export, else None.

    with_payload=True returns (kind, payload) where payload is whatever
    the sniff already extracted (CSV text / XLSX rows; None for PDFs) —
    the caller hands it to the parser so a file is read once end to end.
    """
    from pathlib import Path

    from .brokers import sniff_csv
    from .pdftext import looks_like_pdf

    def _ret(kind, payload=None):
        return (kind, payload) if with_payload else kind

    p = Path(path)
    if looks_like_pdf(p):
        return _ret("cas")
    if p.suffix.casefold() in (".xlsx", ".xlsm"):
        from .brokers import rows_from_xlsx
        try:
            # payload callers get the full extraction (the openpyxl OPEN
            # dominates the cost, not the rows); sniff-only callers stop
            # at the header region
            rows = rows_from_xlsx(p) if with_payload else \
                rows_from_xlsx(p, max_rows=30)
        except Exception:                  # noqa: BLE001 — not an xlsx we
            return _ret(None)              # can read = not an import file
        if sniff_csv(rows) is not None:
            return _ret("equity_xlsx", rows)
        return _ret(None)
    # content decides, whatever the extension — a tradebook saved as .txt
    # is still a tradebook (statements are small; cap guards a stray blob)
    try:
        if p.stat().st_size > 50 * 1024 * 1024:
            return _ret(None)
        text = p.read_text(encoding="utf-8-sig", errors="replace")
    except OSError:
        return _ret(None)
    if sniff_csv(text) is not None:
        return _ret("equity_csv", text)
    return _ret(None)
