# -*- coding: utf-8 -*-
"""ACME Engineering Pte Ltd - TAX INVOICE with HS code + country of origin per line
(The FIRST line above is shown on Settings > Suppliers - write the supplier and document type there.)

PERMIT SERVER READER PLUGIN - one new supplier layout, no exe rebuild needed.

HOW IT WORKS
  1. Save this file as  PERMIT_SERVER_DATA\\plugins\\<supplier>.py   (lower case, no spaces, not starting with _)
     or Settings > Suppliers > "Add plugin" on the server PC.
  2. Settings > Suppliers > Reload  (or restart the server). The list shows OK or the error.
  3. Settings > Suppliers > Test a PDF: shows what this reader makes of a PDF (nothing saved).
  4. New job with "Auto" or "Any supplier": when match() says yes, read() is used for that PDF.
     Everything after reading is the normal server: HS list, HS UOM / qty, CASC, product master, validation,
     dec-upload Excel, CHECK_REPORT, FrtImport XML, edits, undo ...

RULES FOR THE CODE
  - Only Python's standard library + modules already inside PERMIT_SERVER.exe:
      re, os, datetime, json, pdfplumber, openpyxl, permit_engine, readers, export_docs
  - match(text) must be STRICT (only this supplier's layout) and fast. Plugins are tried BEFORE the built-in layouts.
  - read() returns a list of documents made with _inv() / _line() (invoice) or R.new_pl() (packing list).
  - Never guess: leave a value empty (or None for numbers) and add a note - the server reports it.
  - Numbers as float (1,234.50 -> 1234.5 with _num). HS = 8 digits (_hs8). COO = 2-letter code (C.code('Germany') -> 'DE').
"""
import re
from export_docs import _inv, _line, _finish, _num, _hs8, _date_any, N   # helpers used by the built-in layouts
from permit_engine import clean

NAME = 'ACME'                    # short name shown in the log / Settings (upper case)


def match(text):
    """True only for this supplier's documents. text = all pages of the PDF (text layer, or OCR for scans)."""
    return 'ACME ENGINEERING PTE' in text.upper() and re.search(r'TAX\s+INVOICE', text, re.I) is not None


def read(path, pages, C):
    """path = the PDF, pages = [{'text': ..., 'words': [...]}, ...] one per page, C = country list (C.code(name)).
    Returns [invoice document] (more documents when one PDF holds several invoices / a packing list)."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'ACME invoice', 'ACME ENGINEERING PTE LTD')      # supplier name = exporter / seller
    d['title'] = 'TAX INVOICE'

    # ---- header values (one regex each; leave '' when not found)
    m = re.search(r'Invoice\s*No\.?\s*[:#]?\s*([A-Z0-9-]{5,})', t, re.I)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice\s*Date\s*[:#]?\s*(.+)', t, re.I)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''               # -> dd/mm/yyyy
    m = re.search(r'\b(EXW|FCA|FOB|CFR|CIF|CPT|CIP|DAP|DDP)\b', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'(?:UEN|Co\.?\s*Reg\.?\s*No)\s*[:.]?\s*(\d{8,9}[A-Z])', t, re.I)
    d['uen'] = m.group(1) if m else ''                                   # Singapore UEN of the exporter / importer
    m = re.search(r'Ship\s*To\s*:?\s*\n(.+)', t, re.I)
    d['ship_to'] = clean(m.group(1)).upper() if m else ''
    d['ship_to_country'] = C.find_in(t[t.upper().find('SHIP TO'):][:400]) if 'SHIP TO' in t.upper() else ''
    m = re.search(r'(?:HAWB|HBL|House\s*B/?L)\s*(?:No\.?)?\s*[:#]?\s*([A-Z0-9]{6,})', t, re.I)
    d['hawb'] = m.group(1) if m else ''
    m = re.search(r'Gross\s*Weight\s*[:]?\s*(' + N + r')\s*KG', t, re.I)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'(\d+)\s*(CARTONS?|CTNS?|PALLETS?|PKGS?|PACKAGES?)\b', t, re.I)
    if m:
        d['boxes'], d['package_unit'] = int(m.group(1)), m.group(2).upper()
    m = re.search(r'Total\s*(?:Amount|Value)\s*[:]?\s*([A-Z]{3})?\s*(' + N + r')', t, re.I)
    if m:
        d['currency'] = m.group(1) or ''
        d['total'] = _num(m.group(2))

    # ---- item lines. Example line:
    #   1  AB-1001  HYDRAULIC PUMP 24V   84137099  DE  10 PCE  125.00  1,250.00  12.500
    LN = re.compile(r'^\s*(\d{1,3})\s+(\S+)\s+(.+?)\s+(\d{6,10})\s+([A-Z]{2})\s+(' + N + r')\s*([A-Z]{2,3})\s+(' + N + r')\s+('
                    + N + r')(?:\s+(' + N + r'))?\s*$')
    for raw in t.splitlines():
        m = LN.match(raw)
        if not m:
            continue
        d['lines'].append(_line(line=m.group(1), item=m.group(2), desc=clean(m.group(3)), hs=_hs8(m.group(4)),
                                coo=m.group(5), qty=_num(m.group(6)), uom=m.group(7), price=_num(m.group(8)),
                                amount=_num(m.group(9)), nw=_num(m.group(10)) if m.group(10) else None))
    if not d['currency']:
        m = re.search(r'\b(USD|EUR|SGD|GBP|JPY|CNY|AUD)\b', t)
        d['currency'] = m.group(1) if m else ''
    if not d['incoterm']:
        d['notes'].append('incoterm not printed - check')
    return [_finish(d, C)]         # _finish: line numbers, currency per line, total check, missing HS / weight notes
