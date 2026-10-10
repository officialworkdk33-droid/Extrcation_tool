# -*- coding: utf-8 -*-
"""Built-in readers for known supplier layouts (exports AND imports) - recognised automatically:

  UD Trucks Singapore     "I N V O I C E" parts invoice + PACKING SUMMARY         (SG -> e.g. Philippines)
  Kalmar / Bromma         TAX INVOICE "No.1526014343", Commodity Code / Net Weight / Country of Origin per line
  Getinge Singapore       delivery note + Commercial Invoice (scanned - OCR), SG HS Code per line
  Cummins Asia Pacific    TAX INVOICE, HARMONIZED TARIFF # + COO per line, box table (gross / net)
  L Beauty (Luxasia)      TAX INVOICE, Product code / UPC / Cty of Origin, 10-digit HS line
  BMW Group parts invoice "For Customs Purposes only", HU pallets, COO + tariff per item (DE -> AU)
  Perkins Engines AP      TAX INVOICE SUMMARY, country of origin / tariff summary lines
  Legrand Singapore       Tax Invoice, Commercial Code / Commodity Code / Country of origin per line
  Dalian Nurmi Hydraulics scanned pack: summary page + invoice per order + delivery note (import to Sandvik SG)

Every reader returns the common document shape of readers.py (new_inv / new_line), so HS list, CASC, weights,
consolidation, Excel and XML work the same as for every other supplier."""
import os, re, datetime, collections
import permit_engine as E
from permit_engine import clean, fnum, NUM, Countries

N = r'-?(?:\d{1,3}(?:,\d{3})+|\d*)(?:\.\d+)?'


def _num(s):
    s = (s or '').strip().replace(' ', '')
    if not s or not re.search(r'\d', s):
        return None
    try:
        return float(s.replace(',', ''))
    except ValueError:
        return None


def _hs8(s):
    d = re.sub(r'\D', '', s or '')
    return d[:8] if len(d) >= 8 else d


def _dmy(d, m, y):
    y = int(y)
    y = y + 2000 if y < 100 else y
    return '%02d/%02d/%04d' % (int(d), int(m), y)


MONTHS = {m: i for i, m in enumerate(('JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'), 1)}


def _date_any(s):
    s = s or ''
    m = re.search(r'\b(\d{1,2})[./](\d{1,2})[./](\d{4})\b', s)
    if m:
        return _dmy(m.group(1), m.group(2), m.group(3))
    m = re.search(r'\b(\d{4})[/-](\d{1,2})[/-](\d{1,2})\b', s)
    if m:
        return _dmy(m.group(3), m.group(2), m.group(1))
    m = re.search(r'\b([A-Za-z]{3})[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})\b', s)
    if m and m.group(1).upper() in MONTHS:
        return _dmy(m.group(2), MONTHS[m.group(1).upper()], m.group(3))
    return ''


def _inv(path, reader, supplier):
    import readers as R
    d = R.new_inv(path, reader)
    d.update(supplier=supplier, invoice_from=supplier, ship_from=supplier, ship_from_country='SG', title='COMMERCIAL INVOICE')
    return d


def _line(**k):
    import readers as R
    return R.new_line(**k)


def _pdf_text(path, x_tol=3, crop_x0=None):
    """Text of every page again from the PDF (other spacing / without a side column)."""
    import pdfplumber
    out = []
    with pdfplumber.open(path) as pdf:
        for p in pdf.pages:
            if crop_x0:
                p = p.crop((crop_x0, 0, p.width, p.height))
            out.append(p.extract_text(x_tolerance=x_tol) or '')
    return out


def _finish(d, C, dest_name=''):
    if dest_name and not d.get('ship_to_country'):
        d['ship_to_country'] = C.code(dest_name) or ''
    if not d.get('direction'):
        if d.get('ship_from_country') == 'SG' and d.get('ship_to_country') and d['ship_to_country'] != 'SG':
            d['direction'] = 'EXPORT'
        elif d.get('ship_to_country') == 'SG':
            d['direction'] = 'IMPORT'
        else:
            d['direction'] = 'EXPORT'
    for i, l in enumerate(d['lines'], 1):
        l['line'] = l['line'] or str(i)
        l['currency'] = l['currency'] or d['currency']
        if l['qty'] is None and l['price'] and l['amount']:
            l['qty'] = round(l['amount'] / l['price'], 4)
    s = round(sum(l['amount'] or 0 for l in d['lines']), 2)
    if d.get('total') is None:
        d['total'] = s
    elif abs(s - d['total']) > max(0.05, abs(d['total']) * 0.0005):
        d['notes'].append('item lines add up to %s but the invoice goods value is %s - check' % (
            '{:,.2f}'.format(s), '{:,.2f}'.format(d['total'])))
    if not d['lines']:
        d['notes'].append('no item lines could be read')
    if not any(l.get('nw') for l in d['lines']) and not d.get('total_nw'):
        d['notes'].append('no net weight on the invoice - type the gross weight on tab 2 / shipments (net = 85 %)')
    nohs = [l['item'] for l in d['lines'] if not l['hs']]
    if nohs:
        d['notes'].append('no HS / tariff code on the invoice for: %s' % ', '.join(nohs[:15]))
    return d


# ======================================================================== UD Trucks Singapore
def ud_trucks(path, pages, C):
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'UD Trucks invoice', 'UD TRUCKS SINGAPORE (PTE.) LTD.')
    m = re.search(r'Inv No\..*\n.*?((?:\d ){5,}\d)\s+([A-Z]{3})\s*$', t, re.M)
    if m:
        dig = m.group(1).replace(' ', '')
        d['invoice_no'] = dig[:-1] if len(dig) > 6 else dig          # last digit = page number
        d['incoterm'] = m.group(2)
    m = re.search(r'\b(\d{2})-(\d{2})-(\d{2})\s+\d+\s+DAYS', t)
    if m:                                                          # yy-mm-dd
        d['invoice_date'] = _dmy(m.group(3), m.group(2), m.group(1))
    d['currency'] = 'USD' if '(USD)' in t else ''
    # consignee: letters printed with spaces ("C I V I C   M E R C H ...") - rebuild from the characters
    name = ''
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            pg = pdf.pages[0]
            ws = pg.extract_words()
            hd = next((w for w in ws if w['text'] == 'Payer'), None)
            if hd:
                row = [c for c in pg.chars if hd['top'] + 4 < c['top'] < hd['top'] + 16]
                row.sort(key=lambda c: c['x0'])
                s = ''.join(c['text'] for c in row)
                s = re.sub(r' {2,}', '|', s).replace(' ', '').strip('|')
                parts = [x for x in s.split('|') if x]
                words = []
                for x in parts:
                    if re.fullmatch(r'\d+[A-Z]*', x):
                        break
                    words.append(x)
                half = len(words) // 2
                name = ' '.join(words[:half] if half and words[:half] == words[half:2 * half] else words)
    except Exception:
        pass
    d['ship_to'] = d['invoice_to'] = clean(name)
    m = re.search(r'^(PHILIPPINES|[A-Z]{4,}(?: [A-Z]{4,})?)\s+\1\s*$', t, re.M)
    country = m.group(1) if m else ''
    d['ship_to_country'] = C.code(country) or ''
    m = re.search(r'Co\.\s*Reg\.\s*No\.\s*:\s*(\d{8,9}[A-Z])', t)
    d['uen'] = m.group(1) if m else ''
    case = ''
    LN = re.compile(r'^(\d{6}) (\d{6}) (\d{6}) (\S{1,3}) (\S+) (.+?) (\d+) (' + N + r') (' + N + r') (\d) ([A-Z]{2}) (\S+)'
                    r'(?: (\d{8,10}))? (' + N + r')\s*$')
    for p in pages:
        for raw in p['text'].splitlines():
            ln = raw.strip()
            mc = re.match(r'PROD\.NO:\s*(\d+)\s+CASE NO:\s*(\d+)', ln)
            if mc:
                case = '%s/%s' % (mc.group(1), mc.group(2))
                continue
            m = LN.match(ln)
            if m:
                d['lines'].append(_line(item=m.group(5), desc=clean(m.group(6)), qty=fnum(m.group(7)), uom='PCE',
                                        price=_num(m.group(8)), amount=_num(m.group(9)), coo=m.group(11),
                                        hs=_hs8(m.group(13) or ''), nw=_num(m.group(14)), order=m.group(1),
                                        boxes=[case] if case else [], currency=d['currency']))
    m = re.search(r'SUB TOTAL\s*:\s*(' + N + ')', t)
    d['total'] = _num(m.group(1)) if m else None
    ch = []
    for lab, rx in (('freight', r'FREIGHT.*?\s(\d[\d,]*\.\d{2})\s*$'), ('insurance', r'INSURANCE\s+(\d[\d,]*\.\d{2}|\.\d{2})\s*$')):
        mm = re.search(rx, t, re.M)
        if mm and _num(mm.group(1)):
            ch.append('%s %s' % (lab, mm.group(1)))
    m = re.search(r'INVOICE INCL\.GST\s+(' + N + ')', t)
    if ch:
        d['notes'].append('%s on the invoice (%s) - not in the item values; invoice incl. charges %s' % (
            ' + '.join(ch), d['incoterm'], m.group(1) if m else '?'))
    m = re.search(r'TOTAL\s*:\s*(\d+)\s+PACKAGE\s+(' + N + r')\s+(' + N + r')', t)
    if m:
        d['boxes'], d['total_gw'], d['total_nw'] = int(m.group(1)), _num(m.group(2)), _num(m.group(3))
        d['package_unit'] = 'PKG'
    return _finish(d, C)


# ======================================================================== Kalmar / Bromma
def kalmar(path, pages, C):
    tx = _pdf_text(path, x_tol=1)
    t = '\n'.join(tx)
    d = _inv(path, 'Kalmar / Bromma invoice', 'KALMAR PTE. LTD.')
    d['title'] = 'TAX INVOICE'
    m = re.search(r'No\.\s*(\d{8,})', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice Date\s*\n(?:.*\n){0,3}?.*?(\d{2}\.\d{2}\.\d{4})', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'Delivery Terms Document Currency\s*\n.*?\b(EXW|FCA|FOB|CFR|CIF|CPT|CIP|DAP|DPU|DDP)\b(.*?)\b([A-Z]{3})\s*$', t, re.M)
    if m:
        d['incoterm'], d['currency'] = m.group(1), m.group(3)
    m = re.search(r'Delivery Address.*\n(.+?)\s+\d+ Days', t)
    d['ship_to'] = d['invoice_to'] = clean(m.group(1)) if m else ''
    da = t[t.find('Delivery Address'):t.find('Customer Billing Contact')] if 'Delivery Address' in t else ''
    d['ship_to_country'] = C.find_in(da) or ''
    m = re.search(r'Company\s*Reg\.?\s*No\.?\s*(\d{8,9}[A-Z])', t)
    d['uen'] = m.group(1) if m else ''
    m = re.search(r'Shipment Number\s*\n\s*(\d+)', t)
    d['transport'] = m.group(1) if m else ''
    m = re.search(r'Forwarding Agent\s*\n.*?\s(\S.*?)$', t, re.M)
    d['carrier'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'Gross Weight\s*/\s*Total Volume.*?(' + N + r')\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'NO OF PACKAGES\s+(\d+)', t)
    d['boxes'] = int(m.group(1)) if m else None
    cur = None
    for ln in t.splitlines():
        ln = ln.strip()
        m = re.match(r'^(\d{1,4}) (\S+) (.+?) (' + N + r') ([A-Z]{2,4}) (' + N + r') (' + N + r')$', ln)
        if m and not ln.startswith('Line'):
            cur = _line(line=m.group(1), item=m.group(2), desc=clean(m.group(3)), qty=_num(m.group(4)), uom=m.group(5),
                        price=_num(m.group(6)), amount=_num(m.group(7)), currency=d['currency'])
            d['lines'].append(cur)
            continue
        if cur is None:
            continue
        m = re.match(r'Commodity Code\s+(\d{6,10})', ln)
        if m:
            cur['hs'] = _hs8(m.group(1))
        m = re.match(r'Net Weight\s+(' + N + r')\s*KG(?:\s*/\s*(\w+))?', ln)
        if m:
            w = _num(m.group(1))
            cur['nw'] = round(w * (cur['qty'] or 1), 4) if m.group(2) else w     # "0.28 KG / PCS" = per piece
        m = re.match(r'Country of Origin\s+([A-Z]{2})\b', ln)
        if m:
            cur['coo'] = m.group(1)
    m = re.search(r'Freight\s*Cost\s+(' + N + ')', t)
    fr = _num(m.group(1)) if m else None
    m = re.search(r'Total\s*Value\s+(' + N + ')', t)
    tot = _num(m.group(1)) if m else None
    d['total'] = round(tot - (fr or 0), 2) if tot is not None else None
    if fr:
        d['notes'].append('Freight cost %s on the invoice (%s) - total value %s incl. freight; items carry the goods value'
                          % (fr, d['incoterm'], tot))
    return _finish(d, C)


# ======================================================================== Getinge (scanned)
def getinge(path, pages, C):
    ci = [p for p in pages if re.search(r'Commercial\s+Invoice', p['text'], re.I)]
    dn = [p for p in pages if p not in ci and re.search(r'Delivery', p['text'], re.I)]
    t = '\n'.join(p['text'] for p in ci) or '\n'.join(p['text'] for p in pages)
    full = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Getinge commercial invoice', 'GETINGE SINGAPORE PTE. LTD.')
    d['ocr'] = any(p.get('ocr') for p in pages)
    m = re.search(r'\b(56\d{8})\s+(\d{4}/\d{2}/\d{2})', t) or re.search(r'\b(\d{10})\s*/?\s*(\d{4}/\d{2}/\d{2})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _date_any(m.group(2))
    m = re.search(r'Delivery note no\.?:?\s*(\d+)', t)
    d['delivery'] = m.group(1) if m else ''
    m = re.search(r'Terms of delivery:?\s*([A-Z]{3})\b\s*(.*)', t)
    if m:
        d['incoterm'], d['incoterm_place'] = m.group(1), clean(m.group(2))
    m = re.search(r'(?:Unit price|Total price)[^\n]*\n[^\n]*?\b([A-Z]{3})\s+([A-Z]{3})\b', t)
    d['currency'] = m.group(2) if m else ('JPY' if 'JPY' in t else '')
    m = re.search(r'GETINGE SINGAPORE PTE\. LTD\.\s*\n?(?:.*\n)?\s*(.+?K\.?K\.?)\b', t)
    d['ship_to'] = d['invoice_to'] = clean(m.group(1)).upper() if m else ''
    d['ship_to_country'] = 'JP' if re.search(r'\bJAPAN\b', t, re.I) else C.find_in(t[:800]) or ''
    m = re.search(r'UEN:?\s*(\d{8,9}[A-Z])', full) or re.search(r'Business Reg\. No\.:?\s*(\d{8,9}[A-Z])', full)
    d['uen'] = m.group(1) if m else ''
    m = re.search(r'Gross weight\s+(' + N + r')\s*KG', full, re.I)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'Quantity of packages:?\s*(\d+)', full)
    d['boxes'] = int(m.group(1)) if m else None
    cur = None
    for p in ci:
        for raw in p['text'].splitlines():
            ln = raw.strip()
            m = re.match(r'^(\d{2,4})\s+([0-9A-Z]{6,})\s+(\d+)\s*([A-Z]{2,3})\s*$', ln)
            if m:
                cur = _line(line=m.group(1), item=m.group(2), qty=fnum(m.group(3)), uom=m.group(4), currency=d['currency'])
                cur['_need_desc'] = True
                d['lines'].append(cur)
                continue
            if cur is None:
                continue
            if cur.get('_need_desc') and ln and not re.match(r'(Item net|Country|SG HS)', ln, re.I):
                cur['desc'] = clean(ln)
                cur['_need_desc'] = False
                continue
            m = re.match(r'Item net weight\s+(' + N + r')\s*KG', ln, re.I)
            if m:
                cur['nw'] = _num(m.group(1))
            m = re.match(r'Country of origin\s+([A-Z]{2})\b', ln, re.I)
            if m:
                cur['coo'] = m.group(1).upper()
            m = re.match(r'SG HS Code:?\s+(\d{8})', ln, re.I)
            if m:
                cur['hs'] = m.group(1)
            m = re.match(r'^(' + N + r')\s+(' + N + r')$', ln)
            if m and cur['amount'] is None:
                cur['price'], cur['amount'] = _num(m.group(1)), _num(m.group(2))
    for l in d['lines']:
        l.pop('_need_desc', None)
        if l['price'] and l['amount'] and l['qty'] and abs(l['price'] * l['qty'] - l['amount']) > 1:
            d['notes'].append('Item %s: qty x price (%s x %s) is not the line total %s - OCR, check' % (
                l['item'], l['qty'], l['price'], l['amount']))
    m = re.search(r'[Il1]tem sum\s+(' + N + ')', t)
    d['total'] = _num(m.group(1)) if m else None
    for hs in re.findall(r'SG HS Code\s*:\s*(\d{8})', t):
        if not any(l['hs'] == hs for l in d['lines']):
            d['notes'].append('HS summary at the end shows %s, which no item line has (OCR?) - check the HS codes' % hs)
    if d['ocr']:
        d['notes_ocr'] = 'scanned invoice read by OCR - check item no / HS / qty / amounts'
    return _finish(d, C)


# ======================================================================== Cummins Asia Pacific
def cummins(path, pages, C):
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Cummins invoice', 'CUMMINS ASIA PACIFIC PTE. LTD.')
    d['title'] = 'TAX INVOICE'
    m = re.search(r'INVOICE NO:\s*(\d+)', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice Date:\s*(\S+)', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'CURR\s*:\s*([A-Z]{3})', t)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'SHIP TO:.*\n(.+?)\s+Invoice Date', t)
    if m:
        s = m.group(1)
        half = s.find(' CUMMINS', 3)
        d['ship_to'] = d['invoice_to'] = clean(s[half:] if half > 0 else s).upper()
    m = re.search(r'Invoice Date.*\n(?:.*\n){1,3}?\s*([A-Z][A-Z ]+?)\s+\1\s*$', t, re.M)
    d['ship_to_country'] = C.code(m.group(1)) if m else C.find_in(t[t.find('SHIP TO'):t.find('SOLD TO')]) or ''
    m = re.search(r'BUSINESS REG NO:\s*(\d{8,9}[A-Z])', t)
    d['uen'] = m.group(1) if m else ''
    m = re.search(r'Shipped Via\s*:\s*(.+?)\s+Payment', t)
    d['carrier'] = clean(m.group(1)) if m else ''
    if re.search(r'EX-?WORKS', t):
        d['incoterm'] = 'EXW'
    LN = re.compile(r'^(\d{1,3}) (\d{6,}) (\d{4,}) (\S+) (\S+) (\S+) (\S+) (OUT|EXC|NEW) (.+?) (\d+) (' + N + r') (' + N + r')$')
    cur = None
    for p in pages:
        for raw in p['text'].splitlines():
            ln = raw.strip()
            m = LN.match(ln)
            if m:
                cur = _line(line='%s-%s' % (m.group(2), m.group(1)), item=m.group(4), desc=clean(m.group(9)), qty=fnum(m.group(10)), uom='EA',
                            price=_num(m.group(11)), amount=_num(m.group(12)), order=m.group(3), boxes=[m.group(6)],
                            currency=d['currency'])
                cur['po'] = m.group(2)
                d['lines'].append(cur)
                continue
            m = re.match(r'HARMONIZED TARIFF #\s*(\d{6,10})\s+([A-Z]{2})\b', ln)
            if m and cur is not None:
                cur['hs'], cur['coo'] = _hs8(m.group(1)), m.group(2)
                continue
            if cur is not None and (ln.startswith('Order Sub-Total') or ln.startswith('TOTAL PIECES')):
                cur = None
    m = re.search(r'Invoice Sub-Total\s*:\s*(' + N + ')', t)
    d['total'] = _num(m.group(1)) if m else None
    m = re.search(r'GROSS WEIGHT\s*:\s*(' + N + r')\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    boxes = re.findall(r'^\d+\s+(G\d+)\s+.*?(' + N + r')\s*KG\s+(' + N + r')\s*KG', t, re.M)
    if boxes:
        d['boxes'] = len(boxes)
        d['total_nw'] = round(sum(_num(b[2]) or 0 for b in boxes), 3)
        d['box_weights'] = {b[0]: (_num(b[1]), _num(b[2])) for b in boxes}
    return _finish(d, C)


# ======================================================================== L Beauty (Luxasia)
def lbeauty(path, pages, C):
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'L Beauty invoice', 'L BEAUTY PTE LTD')
    d['title'] = 'TAX INVOICE'
    m = re.search(r'Invoice No\.\s*(\d+)', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice Date\s+(\S+)', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'Currency\s+([A-Z]{3})', t)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'Incoterms\s+([A-Z]{3})', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'Customer Code:.*\n(.+?)\s+Delivery Date', t)
    name = clean(m.group(1)) if m else ''
    m2 = re.search(r'Delivery Date.*\n(\S.*?)\s+PO No', t)
    if m2 and not re.search(r'\d', m2.group(1)):
        name += ' ' + clean(m2.group(1))
    d['ship_to'] = d['invoice_to'] = name.upper()
    d['ship_to_country'] = C.find_in(t[t.find('Bill to'):t.find('Bank Details')] if 'Bill to' in t else t) or ''
    m = re.search(r'CO Reg No\.?:?\s*(\d{8,9}[A-Z])', t)
    d['uen'] = m.group(1) if m else ''
    m = re.search(r'Shipment Mode\s+(?:by\s+)?(\w+)', t)
    d['ship_mode'] = m.group(1) if m else ''
    cur = None
    for raw in t.splitlines():
        ln = raw.strip()
        m = re.match(r'^(\d{1,4}) (\S+) (\d{8,14}) ([A-Za-z][A-Za-z ]+?) (' + N + r') ([A-Z]{2,3}) (' + N + r')(?: (' + N + r'))? (' + N + r')$', ln)
        if m:
            cur = _line(line=m.group(1), item=m.group(2), qty=_num(m.group(5)), uom=m.group(6), price=_num(m.group(7)),
                        amount=_num(m.group(9)), coo=C.code(m.group(4)) or '', currency=d['currency'])
            cur['upc'] = m.group(3)
            cur['_n'] = 0
            d['lines'].append(cur)
            continue
        if cur is None:
            continue
        if re.fullmatch(r'\d{8,10}', ln):
            cur['hs'] = _hs8(ln)
            cur = None
            continue
        if not cur['desc']:
            cur['desc'] = clean(ln)
    for l in d['lines']:
        l.pop('_n', None)
    m = re.search(r'Total\s+[A-Z]{3}\s+Excl\.?\s*GST\s+(' + N + ')', t)
    d['total'] = _num(m.group(1)) if m else None
    return _finish(d, C)


# ======================================================================== BMW Group parts invoice
def bmw(path, pages, C):
    tx = _pdf_text(path, crop_x0=101)                  # left column = BMW address block
    t = '\n'.join(tx)
    d = _inv(path, 'BMW parts invoice', 'BAYERISCHE MOTOREN WERKE AG')
    d['ship_from_country'] = 'DE'
    m = re.search(r'Invoice No\..*\n\s*(\d{8,})\s+(\d{2}\.\d{2}\.\d{4})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _date_any(m.group(2))
    m = re.search(r'Total Price per\s+\S+\s*\n?.*?\(([A-Z]{3})\)', t) or re.search(r'Invoice Amount in ([A-Z]{3})', t)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'Incoterms\s+([A-Z]{3})\s*(.*)', t)
    if m:
        d['incoterm'], d['incoterm_place'] = m.group(1), clean(m.group(2))
    m = re.search(r'Shipment number\s+(\d+)', t)
    d['transport'] = m.group(1) if m else ''
    m = re.search(r'Transport Mode\s+(\w+)', t)
    d['ship_mode'] = m.group(1) if m else ''
    m = re.search(r'^(BMW [A-Z][A-Za-z]+ Ltd\.?)', t, re.M)
    d['ship_to'] = d['invoice_to'] = (m.group(1) if m else '').upper()
    d['ship_to_country'] = C.find_in(t[:1500], exclude=('DE',)) or ''
    inv_pages = [x for x in tx if 'Parts invoice' in x]
    cur, hu = None, None
    LN = re.compile(r'(\d{5}) ([A-Z0-9]{2}) (\w{11}) (.+?) (\d+) (\d{2}) (' + N + r') (' + N + r') (' + N + r') %')
    for pg in inv_pages:
        for raw in pg.splitlines():
            ln = raw.strip()
            mh = re.search(r'HU ID:\s*(\d+)\s+(\w+).*?Gross Weight:\s*(' + N + r')\s*KG\s+Load\. Weight:\s*(' + N + r')\s*KG', ln)
            if mh:
                hu = dict(id=mh.group(1), gw=_num(mh.group(3)), nw=_num(mh.group(4)))
                continue
            m = LN.search(ln)
            if m:
                cur = _line(line=m.group(1).lstrip('0') or '0', item=m.group(3), desc=clean(m.group(4)),
                            qty=fnum(m.group(5)), uom='PCE', price=_num(m.group(7)), amount=_num(m.group(8)),
                            currency=d['currency'], boxes=[hu['id']] if hu else [])
                cur['_hu'], cur['_k'] = hu, 0
                d['lines'].append(cur)
                continue
            if cur is None:
                continue
            m = re.search(r'(?:^|\s)([A-Z]{2}) (\d{8,11})$', ln)
            if m and not cur['hs']:
                cur['coo'], cur['hs'] = m.group(1), _hs8(m.group(2))
                continue
            if cur['_k'] == 0 and ln and not re.match(r'(HU ID|Rush Order|[A-Z]\d{6}\b)', ln):
                cur['desc'] = clean(cur['desc'] + ' ' + ln)
                cur['_k'] = 1
    # pallet weights -> lines (one line per pallet here; several lines on one pallet share it by value)
    by_hu = {}
    for l in d['lines']:
        if l.get('_hu'):
            by_hu.setdefault(l['_hu']['id'], []).append(l)
    for hid, ls in by_hu.items():
        h = ls[0]['_hu']
        tot = sum(x['amount'] or 0 for x in ls) or 1
        for x in ls:
            x['nw'] = round(h['nw'] * (x['amount'] or 0) / tot, 3) if h['nw'] else None
            x['gw'] = round(h['gw'] * (x['amount'] or 0) / tot, 3) if h['gw'] else None
    for l in d['lines']:
        l.pop('_hu', None)
        l.pop('_k', None)
    m = re.search(r'Invoice Amount in [A-Z]{3}\s*\n\s*(' + N + r')\s+(' + N + r')\s+(' + N + ')', t)
    d['total'] = _num(m.group(3)) if m else None
    m = re.search(r'Total Net Weight\s+(' + N + r')\s*KG', t)
    d['total_nw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total HU gross weight\s+(' + N + r')\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'^(\d+) Packages\s', t, re.M)
    d['boxes'] = int(m.group(1)) if m else len(by_hu) or None
    if d['ship_to_country'] != 'SG':
        d['notes'].append('BMW AG (DE) to %s - Singapore is not the shipper or the consignee: transhipment / '
                          're-export? Check direction and permit type on tab 3' % (d['ship_to_country'] or '?'))
    return _finish(d, C)


# ======================================================================== Perkins Engines Asia Pacific
def perkins(path, pages, C):
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Perkins invoice summary', 'PERKINS ENGINES (ASIA PACIFIC) PTE LTD')
    d['title'] = 'TAX INVOICE'
    m = re.search(r'Invoice Number:\s*(\d+)', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice Date:\s*(\S+)', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'Shipment Number:\s*(\d+)', t)
    d['transport'] = m.group(1) if m else ''
    m = re.search(r'Ship To.*\n\s*(.+?)\s+Singapore \d{6}', t)
    d['ship_to'] = d['invoice_to'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'Shipment Number.*', t)
    d['ship_to_country'] = C.find_in(t[:t.find('Shipment Number')] if 'Shipment Number' in t else t[:900]) or ''
    m = re.search(r'Co\.\s*Reg\.?\s*No\s*:\s*(\d{8,9}[A-Z])', t)
    d['uen'] = m.group(1) if m else ''
    # weight may come out split by the PDF spacing ("40.600 KG" -> "4 0.600 KG"): digits + optional space + digits
    W = r'(\d[\d,]*(?: \d[\d,]*)?\.\d+|\d[\d,]*)'
    Q = r'(\d{1,3}(?:,\d{3})+|\d+)'           # qty, also 1,460
    LN = re.compile(r'^([A-Z]{2}) (\S+) (.+?) (\d{10}) (?:([A-Z0-9]*[A-Z][A-Z0-9.]*) )?' + W + r' KG ' + Q + r' (' + N + r') ([A-Z]{3})$')
    unread = []
    for raw in t.splitlines():
        m = LN.match(raw.strip())
        if m:
            d['lines'].append(_line(item=m.group(2), desc=clean(m.group(3)), hs=_hs8(m.group(4)), hs_raw10=m.group(4), eccn=m.group(5) or '',
                                    nw=_num(m.group(6).replace(' ', '')), qty=_num(m.group(7)), uom='PCE',
                                    amount=_num(m.group(8)), currency=m.group(9), coo=m.group(1)))
        elif re.match(r'^[A-Z]{2} \S+ .+ \d{10}\b.*\b[A-Z]{3}$', raw.strip()):
            unread.append(raw.strip())
    if unread:
        d['notes'].append('%d item line(s) not readable - check them on the invoice: %s' % (len(unread), ' | '.join(unread[:5])))
    d['currency'] = d['lines'][0]['currency'] if d['lines'] else ''
    for l in d['lines']:
        if l['qty'] and l['amount'] is not None:
            l['price'] = round(l['amount'] / l['qty'], 4)
    tariff = re.findall(r'^(\d{10}) ' + W + r' KG ' + Q + r' (' + N + r') [A-Z]{3}$', t, re.M)
    if tariff:
        d['total'] = round(sum(_num(x[3]) or 0 for x in tariff), 2)
        d['total_nw'] = round(sum(_num(x[1].replace(' ', '')) or 0 for x in tariff), 3)
    bad_t = [ln.strip() for ln in t.splitlines() if re.match(r'^\d{10} .*KG', ln.strip())
             and not re.match(r'^(\d{10}) ' + W + r' KG ' + Q + r' (' + N + r') [A-Z]{3}$', ln.strip())]
    if bad_t:
        d['notes'].append('%d tariff summary line(s) not readable - check: %s' % (len(bad_t), ' | '.join(bad_t[:5])))
    # cross-check every tariff code: items (qty / net weight / value) against the TARIFF CODE SUMMARY
    if tariff:
        agg = collections.defaultdict(lambda: [0.0, 0.0, 0.0])
        for l in d['lines']:
            a_ = agg[l['hs_raw10'] if l.get('hs_raw10') else l['hs']]
            a_[0] += l['nw'] or 0
            a_[1] += l['qty'] or 0
            a_[2] += l['amount'] or 0
        summ = collections.defaultdict(lambda: [0.0, 0.0, 0.0])
        for h, w_, q_, am in tariff:
            s_ = summ[h]
            s_[0] += _num(w_.replace(' ', '')) or 0
            s_[1] += _num(q_) or 0
            s_[2] += _num(am) or 0
        diff = []
        for h, (w_, q_, am) in summ.items():
            a_ = agg.get(h, [0, 0, 0])
            if abs(a_[1] - q_) > 0.001 or abs(a_[2] - am) > 0.01:
                diff.append('%s: summary qty %s / %s, items qty %s / %s' % (h, fmt_kg(q_), '{:,.2f}'.format(am), fmt_kg(a_[1]),
                                                                           '{:,.2f}'.format(a_[2])))
            elif abs(a_[0] - w_) > 0.0015:
                diff.append('%s: summary net %s KG, items %s KG' % (h, fmt_kg(w_), fmt_kg(round(a_[0], 3))))
        extra = [h for h in agg if h not in summ]
        if diff or extra:
            d['notes'].append('items vs TARIFF CODE SUMMARY differ - check: %s' % '; '.join(diff[:8] + ['%s not in the summary' % h for h in extra[:4]]))
    m = re.search(r'Total Amount\s+' + W + r'\s*KG', t)
    if m:
        d['total_nw'] = _num(m.group(1).replace(' ', ''))
    d['notes'].append('Invoice SUMMARY (by country of origin / tariff) - one item row per part as listed; '
                      'no gross weight / packages on it - type them on tab 2 or add the packing list')
    return _finish(d, C)


# ======================================================================== Perkins Engines Company Ltd (UK) "Customs Only Invoice"
PKC_SKIP = re.compile(r'^(Perkins Engines Company|Parts Distribution Centre|Birchwood Park|Invoice Detail|Case Summary|Case Details|'
                      r'General Summary|Country of Origin / Tariff|Warrington|Summary WA3|WA3 6YF|Telephone Number|VAT Registration|HULMED|'
                      r'Billing Doc\.|Order Type Customer|Number Number|\d{9} \d{4,6} Standard Order|Customer Item Ans|Part Number Ordered|'
                      r'Order Number No\.|"Ans Code"|\'B\' =>|Perkins®|Perkins\s*\S* is the Regis)')


def perkins_cinv(path, pages, C):
    """Perkins Engines Company Ltd (Warrington, UK) Customs Only Invoice -> Perkins Engines (Asia Pacific), import.
    Invoice Detail: per item line (customer order, item no., part ordered / supplied, qty sent, unit price, value) and
    the block under it (description, Tariff Code, Country of Origin, ECCN, Qty Sent Net Wt = net KG of the line).
    Missing tariff / COO are taken from the Country of Origin / Tariff Summary (same part); Case Summary = packages."""
    d = _inv(path, 'Perkins customs only invoice', 'PERKINS ENGINES COMPANY LTD')
    d.update(ship_from_country='GB', title='CUSTOMS ONLY INVOICE', invoice_from='PERKINS ENGINES COMPANY LTD')
    t = '\n'.join(p['text'] for p in pages)
    m = re.search(r'Billing Document Number\s*(\d{6,12})', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'dd/mm/yyyy\s*(\d{1,2}/\d{1,2}/\d{4})', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'^Sold-to:\s*\n(?:Date\s*\n)?(.+?)\s+Billing Document', t, re.M)
    name = clean(m.group(1)).upper() if m else ''
    if name and not re.search(r'PTE', name):
        name += ' PTE LTD'
    d['ship_to'] = d['invoice_to'] = name or 'PERKINS ENGINES (ASIA PACIFIC) PTE LTD'
    d['ship_to_country'] = 'SG' if re.search(r'\bSingapore\b', t[:1500], re.I) else (C.find_in(t[:1500]) or '')
    m = re.search(r'^(EXW|FCA|FAS|FOB|CFR|CIF|CPT|CIP|DAP|DPU|DDP)\s+([A-Z ]+?)\s+([A-Z]{3})\s+(' + N + r')\s*$', t, re.M)
    if m:
        d['incoterm'], d['currency'], d['total'] = m.group(1), m.group(3), _num(m.group(4))
    else:
        m = re.search(r'All Product Prices\s+(' + N + r')', t)
        d['total'] = _num(m.group(1)) if m else None
    m = re.search(r'TOTAL\s+(' + N + r')\s+(' + N + r')\s+(' + N + r')\s*$', t, re.M)
    if m:
        d['total_gw'], d['total_nw'], d['cbm'] = _num(m.group(1)), _num(m.group(2)), _num(m.group(3))
    m = re.search(r'Port of Loading\s+(.+)', t)
    d['loading_port'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'Port of Discharge\s+(.+)', t)
    d['discharge_port'] = clean(m.group(1)).upper() if m else ''
    Q = r'(\d{1,3}(?:,\d{3})+|\d+)'
    IT = re.compile(r'^(\d{10}) (\d{1,6}) (\S+) (\S+)(?: ([A-Z]))? ' + Q + r' ' + Q + r' (\d+) (' + N + r') (' + N + r')$')
    KEY = re.compile(r'^(Order Type|Tariff Code|Country of Origin|ECCN|Customer Material Number|Qty Sent Net Wt|Customer\'s Item reference)\s*(?:\([^)]*\))?\s*:\s*(.*)$')
    cur, unread = None, []
    detail = [p['text'] for p in pages if re.search(r'^Invoice Detail\b', p['text'], re.M)]
    for txt in detail:
        for raw in txt.splitlines():
            ln = raw.strip()
            if not ln or PKC_SKIP.match(ln):
                continue
            if re.match(r'^Cumulative Total\b', ln):
                cur = None
                continue
            m = IT.match(ln)
            if m:
                qty, price, amt = _num(m.group(6)), _num(m.group(9)), _num(m.group(10))
                cur = _line(item=m.group(4), desc='', qty=qty, uom='PCE', price=price, amount=amt, currency=d.get('currency') or '')
                cur['po'] = m.group(1)
                cur['line'] = m.group(2)
                cur['part_ordered'] = m.group(3)
                d['lines'].append(cur)
                continue
            if re.match(r'^\d{10} \d{1,6} \S+', ln):
                unread.append(ln)
                cur = None
                continue
            if cur is None:
                continue
            k = KEY.match(ln)
            if k:
                key, val = k.group(1), clean(k.group(2))
                if key == 'Tariff Code' and re.match(r'\d{8,10}$', val):
                    cur['hs'], cur['hs_raw10'] = _hs8(val), val
                elif key == 'Country of Origin' and re.match(r'[A-Z]{2}$', val):
                    cur['coo'] = val
                elif key == 'ECCN':
                    cur['eccn'] = val
                elif key == 'Qty Sent Net Wt':
                    cur['nw'] = _num(val)
                elif key == 'Customer Material Number':
                    cur['cust_part'] = val
                continue
            if not cur['desc']:
                cur['desc'] = clean(ln).upper()
            elif not cur.get('hs') and not cur.get('coo') and len(cur['desc']) < 60:
                cur['desc'] = clean(cur['desc'] + ' ' + ln).upper()
    if unread:
        d['notes'].append('%d item line(s) not readable - check them on the invoice: %s' % (len(unread), ' | '.join(unread[:5])))
    # Country of Origin / Tariff Summary: part -> (COO, tariff, net, qty, value)
    summ = []
    SM = re.compile(r'^([A-Z]{2}) (\S+) (.+?) (\d{10}) (?:(?:NLR|EAR99|[0-9][A-E]\d{3}\S*) )?(' + N + r') ' + Q + r' (' + N + r')$')
    for p in pages:
        if 'Country of Origin / Tariff' not in p['text']:
            continue
        for raw in p['text'].splitlines():
            m = SM.match(raw.strip())
            if m:
                summ.append(dict(coo=m.group(1), part=m.group(2), tariff=m.group(4), nw=_num(m.group(5)), qty=_num(m.group(6)), value=_num(m.group(7))))
    by_part = collections.defaultdict(list)
    for x in summ:
        by_part[x['part']].append(x)
    filled = []
    for l in d['lines']:
        cand = by_part.get(l['item']) or []
        if not l.get('hs'):
            hs = {x['tariff'] for x in cand}
            if len(hs) == 1:
                l['hs_raw10'] = hs.pop()
                l['hs'] = _hs8(l['hs_raw10'])
                filled.append('%s tariff' % l['item'])
        if not l.get('coo'):
            co = {x['coo'] for x in cand}
            if len(co) == 1:
                l['coo'] = co.pop()
                filled.append('%s COO' % l['item'])
    if filled:
        d['notes'].append('taken from the Country of Origin / Tariff Summary (not under the item): %s' % ', '.join(filled[:12]))
    if summ:     # cross-check: per part qty / value
        agg = collections.defaultdict(lambda: [0.0, 0.0])
        for l in d['lines']:
            agg[l['item']][0] += l['qty'] or 0
            agg[l['item']][1] += l['amount'] or 0
        sp = collections.defaultdict(lambda: [0.0, 0.0])
        for x in summ:
            sp[x['part']][0] += x['qty'] or 0
            sp[x['part']][1] += x['value'] or 0
        diff = ['%s: summary %s / %.2f, items %s / %.2f' % (k, fmt_kg(v[0]), v[1], fmt_kg(agg.get(k, [0, 0])[0]), agg.get(k, [0, 0])[1])
                for k, v in sp.items() if abs(agg.get(k, [0, 0])[0] - v[0]) > 0.001 or abs(agg.get(k, [0, 0])[1] - v[1]) > 0.01]
        diff += ['%s not in the summary' % k for k in agg if k not in sp]
        if diff:
            d['notes'].append('items vs Country of Origin / Tariff Summary differ - check: %s' % '; '.join(diff[:8]))
    # Case Summary: handling units = packages (gross / net per unit)
    hu = {}
    for p in pages:
        if 'Case Summary' not in p['text']:
            continue
        for raw in p['text'].splitlines():
            m = re.match(r'^(\d{10}) (.+?) (\d+) (\d+) (\d+) (' + N + r') (' + N + r') (' + N + r') (\d{10})$', raw.strip())
            if m:
                hu[m.group(1)] = (clean(m.group(2)).upper(), _num(m.group(6)), _num(m.group(7)))
    if hu:
        d['packages'] = len(hu)
        kinds = {v[0] for v in hu.values()}
        d['package_unit'] = 'PLT' if all(re.search(r'PALLET|EURO|CHEP|DW\d|TRIWALL|CRATE', k) for k in kinds) else 'PKG'
        g = round(sum(v[1] or 0 for v in hu.values()), 3)
        if d.get('total_gw') and abs(g - d['total_gw']) > 0.5:
            d['notes'].append('Case Summary gross %s KG but the invoice total is %s KG' % (fmt_kg(g), fmt_kg(d['total_gw'])))
    s_nw = round(sum(l['nw'] or 0 for l in d['lines']), 3)
    sum_nw = round(sum(x['nw'] or 0 for x in summ), 3)
    if d.get('total_nw') and abs(s_nw - d['total_nw']) > 0.01:
        if summ and abs(s_nw - sum_nw) <= 0.01:
            d['notes'].append('item net weights %s KG (same as the Country of Origin / Tariff Summary) used - page 1 / Case Summary '
                              'shows %s KG net (handling units)' % (fmt_kg(s_nw), fmt_kg(d['total_nw'])))
            d['total_nw'] = s_nw
        else:
            d['notes'].append('item net weights add up to %s KG, invoice total net %s KG' % (fmt_kg(s_nw), fmt_kg(d['total_nw'])))
    for l in d['lines']:
        l['currency'] = d.get('currency') or l['currency']
        if l['qty'] and l['amount'] is not None and (not l['price'] or abs(l['qty'] * l['price'] - l['amount']) > 0.011):
            l['price'] = round(l['amount'] / l['qty'], 4)       # "price per pack" (pack of 10 ...) -> price per piece
    return _finish(d, C)


# ======================================================================== DHL "Sea Freight FCL / LCL Arrival Customs Instructions Import / Export"
def dhl_ci_sea(path, pages, C):
    """DHL (vCargoCloud) sea freight customs instructions: ocean + house B/L, vessel / voyage, carrier, load / discharge
    ports, ETD / ETA, packages, weight, container + seal, incoterm, place of release / receipt, permit type, remarks."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'DHL sea customs instructions', 3)
    exp = bool(re.search(r'Customs Instructions Export', t))
    cg = {'mode': 'Sea'}
    m = re.search(r'SHIPMENT\s*([A-Z]\d{8,12})', t)
    d['fwd_job'] = m.group(1) if m else ''
    if d['fwd_job']:
        cg['shipment_no'] = d['fwd_job']
    m = re.search(r'OCEAN BILL OF LADING\s+HOUSE BILL OF LADING\s*\n(.*)\n', t)
    if m:
        toks = re.findall(r'\b([A-Z][A-Z0-9]{7,19})\b', m.group(1).split('SRN:')[-1])
        toks = [x for x in toks if not re.match(r'^(SING|FFW|CLT|OTH)', x)]
        if len(toks) >= 2:
            d['mawb'], d['hawb'] = toks[-2], toks[-1]
        elif toks:
            d['mawb'] = toks[-1]
    for k in ('hawb', 'mawb'):
        if d[k]:
            cg[k] = d[k]
    m = re.search(r'SRN:\s*(\d{6,12})', t)
    if m:
        d['invoice_refs'] = [m.group(1)]
    m = re.search(r'^(SEA|AIR)\s+(.+?)\s*/\s*([A-Z0-9]+)\s*/\s*(\d{7})?\s*(.*?)\s+([A-Z]{5})\s+([A-Z]{5})\s+(\d{1,2}-[A-Za-z]{3}-\d{2,4})?\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})?\s*$', t, re.M)
    if m:
        cg['transport_id'], cg['flight'] = clean(m.group(2)).upper(), m.group(3)
        if m.group(5):
            cg['carrier_agent_hint'] = clean(m.group(5)).upper()
        cg['loading_code'], cg['discharge_code'] = m.group(6), m.group(7)
        if m.group(8):
            cg['etd'] = _dmon(m.group(8))
        if m.group(9):
            cg['eta'] = _dmon(m.group(9))
        cg['date'] = cg.get('etd' if exp else 'eta', '')
    m = re.search(r'^(\d+)\s+([A-Z]{2,4})\s*\(OUTER\).*?(' + N + r')\s*KG', t, re.M)
    if m:
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = int(m.group(1)), m.group(2), _num(m.group(3))
        d['pieces'], d['gw'] = cg['packages_qty'], cg['gw_kg']
    conts = []
    for mc in re.finditer(r'CONTAINER:\s*([A-Z]{4}\d{7})\s+TYPE:\s*(\d{2})([A-Z0-9]{2})\S*.*\n\s*SEAL:\s*(\S+).*?GROSS:\s*(' + N + r')\s*KG', t):
        size = mc.group(2)
        conts.append(dict(no=mc.group(1), size=size, size_text=('LCL' if 'LCL' in t[:300] else 'FCL') + size, seal=mc.group(4),
                          tne=(_num(mc.group(5)) or 0) / 1000.0 or None, type=size + mc.group(3)))
    if conts:
        cg['containers'] = [c['no'] for c in conts]
        cg['container_info'] = conts
    m = re.search(r'Harmonized Code:\s*(\d{4,10})', t)
    if m:
        d['awb_hs'] = m.group(1)
    m = re.search(r'INCO TERM\s*\n\s*([A-Z]{3})\b', t)
    if m:
        cg['term_type'] = d['awb_terms'] = m.group(1)
    m = re.search(r'SHIPPER\s+CONSIGNEE\s*\n(.+)', t)
    if m:
        row = clean(m.group(1)).upper()
        mm = re.search(r'(.+?\b(?:LTD|LIMITED|INC|GMBH|BV|B\.V|AG|SA|S\.A|NV/SA|SRL|SPA|PLC|KG)\.?)\s+(.+)$', row)
        d['shipper'], d['consignee'] = (mm.group(1), mm.group(2)) if mm else ('', row)
        cg['party_name'] = d['shipper'] if exp else d['consignee']
    u = re.search(r'GST:\s*(\d{8,9}[A-Z]|[STR]\d{2}[A-Z]{2}\d{4}[A-Z])\b', t)
    if u:
        cg['uen'] = u.group(1)
    for key, lab in (('release', 'Place of Release'), ('receipt', 'Place of Receipt'), ('permit_remarks', 'Permit Remarks'),
                     ('special', 'Special Instructions')):
        mm = re.search(lab + r'\s*:[ \t]*([^\n]*)', t, re.I)
        v = clean(mm.group(1)) if mm else ''
        if key in ('release', 'receipt') and re.fullmatch(r'(IG|IT|IM|OD|OT|GN)', v.upper()):
            d['notes'].append('%s "%s" is a permit type, not a location - left blank' % (lab, v))
            v = ''
        if v:
            cg[key] = v.upper() if key in ('release', 'receipt') else v
    mm = re.search(r'Type of Permit\s*:[ \t]*([^\n]*)', t, re.I)
    if mm:
        pt = re.findall(r'\b(IG|ME|IT|IM|OD|OT|GN)\b', mm.group(1).upper())
        if pt:
            cg['permit_type'] = pt[-1]
    m = re.search(r'Co\.?\s*Reg\.?\s*No:?\s*(\d{8,9}[A-Z])', t)
    if m and 'DHL GLOBAL FORWARDING (SINGAPORE)' in t.upper():
        d['issued_by'] = 'DHL GLOBAL FORWARDING (SINGAPORE) PTE LTD'
        cg['forwarder'], cg['forwarder_uen'] = d['issued_by'], m.group(1)
    car = cg.pop('carrier_agent_hint', '')        # the shipping line, not its Singapore agent - the arrival notice names that
    if re.search(r'OCEAN NETWORK EXPRESS', car):
        cg['carrier_agent'] = 'OCEAN NETWORK EXPRESS (SINGAPORE) PTE LTD'
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    d['direction'] = 'EXPORT' if exp else 'IMPORT'
    return [d]


# ======================================================================== ONE (Ocean Network Express) arrival notice / sea waybill
def one_an(path, pages, C):
    """ONE arrival notice: ocean B/L, vessel + voyage, ETA, container + seal, pieces, weight, port of discharge."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'ONE arrival notice', 2)
    cg = {'mode': 'Sea'}
    m = re.search(r'B/L No\.?\s*:\s*([A-Z0-9]{8,20})', t)
    if m:
        d['mawb'] = cg['mawb'] = m.group(1)
    m = re.search(r'Arrival Vessel\s*:\s*(.+?)\s+([0-9]{3,4}[A-Z]?)\s+O/BL', t)
    if m:
        cg['transport_id'], cg['flight'] = clean(m.group(1)).upper(), m.group(2)
    m = re.search(r'Port of Discharging\s+([A-Z ]+?)\s*(?:SHIPPER|\n)', t)
    if m:
        cg['discharge_name'] = clean(m.group(1)).upper()
    m = re.search(r'ETA\s+(\d{1,2}-[A-Z]{3}-\d{2,4})', t)
    if m:
        cg['date'] = cg['eta'] = _dmon(m.group(1))
    m = re.search(r'Place of Receipt\s+(.+)', t)
    if m:
        cg['loading_name'] = clean(m.group(1).split(',')[0]).upper()
    m = re.search(r'Total Piece Count\s+(\d+)\s*(\w+)?\s+Total Weight\s*:\s*(' + N + r')\s*KG', t)
    if m:
        cg['packages_qty'], cg['gw_kg'] = int(m.group(1)), _num(m.group(3))
        d['pieces'], d['gw'] = cg['packages_qty'], cg['gw_kg']
    conts = [dict(no=mc.group(1), seal=mc.group(3), tne=None) for mc in re.finditer(r'^([A-Z]{4}\d{7})\s+(\w{2})\s+(\S+)\s+(\d+)\b', t, re.M)]
    if conts:
        cg['containers'] = [c['no'] for c in conts]
        cg['container_info'] = conts
    cg['carrier_agent'] = 'OCEAN NETWORK EXPRESS (SINGAPORE) PTE LTD'
    d['issued_by'] = 'OCEAN NETWORK EXPRESS (SINGAPORE) PTE LTD'
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]


def one_swb(path, pages, C):
    """ONE sea waybill / bill of lading: B/L no., vessel + voyage, ports, container / seal / type, packages, weight, laden on board."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'ONE sea waybill', 2)
    cg = {'mode': 'Sea'}
    m = re.search(r'SEA WAYBILL NO\.\s*\n?\s*(?:[A-Z0-9]+\s+)?(ONEY[A-Z0-9]{6,16})', t) or re.search(r'B/L NO\.?:?\s*(ONEY[A-Z0-9]{6,16})', t)
    if m:
        d['mawb'] = cg['mawb'] = m.group(1)
    m = re.search(r'OCEAN VESSEL VOYAGE NO\..*\n(.+?)\s+([0-9]{3,4}[A-Z]?)\s+(.+)', t)
    if m:
        cg['transport_id'], cg['flight'] = clean(m.group(1)).upper(), m.group(2)
        cg['loading_name'] = clean(m.group(3).split(',')[0]).upper()
    m = re.search(r'PORT OF DISCHARGE.*\n\s*([A-Z ]+?)\s{1,}([A-Z ]+?)\s+(FCL|LCL)', t)
    if m:
        cg['discharge_name'] = clean(m.group(1)).upper()
    m = re.search(r'([A-Z]{4}\d{7})\s*/\s*(\S+)\s*/\s*(\d+)\s*([A-Z]+)\s*/\s*(FCL|LCL)\s*/\s*(?:FCL|LCL)\s*/\s*(\d{2})([A-Z]{2})\s*/\s*(' + N + r')\s*KGS', t)
    if m:
        cg['containers'] = [m.group(1)]
        cg['container_info'] = [dict(no=m.group(1), seal=m.group(2), size=m.group(6), size_text=m.group(5) + m.group(6),
                                     type=m.group(6) + m.group(7), tne=(_num(m.group(8)) or 0) / 1000.0 or None)]
        cg['packages_qty'], cg['gw_kg'] = int(m.group(3)), _num(m.group(8))
        d['pieces'], d['gw'] = cg['packages_qty'], cg['gw_kg']
    m = re.search(r'LADEN ON BOARD\s*\n?\s*(\d{1,2} [A-Z]{3} \d{4})', t)
    if m:
        cg['etd'] = _dmon(m.group(1))
    m = re.search(r'HS CODE\s*:\s*(\d{4,10})', t)
    if m:
        d['awb_hs'] = m.group(1)
    m = re.search(r'INVOICE\s*\n?\s*NO:?\s*(\d{6,12})', t)
    if m:
        d['invoice_refs'] = [m.group(1)]
    d['issued_by'] = 'OCEAN NETWORK EXPRESS'
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]


# ======================================================================== Legrand Singapore
def legrand(path, pages, C):
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Legrand invoice', 'LEGRAND (S) PTE LTD')
    d['title'] = 'TAX INVOICE'
    m = re.search(r'Invoice No\.\s*(\S+)', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice date\.?\s*(.+)', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    for k, rx in (('currency', r'Currency\s+([A-Z]{3})'), ('incoterm', r'Delivery Terms\s+([A-Z]{3})'),
                  ('ship_mode', r'Shipping Condition\s+(\w+)')):
        m = re.search(rx, t)
        d[k] = m.group(1) if m else ''
    m = re.search(r'Bill-to Address\s*\n\s*(.+)', t)
    d['ship_to'] = d['invoice_to'] = clean(m.group(1)).upper() if m else ''
    a, b = t.find('Bill-to Address'), t.find('Customer No')
    d['ship_to_country'] = C.find_in(t[a:b] if 0 <= a < b else t[:400]) or ''
    so = dn = po = ''
    cur = None
    LN = re.compile(r'^(\d{1,4}) (\S+) (.+?) (\d{8,10}) ([A-Z][A-Za-z ]+?) (' + N + r') ([A-Z]{1,3}) (' + N + r') (' + N + r')$')
    for raw in t.splitlines():
        ln = raw.strip()
        m = re.match(r'Sales Order:\s*(\d+)\s+Delivery No\.:\s*(\d+)\s+Customer PO No\.:\s*(\S+)', ln)
        if m:
            so, dn, po = m.groups()
            cur = None
            continue
        m = LN.match(ln)
        if m:
            cur = _line(line=m.group(1), item=m.group(2), desc=clean(m.group(3)), hs=_hs8(m.group(4)),
                        coo=C.code(m.group(5)) or '', qty=_num(m.group(6)), uom=m.group(7), price=_num(m.group(8)),
                        amount=_num(m.group(9)), order=so, delivery=dn, currency=d['currency'])
            cur['po'] = po
            d['lines'].append(cur)
            continue
        if cur is not None and ln and not re.match(r'(Commercial|Code|No\.|All supplies|Page|Sales Order|Net Value)', ln) \
                and len(ln) <= 30 and not re.search(r'\d{6,}', ln):
            cur['desc'] = clean(cur['desc'] + ' ' + ln)
            cur = None
        elif cur is not None:
            cur = None
    m = re.search(r'Net Value \(Excluding GST\)\s+(' + N + ')', t)
    d['total'] = _num(m.group(1)) if m else None
    dns = sorted({l['delivery'] for l in d['lines'] if l['delivery']})
    d['delivery'] = ', '.join(dns)
    return _finish(d, C)



# ======================================================================== Dalian Nurmi Hydraulics (scanned pack -> Sandvik SG)
def _o2zero(code):
    """OCR reads 0 as O inside part numbers (BGO1526031 -> BG01526031)."""
    return re.sub(r'(?<=[A-Z0-9])O(?=\d)|(?<=\d)O(?=[\dA-Z])', '0', code or '')


def nurmi(path, pages, C):
    """Summary page ("Contact") + one INVOICE page per order + Delivery Note (net / gross per article).
    The summary page repeats the invoices - it is NOT read as items, only used to cross-check the totals."""
    import readers as R
    kinds = []
    for p in pages:
        head = '\n'.join(p['text'].splitlines()[:6])
        kinds.append('DN' if re.search(r'Delivery\s+Note', head, re.I) else 'INV' if re.search(r'INVOICE', head) else
                     'SUM' if re.search(r'\bCont(?:r)?act\b', head) else '')
    invs, notes_all = [], []
    summary_total = None
    for p, k in zip(pages, kinds):
        t = p['text']
        if k == 'SUM':
            m = re.search(r'TOTAL AMOUNT[^\n]*?\b([A-Z]{3})\s+(' + N + r')\s*$', t, re.M)
            summary_total = _num(m.group(2)) if m else None
            continue
        if k != 'INV':
            continue
        d = _inv(path, 'Nurmi invoice', 'DALIAN NURMI HYDRAULICS LIMITED LIABILITY COMPANY')
        d.update(ship_from_country='CN', ocr=bool(p.get('ocr')))
        m = re.search(r'(\d{4})\.(\d{1,2})\.(\d{1,2})\s+(DNH[0-9O]{6,}(?:-\d+)?)', t)
        if m:
            d['invoice_date'] = _dmy(m.group(3), m.group(2), m.group(1))
            d['invoice_no'] = _o2zero(m.group(4))
        m = re.search(r'TOTAL AMOUNT[^\n]*?\b([A-Z]{3})\s+(' + N + r')\s*$', t, re.M)
        if m:
            d['currency'], d['total'] = m.group(1), _num(m.group(2))
        m = re.search(r'Terms of Delivery.*\n\s*([A-Z]{3})\b\s*([A-Z ]*)', t)
        if m:
            d['incoterm'], d['incoterm_place'] = m.group(1), clean(m.group(2))
        m = re.search(r'Your order number\s*\n.*?(\d{10})', t)
        d['customer_po'] = d['order_no'] = m.group(1) if m else ''
        m = re.search(r'Consignee\s*\n.*?((?:Sandvik|[A-Z][A-Za-z]+) [^\n]*?(?:Ltd|LTD|Pte|PTE)\.?)\s*$', t, re.M)
        d['ship_to'] = clean(m.group(1)).upper() if m else ''
        m = re.search(r'(C/O [^\n]+?PTE\.? ?LTD)', t, re.I)
        if m:
            d['ship_to_block'] = [d['ship_to'], clean(m.group(1)).upper()]
        m = re.search(r'Invoice To\s+Consignee\s*\n\s*(.+?)\s{1,}(?:Sandvik SMC|[A-Z][a-z]+ [A-Z]{3})', t)
        d['invoice_to'] = clean(m.group(1)).upper() if m else ''
        d['ship_to_country'] = 'SG' if re.search(r'SINGAPORE\s+\d{6}', t) else ''
        m = re.search(r'Way of Delivery.*\n.*?\b(DHL|[A-Z]{3,})\s*$', t, re.M)
        d['carrier'] = m.group(1) if m else ''
        for raw in t.splitlines():
            ln = raw.strip()
            m = re.match(r'^(\d{2,4})\s+(.+?)\s+(\d{8,10})\s+(' + N + r')\s+(' + N + r')\s+(' + N + r')$', ln)
            if m:
                txt = m.group(2).split()
                item = _o2zero(txt[-1]) if len(txt) > 1 else ''
                d['lines'].append(_line(line=m.group(1), item=item, desc=clean(' '.join(txt[:-1]) if item else m.group(2)),
                                        hs=_hs8(m.group(3)), qty=_num(m.group(4)), uom='PCE', price=_num(m.group(5)),
                                        amount=_num(m.group(6)), order=d['customer_po'], currency=d['currency'], coo='CN'))
            m = re.match(r'^Package Fee\s+(' + N + r')$', ln)
            if m:
                d['charges'], d['charges_label'] = _num(m.group(1)), 'Package fee'
        s_ = round(sum(l['amount'] or 0 for l in d['lines']) + (d.get('charges') or 0), 2)
        if d['total'] is not None and abs(s_ - d['total']) > 0.01:
            d['notes'].append('items + package fee = %s but the invoice total is %s - OCR, check' % ('{:,.2f}'.format(s_), '{:,.2f}'.format(d['total'])))
        d['notes'].append('country of origin not printed - CN taken from the manufacturer (Dalian, China); check')
        if d['ocr']:
            d['notes_ocr'] = 'scanned invoice read by OCR - check item no / HS / qty / amounts'
        _finish(d, C)
        d['total'] = d['total'] if d['total'] is not None else s_
        d['notes'] = [n for n in d['notes'] if not n.startswith('item lines add up to')]   # fee is in the total
        invs.append(d)
    # summary page = all invoices together
    if summary_total is not None and invs:
        tot = round(sum(i['total'] or 0 for i in invs), 2)
        msg = ('summary page total %s = the %d invoices %s' if abs(tot - summary_total) <= 0.01 else
               'summary page total %s but the %d invoices add up to %s - an invoice page is missing or misread') % (
            '{:,.2f}'.format(summary_total), len(invs), '{:,.2f}'.format(tot))
        invs[0]['notes'].append(msg + ' (summary page not counted as items)')
    # delivery note -> packing list: net / gross per article, packages, totals
    out = list(invs)
    for p, k in zip(pages, kinds):
        if k != 'DN':
            continue
        t = p['text']
        pl = R.new_pl(path, 'Nurmi delivery note')
        pl['ocr'] = bool(p.get('ocr'))
        pl['invoice_refs'] = [i['invoice_no'] for i in invs]
        m = re.search(r'TOTAL WEIGHT[^\n]*?(' + N + r')\s*KG\s*/\s*(' + N + r')\s*KG', t)
        if m:
            pl['total_nw'], pl['total_gw'] = _num(m.group(1)), _num(m.group(2))
        m = re.search(r'^Total:?\s+(\d+)\s+(' + N + r')\s+(' + N + r')\s*$', t, re.M)
        if m:
            pl['total_nw'] = pl['total_nw'] or _num(m.group(2))
            pl['total_gw'] = pl['total_gw'] or _num(m.group(3))
        m = re.search(r'TOTALLY in \(?(\d+)\)?\s*(\w+(?: \w+)?)', t)
        if m:
            pl['packages'], pl['package_unit'] = int(m.group(1)), m.group(2).upper()
        rows = []
        for raw in t.splitlines():
            m = re.match(r'^(.+?)\s+(\d+)\s+(' + N + r')\s+(' + N + r')\s*$', raw.strip())
            if m and not raw.strip().startswith(('Total', 'Pos')):
                rows.append(dict(text=_o2zero(m.group(1)), qty=_num(m.group(2)), nw=_num(m.group(3)), gw=_num(m.group(4))))
        lines = [l for i in invs for l in i['lines']]
        for r in rows:
            sq = re.sub(r'\W', '', r['text'].upper())
            hits = [l for l in lines if l['item'] and re.sub(r'\W', '', l['item'].upper()) in sq]
            if not hits:
                continue
            q = sum(h['qty'] or 0 for h in hits) or 1
            for h in hits:
                f = (h['qty'] or 0) / q
                h['nw'], h['gw'] = round(r['nw'] * f, 3), round(r['gw'] * f, 3)
            pl['lines'].append(R.new_line(item=hits[0]['item'], desc=hits[0]['desc'], qty=r['qty'], nw=r['nw'], gw=r['gw']))
            if r['qty'] != q:
                invs[0]['notes'].append('Delivery note qty %s for %s, invoices %s - check' % (r['qty'], hits[0]['item'], q))
        for i in invs:                                  # weights now come from the delivery note
            if all(l.get('nw') for l in i['lines']):
                i['notes'] = [n for n in i['notes'] if not n.startswith('no net weight')]
        miss = [l['item'] for l in lines if l.get('nw') is None]
        if miss:
            invs[0]['notes'].append('no delivery-note weight found for: %s' % ', '.join(miss))
        out.append(pl)
    return out


# ======================================================================== dispatch
# ======================================================================== shared: dd-Mon-yy dates, AWB documents
def _dmon(s):
    m = re.search(r'\b(\d{1,2})[- ]([A-Za-z]{3})[a-z]*[- ,]+(\d{2,4})\b', s or '')
    if m and m.group(2).upper() in MONTHS:
        return _dmy(m.group(1), MONTHS[m.group(2).upper()], m.group(3))
    return _date_any(s or '')


def _awb(path, reader, priority):
    """Transport document (B/L, sea waybill, arrival notice, customs instructions) in the common AWB shape.
    priority: which document's values win when several give the same field (higher = wins)."""
    return dict(kind='AWB', file=os.path.basename(path), path=path, notes=[], reader=reader, hawb='', mawb='',
                pieces=None, gw=None, departure='', legs=[], flights=[], executed='', destination='', consignee='',
                shipper='', fwd_job='', awb_hs='', awb_terms='', issued_by='', cargo={}, priority=priority)


def _port(code_name):
    """'CNSHA = Shanghai, China' -> (CNSHA, SHANGHAI)"""
    m = re.match(r'\s*([A-Z]{5})\s*(?:=\s*([^,\n]+))?', code_name or '')
    if not m:
        return '', clean(code_name or '').upper()
    return m.group(1), clean(m.group(2) or '').upper()


def _uen(t):
    m = re.search(r'\bUEN\s*:?\s*(\d{8,9}[A-Z]|[STR]\d{2}[A-Z]{2}\d{4}[A-Z])\b', t)
    return m.group(1) if m else ''


# ======================================================================== DHL customs instructions (vCargoCloud)
def dhl_ci(path, pages, C):
    """DHL 'Customs Instructions' (BROKERAGE BSIN...): importer + UEN, packages, weight, carrier agent, ocean / house B/L,
    vessel / voyage, ports, ETA / ATA, place of release / receipt, container, incoterm, permit type, freight."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'DHL customs instructions', 3)
    cg = {}
    m = re.search(r'BROKERAGE\s*(BSIN\d+|[A-Z]{3,5}\d{5,})', t)
    d['fwd_job'] = m.group(1) if m else ''
    exp = bool(re.search(r'\bEXPORTER\s+BILL TO\b', t))
    m = re.search(r'(?:IMPORTER|EXPORTER)\s+BILL TO\s*\n(.+)', t)
    if m:
        line = clean(m.group(1))
        half = len(line) // 2           # "NAME NAME" (importer and bill-to side by side)
        name = line[:half].strip() if line[:half].strip() == line[half:].strip() else line
        mm = re.match(r'(.+?\b(?:PTE\.?\s*LTD|LIMITED|LTD|PTE)\.?)\s+\S', name, re.I)
        if name == line and mm:             # "VF SINGAPORE ... PTE. LTD. ACTION EXPRESS INTERNATIONAL (AEI) - JORDAN"
            name = mm.group(1)
        cg['party_name'] = re.sub(r'\s+-\s*PO\s*:.*$', '', name).upper()     # "SATAIR PTE LTD - PO: 7000013555 - AIRBUS SM"
    u = _uen(t)
    if u:
        cg['uen'] = u
    m = re.search(r'PACKAGES\s+WEIGHT\s+VOLUME\s*\n\s*(\d+)\s+([A-Z]{2,4})\s*\(OUTER\).*?(' + N + r')\s*KG', t)
    if m:
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = int(m.group(1)), m.group(2), _num(m.group(3))
        d['pieces'], d['gw'] = cg['packages_qty'], cg['gw_kg']
    m = re.search(r'INWARD TRANSPORTATION\s*\n.*?\b(SEA|AIR|ROAD|RAIL)\s*$|OUTWARD TRANSPORTATION\s*\n.*?\b(SEA|AIR|ROAD|RAIL)\s*$', t, re.M)
    if m:
        cg['mode'] = (m.group(1) or m.group(2)).title()
    m = re.search(r'CARRIER AGENT\s+OCEAN BILL OF LADING\s*\n(.*)\n', t)
    if m:
        row = m.group(1)
        mo = re.search(r'([A-Z0-9]{8,20})\s*$', row)
        obl = mo.group(1) if mo else ''
        agent = row[:mo.start()] if mo else row
        agent = clean(agent.split('=', 1)[1] if '=' in agent else agent)
        if agent:
            cg['carrier_agent'] = agent.upper()
        if obl:
            d['mawb'] = obl
    m = re.search(r'HOUSE BILL OF LADING\s*\n(.*)\n', t)
    if m:
        row = m.group(1)
        ml = re.search(r'([A-Z0-9]{8,20}(?:\s*,\s*[A-Z0-9]{8,20})+),?\s*$', row)
        mh = ml or re.search(r'([A-Z0-9]{8,20})\s*$', row)
        if ml:                                  # "6520079427,6520079428,6520079439" - several house B/Ls, one permit
            d['house_list'] = [x.strip() for x in ml.group(1).split(',') if x.strip()]
        elif mh:
            d['hawb'] = mh.group(1)
        parts = [clean(x) for x in (row[:mh.start()] if mh else row).split('/')]
        if parts and parts[0]:
            cg['transport_id'] = parts[0].upper()
        if len(parts) > 1 and parts[1]:
            cg['flight'] = parts[1].upper()
    m = re.search(r'Carrier\s*:\s*(.+)', t)
    if m:
        car = m.group(1)
        mb = re.search(r'B/?L\s*#?\s*:?\s*([A-Z0-9]{8,20})', car)
        if mb and (not d['mawb'] or d['mawb'] == d['hawb']):
            d['mawb'] = mb.group(1)         # the carrier's ocean B/L when the box repeats the house number
        if not cg.get('carrier_agent'):
            names = [clean(x) for x in re.split(r'[:/]', re.sub(r'B/?L\s*#.*', '', car)) if clean(x)]
            if names:
                cg['carrier_agent'] = names[-1].upper()
    if d['mawb'] and d['mawb'] != d['hawb']:
        cg['mawb'] = d['mawb']
    if d['hawb']:
        cg['hawb'] = d['hawb']
    m = re.search(r'PORT OF ORIGIN\s+ETD\s+DESTINATION\s+ETA\s*\n\s*([A-Z]{5})?\s*(\S+)?\s+([A-Z]{5})?\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})?', t)
    eta = m.group(4) if m else ''
    etd = m.group(2) if m and m.group(2) and re.match(r'\d{1,2}-[A-Za-z]{3}-\d{2,4}$', m.group(2)) else ''
    m = re.search(r'PORT OF LOADING\s+ATD\s+PORT OF DISCHARGE\s+ATA\s*\n(.*)\n', t)
    if m:
        row = m.group(1)
        mm = re.match(r'\s*([A-Z]{5}\s*=\s*[^0-9]+?)\s+(\d{1,2}-[A-Za-z]{3}-\d{2,4})?\s*([A-Z]{5}\s*=\s*[^0-9]+?)\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})?\s*$', row)
        if mm:
            cg['loading_code'], cg['loading_name'] = _port(mm.group(1))
            cg['discharge_code'], cg['discharge_name'] = _port(mm.group(3))
            if mm.group(4):
                eta, ata = mm.group(4), True
            if mm.group(2):
                etd = mm.group(2)
    if eta:
        cg['date'] = cg['eta'] = _dmon(eta)          # import: arrival date
    if etd:
        cg['etd'] = _dmon(etd)                        # export: departure date
    m = re.search(r'PLACE OF RELEASE\s+PLACE OF RECEIPT\s*\n\s*(\S+)?\s+(\S+)?\s*$', t, re.M)
    for key, lab in (('release', 'Place of Release'), ('receipt', 'Place of Receipt'), ('permit_remarks', 'Permit Remarks'),
                     ('special', 'Special Instructions')):
        mm = re.search(lab + r'\s*:[ \t]*([^\n]*)', t, re.I)
        v = clean(mm.group(1)) if mm else ''
        if v:
            cg[key] = v.upper() if key in ('release', 'receipt') else v
    mm = re.search(r'Type of Permit\s*:[ \t]*([^\n]*)', t, re.I)
    if mm:
        pt = re.findall(r'\b(IG|ME|IT|IM|OD|OT)\b', mm.group(1).upper())
        if pt:
            cg['permit_type'] = pt[-1]
    conts = []
    for mc in re.finditer(r'^\s*([A-Z]{4}\d{7})\s+(\d{2}[A-Z]{2})\s+(\S+)\s+(' + N + r')\s+(FCL|LCL)\b\s*([A-Z]{3})?', t, re.M):
        size = mc.group(2)[:2]
        conts.append(dict(no=mc.group(1), size=size, size_text=mc.group(5) + size, seal=mc.group(3),
                          tne=(_num(mc.group(4)) or 0) / 1000.0 or None, type=mc.group(2)))
        if mc.group(6):
            d['awb_terms'] = mc.group(6)
    if conts:
        cg['containers'] = [c['no'] for c in conts]
        cg['container_info'] = conts
    fr = re.findall(r'\b(FRT|FREIGHT|BAF|CAF|THC|LSS)\s*[:=]?\s*([A-Z]{3})\s*(' + N + r')', cg.get('special', '') + ' ' + cg.get('permit_remarks', ''), re.I)
    fr = [x for x in fr if x[0].upper() in ('FRT', 'FREIGHT', 'BAF', 'CAF', 'LSS')]
    if fr and len({x[1].upper() for x in fr}) == 1:
        cg['freight_currency'] = fr[0][1].upper()
        cg['freight_amount'] = round(sum(_num(x[2]) or 0 for x in fr), 2)
        cg['freight_note'] = ' + '.join('%s %s %s' % (x[0].upper(), x[1].upper(), x[2]) for x in fr)
    m = re.search(r'Co\.?\s*Reg\.?\s*No:?\s*(\d{8,9}[A-Z])', t)
    if m and 'DHL GLOBAL FORWARDING' in t.upper():
        d['issued_by'] = 'DHL GLOBAL FORWARDING (SINGAPORE) PTE LTD'
        cg['forwarder'], cg['forwarder_uen'] = 'DHL GLOBAL FORWARDING (SINGAPORE) PTE LTD', m.group(1)
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    d['direction'] = 'EXPORT' if exp else 'IMPORT'
    return [d]


# ======================================================================== DSV arrival notice / sea waybill / B/L
def dsv_an(path, pages, C):
    """DSV 'Sea Freight FCL / LCL Arrival Notice': ocean + house B/L, vessel / voyage, ports, ETA / ATA, container."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'DSV arrival notice', 2)
    cg = {'mode': 'Sea'}
    m = re.search(r'OCEAN BILL OF LADING\s+HOUSE BILL OF LADING\s*\n(.*)\n', t)
    if m:
        toks = re.findall(r'\b([A-Z][A-Z0-9]{7,19})\b', m.group(1)) or []
        if len(toks) < 2:                  # numeric ocean B/L (e.g. 2335357770 SDV000018440)
            toks = re.findall(r'\b([A-Z0-9]{8,20})\b(?!,)', m.group(1))
        if len(toks) >= 2:
            d['mawb'], d['hawb'] = toks[-2], toks[-1]
        elif toks:
            d['hawb'] = toks[-1]
    if d['hawb']:
        cg['hawb'] = d['hawb']
    if d['mawb']:
        cg['mawb'] = d['mawb']
    m = re.search(r'^SEA\s+(.+?)\s*/\s*([A-Z0-9]*)\s*/', t, re.M)
    if m:
        cg['transport_id'] = clean(m.group(1)).upper()
        if m.group(2):
            cg['flight'] = m.group(2)
    m = re.search(r'([A-Z]{5})\s*=\s*([A-Z]{5})\s*=\s*(\d{1,2}-[A-Za-z]{3}-\d{2})\s+(\d{1,2}-[A-Za-z]{3}-\d{2})', t)
    if m:
        cg['loading_code'], cg['discharge_code'] = m.group(1), m.group(2)
        cg['date'] = cg['eta'] = _dmon(m.group(4))
        cg['etd'] = _dmon(m.group(3))
    m = re.search(r'^(\d+)\s+(PLT|CTN|PKG|PCS|PAL|BOX|CASE)S?\b(?:\s*\(OUTER\),\s*\d+\s*[A-Z]*\s*\(INNER\))?\s+(' + N + r')\s*KG', t, re.M)
    if m:
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = int(m.group(1)), m.group(2), _num(m.group(3))
        d['pieces'], d['gw'] = cg['packages_qty'], cg['gw_kg']
    conts = []
    for mc in re.finditer(r'^([A-Z]{4}\d{7})\s+(\S+)\s+(\d{2}[A-Z]{2})\s+(FCL|LCL)\s+(' + N + r')\s*KG', t, re.M):
        size = mc.group(3)[:2]
        conts.append(dict(no=mc.group(1), seal=mc.group(2), size=size, size_text=mc.group(4) + size,
                          tne=(_num(mc.group(5)) or 0) / 1000.0 or None, type=mc.group(3)))
    if not conts:                          # "OOCU9889820 0504970 40HICUBE GRP 166.000 KG 0.248 M3 1 PKG" (LCL arrival notice)
        for mc in re.finditer(r'^([A-Z]{4}\d{7})\s+(\S+)\s+(\d{2})([A-Z]+)\s+(?:[A-Z]{2,4}\s+)?(' + N + r')\s*KG', t, re.M):
            lcl = bool(re.search(r'LCL Arrival', t))
            conts.append(dict(no=mc.group(1), seal=mc.group(2), size=mc.group(3), size_text=('LCL' if lcl else 'FCL') + mc.group(3),
                              tne=(_num(mc.group(5)) or 0) / 1000.0 or None, type=mc.group(3) + mc.group(4)[:2]))
    if conts:
        cg['containers'] = [c['no'] for c in conts]
        cg['container_info'] = conts
    m = re.search(r'SHIPPER\s+CONSIGNEE\s*\n(.+)', t)
    if m:
        d['shipper'] = clean(m.group(1)).upper()
    d['issued_by'] = 'DSV AIR & SEA SINGAPORE PTE LTD' if 'myDSV' in t else ''
    if 'myDSV' not in t:
        d['reader'] = 'Sea arrival notice'
        m = re.search(r'CONSIGNOR\s+CONSIGNEE\s*\n(.+?)\s{1,}(?:SANDVIK|[A-Z][A-Z.&]+\s+[A-Z])', t)
        mm = re.search(r'^CARRIER\s+GOODS TO BE CLEARED BY\s*\n(.+?)\s{2,}|^CARRIER.*\n([A-Z][A-Z ]+?(?:LIMITED|LTD|LINE))', t, re.M)
        if mm:
            cg['carrier'] = clean(mm.group(1) or mm.group(2)).upper()
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]


def dsv_bl(path, pages, C):
    """DSV combined transport B/L or sea waybill: B/L No. (house), vessel / voyage, ports, packages, weight, containers."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'DSV bill of lading', 4)
    cg = {'mode': 'Sea'}
    m = re.search(r'(?:B/L No\.|Sea Waybill No\.)\s*\n?\s*(?:.*?\s)?([A-Z]{2,4}\d{6,10})\b', t)
    if m:
        d['hawb'] = cg['hawb'] = m.group(1)
    m = re.search(r'Ocean Vessel\s+Port of loading\s*\n(.*)\n', t)
    row = m.group(1) if m else ''
    mv = re.match(r'\s*(.+?)\s*/\s*([A-Z0-9]+)\s+(.+?,\s*[A-Z ]+?)(?:\s{1,}(?:Express|Original).*)?$', row)
    if mv:
        cg['transport_id'], cg['flight'] = clean(mv.group(1)).upper(), mv.group(2)
        cg['loading_name'] = clean(mv.group(3).split(',')[0]).upper()
    else:
        mp = re.search(r'Pre-carriage by\s+Place of Receipt\s*\n(.*)\n', t)
        if mp:
            mv = re.match(r'\s*(.+?)\s*/\s*([A-Z0-9]+)\s+(.+)$', mp.group(1))
            if mv:
                cg['transport_id'], cg['flight'] = clean(mv.group(1)).upper(), mv.group(2)
                cg['loading_name'] = clean(mv.group(3).split(',')[0]).upper()
    m = re.search(r'([\d ]+)\s*C\s*a\s*r\s*t\s*o\s*n|(\d+)\s+Pallet', t)
    for mc in re.finditer(r'^([A-Z]{4}\d{7})\s+(\S+)\s+(\d{2}[A-Z]{2})\s+(\d+)\s+([A-Z]{2,4})\s+(\S+)\s+(' + N + r')\s*KG', t, re.M):
        size = mc.group(3)[:2]
        lcl = 'CFS' in mc.group(6).upper()
        cg.setdefault('container_info', []).append(dict(no=mc.group(1), seal=mc.group(2), size=size, size_text=('LCL' if lcl else 'FCL') + size,
                                                       tne=(_num(mc.group(7)) or 0) / 1000.0 or None, type=mc.group(3)))
        cg['containers'] = [c['no'] for c in cg['container_info']]
    m = re.search(r'SHIPPED ON BOARD\s+(\d{1,2}-[A-Z]{3}-\d{2,4})', t)
    if m:
        d['executed'] = cg['etd'] = _dmon(m.group(1))         # shipped on board = departure (export)
    d['issued_by'] = 'DSV AIR & SEA SINGAPORE PTE LTD'
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]


# ======================================================================== Shanghai Everest Fine Chemicals (Halliburton)
def everest(path, pages, C):
    """Shanghai Everest Fine Chemicals COMMERCIAL INVOICE / PACKING LIST (to Halliburton Far East, Singapore cross dock):
    one HS + COO for the document, lines per PO with Material No., qty in pounds; packing list weights per material."""
    t = '\n'.join(p['text'] for p in pages)
    is_pl = bool(re.search(r'PACKING\s+LIST', t))
    m = re.search(r'Invoice No\.\s*[：:]\s*([A-Z0-9-]+)', t)
    inv_no = m.group(1) if m else ''
    m = re.search(r'Date\s*:\s*([A-Za-z]{3})\.?\s*(\d{1,2})\.(\d{4})', t)
    date = _dmy(m.group(2), MONTHS.get(m.group(1).upper(), 1), m.group(3)) if m else ''
    m = re.search(r'Country of Origin\s*:\s*([A-Za-z ]+?)\s*$', t, re.M)
    coo = C.code(m.group(1)) if m else ''
    m = re.search(r'HS Code\s*:\s*([\d.]+)', t)
    hs = re.sub(r'\D', '', m.group(1)) if m else ''
    m = re.search(r'INTERMEDIATE CONSIGNEE:.*\n\s*(.+?)\s+HS Code', t)
    imp = clean(m.group(1)).upper() if m else ''
    LN = re.compile(r'^\s*(\d{5})\s+(' + N + r')\s+([A-Za-z]+)\s+Material No\.\s*:\s*(\S+)\s+(.*)$')
    po, lines = '', []
    rows = t.splitlines()
    for i, raw in enumerate(rows):
        mp = re.match(r'\s*PO No\.\s*:\s*(.+)', raw)
        if mp:
            po = clean(mp.group(1)).replace(' ', '')
            continue
        m = LN.match(raw)
        if not m:
            continue
        desc = clean(rows[i + 1]) if i + 1 < len(rows) and not LN.match(rows[i + 1]) else ''
        uom = {'POUNDS': 'LB', 'POUND': 'LB', 'LBS': 'LB', 'KGS': 'KG', 'KG': 'KG', 'PCS': 'PCE'}.get(m.group(3).upper(), m.group(3).upper())
        rest = m.group(5)
        if is_pl:
            ws = re.findall(r'(' + N + r')\s*kgs', rest, re.I)
            lines.append(dict(line=m.group(1), po=po, item=m.group(4), desc=desc, qty=_num(m.group(2)), uom=uom,
                              nw=_num(ws[0]) if ws else None, gw=_num(ws[1]) if len(ws) > 1 else None))
        else:
            am = re.findall(r'\$\s*(' + N + r')', rest)
            lines.append(dict(line=m.group(1), po=po, item=m.group(4), desc=desc, qty=_num(m.group(2)), uom=uom,
                              price=_num(am[0]) if am else None, amount=_num(am[-1]) if len(am) > 1 else None))
    if is_pl:
        import readers as R
        d = R.new_pl(path, 'Everest packing list')
        d['invoice_refs'] = [inv_no] if inv_no else []
        for l in lines:
            d['lines'].append(_line(line=l['line'], item=l['item'], desc=l['desc'], qty=l['qty'], uom=l['uom'],
                                    nw=l['nw'], gw=l['gw'], order=l['po']))
        m = re.search(r'TOTAL:\s*(' + N + r')\s*\w+\s+(' + N + r')\s*kgs\s+(' + N + r')\s*kgs', t, re.I)
        if m:
            d['total_nw'], d['total_gw'] = _num(m.group(2)), _num(m.group(3))
        m = re.search(r'loaded on\s+(\d+)\s+(pallets?|cartons?|cases?)', t, re.I)
        if m:
            d['packages'], d['package_unit'] = int(m.group(1)), 'PLT' if m.group(2).lower().startswith('pallet') else m.group(2).upper()[:3]
        return [d]
    d = _inv(path, 'Everest invoice', 'SHANGHAI EVEREST FINE CHEMICALS CO.,LTD')
    d['ship_from_country'] = 'CN'
    d.update(invoice_no=inv_no, invoice_date=date, ship_to=imp, invoice_to=imp, ship_to_country='SG', direction='IMPORT',
             coo_text='China' if coo == 'CN' else '', hs_global=hs, currency='USD')
    m = re.search(r'Delivery Terms\s*:\s*([A-Z]{3})', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'TOTAL\s*:\s*US\$\s*(' + N + r')', t)
    d['total'] = _num(m.group(1)) if m else None
    for l in lines:
        d['lines'].append(_line(line='%s-%s' % (l['po'], l['line']), item=l['item'], desc=l['desc'], qty=l['qty'], uom=l['uom'],
                                price=l['price'], amount=l['amount'], hs=hs.ljust(8, '0') if len(hs) == 6 else _hs8(hs),
                                coo=coo, order=l['po'], currency='USD'))
    if hs and len(hs) < 8:
        d['notes'].append('HS %s on the invoice has %d digits - %s used (check against the HS list)' % (hs, len(hs), hs.ljust(8, '0')))
    d = _finish(d, C)
    d['notes'] = [n for n in d['notes'] if not n.startswith('no net weight on the invoice')]     # weights: packing list
    return [d]


# ======================================================================== VF (Infor Nexus) invoice + packing list
def vf_invoice(path, pages, C):
    """VF - INVOICE (Infor Nexus) - Timberland / The North Face ... garments: PO line, buyer item #, short + goods
    description, HTS, COO, qty, unit price, amount. Columns are taken from the word positions."""
    import pdfplumber
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'VF invoice (Infor Nexus)', '')
    m = re.search(r'^(.+?)\s+VF - INVOICE', t, re.M)
    d['supplier'] = d['invoice_from'] = d['ship_from'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'Invoice Number\s+Invoice Issue date\s*\n(?:.*?\s)?([A-Z]?\d{6,})\s+(\d{4}-\d{2}-\d{2})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _date_any(m.group(2))
    m = re.search(r'VF Incoterm 1.*\n\s*(.+?)\s+([A-Z]{3})\s*$', t, re.M)
    if m:
        d['currency'] = m.group(2)
        d['incoterm'] = {'FREE ON BOARD': 'FOB', 'FREE CARRIER': 'FCA', 'EX WORKS': 'EXW', 'COST AND FREIGHT': 'CFR',
                         'COST INSURANCE AND FREIGHT': 'CIF'}.get(clean(m.group(1)).upper().split(' CHITTAGONG')[0].strip(), '')
    d['ship_to'] = d['invoice_to'] = 'VF SINGAPORE OVERSEAS SERVICES PTE. LTD.' if 'VF Singapore Overseas' in t else ''
    d['ship_to_country'], d['direction'] = 'SG', 'IMPORT'
    m = re.search(r'BANGLADESH|VIETNAM|CHINA|INDONESIA|CAMBODIA|INDIA|SRI LANKA', t[:600])
    d['ship_from_country'] = C.code(m.group(0)) if m else ''
    m = re.search(r'Invoice Total\s+(' + N + r')', t)
    d['total'] = _num(m.group(1)) if m else None
    m = re.search(r'^Brand\s+Region.*\n\s*([A-Z][A-Z0-9 &\'.-]+?)\s+(?:APAC|EMEA|AMERICAS|NA|LATAM|GLOBAL)\b', t, re.M)
    if m:
        d['brand'] = clean(m.group(1)).upper()        # e.g. TIMBERLAND -> Brand on every item
    m = re.search(r'Total Cartons\s*:\s*(\d+)', t)
    if m:
        d['boxes'], d['package_unit'] = int(m.group(1)), 'CTN'
    LN = re.compile(r'^(\d{15})\s')
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            words = pg.extract_words(x_tolerance=1.5, keep_blank_chars=False)
            hts_x = coo_x = None
            rows = collections.OrderedDict()
            for w in words:
                rows.setdefault(round(w['top'] / 3), []).append(w)
            cur = None
            for k in sorted(rows):
                ws = sorted(rows[k], key=lambda w: w['x0'])
                txt = ' '.join(w['text'] for w in ws)
                m = re.match(r'^(\d{15})\s+(\S+)\s+(\S+)\s+(.+?)\s+(\d{8,10})\s+(.+?)\s+([A-Z]{2})\s+(' + N + r')\s+([A-Z]+)\s+(' + N
                             + r')\s+(' + N + r')\s+([A-Z]+)\s+(' + N + r')\s+(' + N + r')$', txt)
                if m:
                    cur = dict(line=m.group(1), item=m.group(3), short=[m.group(4)], goods=[m.group(6)], hs=m.group(5),
                               coo=m.group(7), qty=_num(m.group(8)), uom=m.group(9), price=_num(m.group(10)), amount=_num(m.group(14)))
                    d['lines'].append(cur)
                    hw = next((w for w in ws if w['text'] == m.group(5)), None)
                    cw = next((w for w in ws if hw and w['x0'] > hw['x0'] and w['text'] == m.group(7)), None)
                    if hw:
                        hts_x = hw['x0']
                    if cw:
                        coo_x = cw['x0']
                    continue
                if re.match(r'^(Grand Total|Total|Additional)', txt):
                    cur = None
                if cur and hts_x:
                    sh_ = ' '.join(w['text'] for w in ws if w['x0'] < hts_x - 2)
                    gd = ' '.join(w['text'] for w in ws if w['x0'] >= hts_x - 2 and (not coo_x or w['x0'] < coo_x - 2))
                    if sh_:
                        cur['short'].append(sh_)
                    if gd:
                        cur['goods'].append(gd)
    out = []
    for l in d['lines']:
        goods = re.sub(r'(\w)-\s+', r'\1-', ' '.join(l['goods'])).upper()
        goods = re.sub(r'\s+', ' ', goods).strip()
        short = re.sub(r'\s+', ' ', ' '.join(l['short'])).upper().strip()
        out.append(_line(line=l['line'], item=l['item'], desc='%s, %s' % (goods, short) if short else goods, hs=_hs8(l['hs']),
                         coo=l['coo'], qty=l['qty'], uom={'PIECE': 'PCE', 'PIECES': 'PCE'}.get(l['uom'], l['uom']),
                         price=l['price'], amount=l['amount'], order=l['line'], currency=d['currency']))
    d['lines'] = out
    d = _finish(d, C)
    d['notes'] = [n for n in d['notes'] if not n.startswith('no net weight on the invoice')]     # weights: packing list
    return [d]


def vf_packing_xls(path, rows, C):
    """VF Infor Nexus PackingList (.xls / .xlsx): PO SUMMARY totals + PACKAGE DETAIL per carton (main line #, qty,
    net net / net / gross KG) -> weights per invoice PO line."""
    import readers as R
    d = R.new_pl(path, 'VF packing list (Infor Nexus)')
    txt = lambda r: [str(v).strip() if v is not None else '' for v in r]
    hdr, hi = None, -1
    for i, r in enumerate(rows):
        c = txt(r)
        if c and c[0] == 'Ref Number' and len(c) > 1:
            d['invoice_refs'] = [c[1]]
        if 'Total Package Count' in c:
            k = c.index('Total Package Count')
            for v in c[k + 1:]:
                if v.replace('.0', '').isdigit():
                    d['packages'], d['package_unit'] = int(float(v)), 'CTN'
                    break
        if c and c[0] == 'Totals' and len([v for v in c if v]) > 5 and d['total_gw'] is None:
            nums = [_num(v) for v in c[1:] if _num(v) is not None]
            if len(nums) >= 5:
                d['total_nw'], d['total_gw'] = nums[3], nums[4]
        if 'Main Line #' in c and 'Item Qty' in c:
            hdr, hi = c, i
    if hdr:
        col = {h: k for k, h in enumerate(hdr) if h}
        rc, nc, gc = col.get('Range'), col.get('Net'), col.get('Gross')
        agg = collections.OrderedDict()
        for r in rows[hi + 1:]:
            c = txt(r) + [''] * 5
            if rc is None or not c[rc].isdigit():
                if agg:
                    break
                continue
            ml = c[col['Main Line #']]
            a = agg.setdefault(ml, dict(item=c[col.get('Buyer Item #', 0)], qty=0.0, nw=0.0, gw=0.0, n=0))
            a['qty'] += _num(c[col['Item Qty']]) or 0
            a['nw'] += (_num(c[nc]) or 0) if nc is not None else 0
            a['gw'] += (_num(c[gc]) or 0) if gc is not None else 0
            a['n'] += 1
        for ml, a in agg.items():
            d['lines'].append(_line(line=ml, item=a['item'], qty=a['qty'], uom='PCE', nw=round(a['nw'], 3), gw=round(a['gw'], 3),
                                    order=ml, boxes=['%d carton(s)' % a['n']]))
    if not d['lines']:
        d['notes'].append('no PACKAGE DETAIL rows read')
    return [d]


# ======================================================================== Shell Eastern Trading (lubricants export)
def shell_invoice(path, pages, C):
    """Shell Eastern Trading Tax Invoice (export, e.g. to Shell Tianjin): one line per delivery (container) with
    price build-up. Item value = Net Price - Second Distrib Cost (rule given for the Shell permits)."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Shell invoice', 'SHELL EASTERN TRADING (PTE) LTD')
    d['title'] = 'TAX INVOICE'
    d['ship_from_country'], d['direction'] = 'SG', 'EXPORT'
    m = re.search(r'Tax Invoice Number\s+Date.*\n\s*(\d{8,})\s+(\d{2}\.\d{2}\.\d{4})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _date_any(m.group(2))
    m = re.search(r'\bUEN\s+(\d{8,9}[A-Z])\b', t)
    d['uen'] = m.group(1) if m else ''
    m = re.search(r'Order Number\s+(\d{6,})', t)
    d['order_no'] = m.group(1) if m else ''
    m = re.search(r'^\d{6}\s+(SHELL\s.+?)$\n(?:.*\n)?^(LIMITED|LTD\.?|CO\.?,? LTD\.?)\s*$', t, re.M)
    d['ship_to'] = d['invoice_to'] = (clean(m.group(1)) + ' ' + m.group(2)).upper() if m else ''
    m = re.search(r'Delivery To\s+\S+\s+.*?,(CHINA|[A-Z ]+?)\s', t)
    d['ship_to_country'] = C.code(m.group(1)) if m else ''
    m = re.search(r'PALLET WEIGHT\s*(' + N + r')\s*KG', t, re.I)
    pallet = _num(m.group(1)) if m else None
    LN = re.compile(r'^(\d{6})\s+(.+?)\s+(' + N + r')\s+([A-Z]{2,3})\s+(' + N + r')\s+([A-Z]{3})\s*/\s*(\d+)\s*([A-Z]{1,3})\s+(' + N + r')\s*$')
    cur, deliv, state = None, '', ''
    for raw in t.splitlines():
        ln = raw.strip()
        md = re.search(r'Delivery Number/Date\s+(\d+)', ln)
        if md:
            deliv = md.group(1)
        m = LN.match(ln)
        if m:
            cur = dict(line=m.group(1), desc=clean(m.group(2)), qty=_num(m.group(3)), uom=m.group(4), currency=m.group(6),
                       order=deliv, nw=None, gw=None, coo='', inc='', second=0.0, net=None, base=_num(m.group(9)))
            d['lines'].append(cur)
            state = 'qty'
            continue
        if not cur:
            continue
        m = re.match(r'^([A-Z]{3})\s+(.+?)\s+([A-Z]{2})\s+(' + N + r')\s*KG$', ln)
        if m and state in ('qty', 'kg'):
            cur['inc'], cur['place'], cur['coo'], cur['nw'] = m.group(1), m.group(2), m.group(3), _num(m.group(4))
            state = 'gw'
            continue
        m = re.match(r'^(' + N + r')\s*KG$', ln)
        if m and state == 'gw':
            cur['gw'] = _num(m.group(1))
            state = 'cost'
            continue
        if m and state == 'qty':
            state = 'kg'
            continue
        m = re.match(r'^Second Distrib Cost\b.*?(' + N + r')$', ln)
        if m:
            cur['second'] = _num(m.group(1)) or 0.0
            continue
        m = re.match(r'^Net Price\s+(' + N + r')\s+([A-Z]{3}).*?\s(' + N + r')$', ln)
        if m:
            cur['net'] = _num(m.group(3))
            cur = None
            state = ''
    out, total_net = [], 0.0
    raw_lines = d['lines']
    for l in raw_lines:
        net = l['net'] if l['net'] is not None else l['base']
        total_net += net or 0
        amt = round((net or 0) - (l['second'] or 0), 2)
        gw = round(l['gw'] + pallet, 3) if l['gw'] is not None and pallet else l['gw']
        out.append(_line(line=l['line'], item=l['desc'], desc=l['desc'], qty=l['qty'], uom=l['uom'], amount=amt,
                         price=round(amt / l['qty'], 4) if l['qty'] else None, coo=l['coo'], order=l['order'], nw=l['nw'],
                         gw=gw, currency=l['currency'], delivery=l['order']))
        if l['net'] is None:
            d['notes'].append('line %s: Net Price not found - amount %s used' % (l['line'], l['base']))
    d['lines'] = out
    d['currency'] = out[0]['currency'] if out else ''
    incs = [l['inc'] for l in raw_lines if l.get('inc')]
    d['incoterm'] = incs[0] if incs else ''
    m = re.search(r'Total Amount \(Excl GST\)\s+([A-Z]{3})\s+(' + N + r')', t)
    inv_total = _num(m.group(2)) if m else None
    d['total'] = round(sum(l['amount'] or 0 for l in out), 2)
    sec = round(sum(l['second'] or 0 for l in raw_lines), 2)
    if inv_total is not None:
        d['notes'].append('item value = Net Price - Second Distrib Cost: invoice %s %s - %s = %s'
                          % (d['currency'], '{:,.2f}'.format(inv_total), '{:,.2f}'.format(sec), '{:,.2f}'.format(d['total'])))
    if pallet:
        d['notes'].append('gross weight per line + pallet weight %s KG (gross weight with pallets)' % fmt_kg(pallet))
    d['total_nw'] = round(sum(l['nw'] or 0 for l in out), 3) or None
    d['total_gw'] = round(sum(l['gw'] or 0 for l in out), 3) or None
    d['pl_first'] = True                # Shell rule: net weight + HS code from the delivery note (packing list)
    d = _finish(d, C)
    d['notes'] = [n for n in d['notes'] if not n.startswith('no HS / tariff code')]     # HS: delivery note / B/L
    return [d]


def fmt_kg(v):
    return ('%.3f' % v).rstrip('0').rstrip('.') if v is not None else ''


def shell_dn(path, pages, C):
    """Shell DELIVERY NOTE (one per container): product, qty, gross / net weight, tariff no., container, seal, pallets."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Shell delivery note')
    m = re.search(r'!\s*(\d{8,})\s*!\s*(\d{6,})\s*!', t)
    deliv = m.group(1) if m else ''
    d['delivery'] = deliv
    m = re.search(r'Vehicle / Vessel ID\s+([A-Z]{4}\d{7})', t)
    cont = m.group(1) if m else ''
    m = re.search(r'SEAL NO\.?\s*([A-Z0-9/ ,]+?)\s*$', t, re.M)
    seal = clean(m.group(1)) if m else ''
    m = re.search(r'PALLET QTY\s*:\s*(\d+)\s*PALLET', t)
    pallets = int(m.group(1)) if m else None
    m = re.search(r'PALLET WEIGHT\s*:?\s*(' + N + r')\s*KG', t)
    pw = _num(m.group(1)) if m else 0.0
    m = re.search(r'^(\d{6,})\s+(.+?)\s+(' + N + r')\s+([A-Z]{2,3})\s+(' + N + r')\s+([A-Z]{2,3})\s+(' + N + r')\s*KG', t, re.M)
    if m:
        nw = re.search(r'^\.\s*(' + N + r')\s*KG', t[m.end():], re.M)
        hs = re.search(r'Tariff No:\s*(\d{6,10})', t)
        d['lines'].append(_line(item=m.group(1), desc=clean(m.group(2)), qty=_num(m.group(5)), uom=m.group(6),
                                gw=round((_num(m.group(7)) or 0) + (pw or 0), 3), nw=_num(nw.group(1)) if nw else None,
                                hs=_hs8(hs.group(1)) if hs else '', order=deliv, delivery=deliv))
    m = re.search(r'Total Gross\s+(' + N + r')\s*KG', t)
    d['total_gw'] = round((_num(m.group(1)) or 0) + (pw or 0), 3) if m else None
    m = re.search(r'Total Net\s+(' + N + r')\s*KG', t)
    d['total_nw'] = _num(m.group(1)) if m else None
    if pallets:
        d['packages'], d['package_unit'] = pallets, 'PAT'
    if cont:
        d['containers'] = [cont]
    d['seal'] = seal
    if pw:
        d['notes'].append('gross weight %s KG + pallets %s KG (gross weight with pallets)' % (fmt_kg((d['total_gw'] or 0) - pw), fmt_kg(pw)))
    return [d]


def danmar_bl(path, pages, C):
    """DHL / Danmar Lines EXPRESS SEA WAYBILL or bill of lading: B.L. No., vessel / voyage, ports, containers with seals,
    total gross weight, HS per product."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'Danmar sea waybill', 4)
    cg = {'mode': 'Sea'}
    m = re.search(r'B\.L\. No\.\s*\n(.+)\n', t)
    if m:
        toks = re.findall(r'\b([A-Z]{3,5}\d{5,12})\b', m.group(1))
        if toks:
            d['hawb'] = cg['hawb'] = toks[-1]
        mm = re.search(r'\b(S\d{8,})\b', m.group(1))
        if mm:
            d['fwd_job'] = mm.group(1)
    m = re.search(r'^\s*([A-Z][A-Z0-9 .\-]+?)\s+([A-Z]?\d{1,4}[A-Z]{0,2})\s+upon proof', t, re.M)
    if m:
        cg['transport_id'], cg['flight'] = clean(m.group(1)).upper(), m.group(2)
    m = re.search(r'Place of receipt\s+Port of loading\s*\n\s*([A-Z][A-Z .]+?)\s+([A-Z][A-Z .]+?)(?:\s{2,}|\s+DHL|\s+For |$)', t, re.M)
    if m and not re.match(r'DHL|FOR\b', m.group(1)):
        cg['loading_name'] = clean(m.group(2)).upper()
    else:                                   # "For the release of goods apply to:" agent printed on the line in between
        m = re.search(r'Place of receipt\s+Port of loading.*\n(?:.*\n){0,2}?\s*([A-Z][A-Z .]+?)\s+\1\b', t)
        if m:
            cg['loading_name'] = clean(m.group(1)).upper()
    m = re.search(r'Port of discharge\s+Place of delivery.*\n(?:.*\n){0,3}?\s*([A-Z][A-Z .]+?,\s*[A-Z][A-Z ]+?)\s+\1', t)
    if m:
        cg['discharge_name'] = clean(m.group(1).split(',')[0]).upper()
    conts = []
    for mc in re.finditer(r'^([A-Z]{4}\d{7})\s+(.+?)\s+(\d{2}[A-Z]{2})\s+(' + N + r')\s*KG\s+(' + N + r')\s*M3\s+(\d+)\s+([A-Z]+)\s+(\S+)', t, re.M):
        size = mc.group(3)[:2]
        conts.append(dict(no=mc.group(1), seal=clean(mc.group(2)).rstrip(','), size=size,
                          size_text=('LCL' if 'CFS' in mc.group(8) else 'FCL') + size,
                          tne=(_num(mc.group(4)) or 0) / 1000.0 or None, type=mc.group(3)))
    if conts:
        cg['containers'] = [c['no'] for c in conts]
        cg['container_info'] = conts
    m = re.search(r'(\d+)\s*x\s*\d{2}[A-Z]{2}\s+CONTAINER\s+(' + N + r')\s*KG', t)
    if m:
        cg['gw_kg'] = d['gw'] = _num(m.group(2))
    m = re.search(r'STC\s+(\d+)\s+Package', t)
    if m:
        d['pieces'] = int(m.group(1))
    m = re.search(r'(\d+)\s+(Package|Pallet|Carton|Case|Crate|Drum|Bag|Box|Skid)\(s\).*?(' + N + r')\s*KG\s+(' + N + r')\s*M3', t)
    if m:                                   # LCL: "4 Package(s) ... 1165.000 KG 1.367 M3" / "1 Pallet(s) A 257.520 KG 0.629 M3"
        d['pieces'] = d['pieces'] or int(m.group(1))
        if not cg.get('gw_kg'):
            cg['gw_kg'] = d['gw'] = _num(m.group(3))
        cg.setdefault('packages_qty', int(m.group(1)))
        cg.setdefault('packages_unit', {'PALLET': 'PAT', 'CARTON': 'CTN', 'BAG': 'BAG', 'DRUM': 'DRM'}.get(m.group(2).upper(), 'PKG'))
    if re.search(r'\bCFS/CFS\b', t) and not conts:
        cg['cargo_lcl'] = True              # LCL, no container of our own
    hsd = {}
    for mh in re.finditer(r'^(.+)\nHS CODE:\s*(\d{6,10})', t, re.M):
        hsd.setdefault(clean(mh.group(1)), _hs8(mh.group(2)))
    d['hs_by_desc'] = hsd
    m = re.search(r'Place and Date of issue.*\n.*?(\d{1,2}-[A-Z]{3}-\d{2,4})', t)
    if m:
        d['executed'] = _dmon(m.group(1))
    d['issued_by'] = 'DHL GLOBAL FORWARDING (SINGAPORE) PTE LTD' if 'DHL GLOBAL FORWARDING (SINGAPORE)' in t else ''
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]



# ======================================================================== RS Components (UK) consolidated invoice (Excel) -> RS Components Pte Ltd
def rs_invoice_xls(path, rows, C):
    """RS Commercial Invoice (consolidated, .xlsx / .xls): Invoice No / Date / Invoice Amount (incl. packing charges) /
    Value Of Goods, one row per line: Line No, Product Code, Item Description, Handling Unit, Qty, UOM, Unit Price,
    Total Amount (+ charges), CCCN Number (HS), Country of Origin, Gross / Nett weight, Sales Invoice Number.
    Item value = "Total Amount including charges" (the invoice amount = all lines incl. the packing charge)."""
    d = _inv(path, 'RS Components consolidated invoice', 'RS COMPONENTS LIMITED')
    d.update(ship_from_country='GB', title='COMMERCIAL INVOICE', invoice_from='RS COMPONENTS LIMITED')
    txt = lambda v: '' if v is None else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip())
    hdr, hi = None, -1
    for i, r in enumerate(rows):
        c = [txt(v) for v in r]
        if not any(c):
            continue
        lab = c[0].lower()
        val = next((v for v in c[1:4] if v), '')
        if lab == 'invoice no':
            d['invoice_no'] = val
        elif lab == 'invoice date':
            d['invoice_date'] = _date_any(val) if not isinstance(r[2], datetime.datetime) else r[2].strftime('%d/%m/%Y')
        elif lab == 'invoice amount':
            d['total'] = _num(val)
            cur = [v for v in c[1:6] if re.fullmatch(r'[A-Z]{3}', v)]
            if cur:
                d['currency'] = cur[0]
        elif lab == 'value of goods':
            d['goods_value'] = _num(val)
        if 'Consignee' in c:
            k = c.index('Consignee')
            d['ship_to'] = d['invoice_to'] = (c[k + 1] if k + 1 < len(c) else '').upper()
        if c[0] == 'Line No' and 'Product Code' in c:
            hdr, hi = c, i
            break
    d['ship_to'] = d.get('ship_to') or 'RS COMPONENTS PTE LTD'
    d['ship_to_country'] = 'SG'
    if not hdr:
        d['notes'].append('item table (Line No / Product Code ...) not found')
        return [_finish(d, C)]
    col = {}
    for k, h in enumerate(hdr):
        if h and h not in col:
            col[h] = k
    g = lambda r, h: r[col[h]] if h in col and col[h] < len(r) else None
    # the weight value columns: "Gross weight" / "Nett weight" (each followed by its "Unit" column)
    cnt = collections.Counter()
    for r in rows[hi + 1:]:
        no = g(r, 'Line No')
        if not isinstance(no, (int, float)) or g(r, 'Product Code') in (None, ''):
            continue
        amt_c = _num(txt(g(r, 'Total Amount including charges')))
        amt = _num(txt(g(r, 'Total Amount')))
        hs = re.sub(r'\D', '', txt(g(r, 'CCCN Number')))
        l = _line(line=str(int(no)), item=txt(g(r, 'Product Code')), desc=clean(txt(g(r, 'Item Description'))),
                  qty=_num(txt(g(r, 'Qty'))), uom=txt(g(r, 'Unit of Measure')).upper(), price=_num(txt(g(r, 'Unit Price'))),
                  amount=amt_c if amt_c is not None else amt, currency=d.get('currency') or '',
                  hs=hs.zfill(8)[:8] if hs else '', coo=txt(g(r, 'Country of Origin')).upper(),
                  gw=_num(txt(g(r, 'Gross weight'))), nw=_num(txt(g(r, 'Nett weight') if 'Nett weight' in col else g(r, 'Net weight'))))
        l['sales_inv'] = txt(g(r, 'Sales Invoice Number')).split('.')[0]
        l['handling_unit'] = txt(g(r, 'Handling Unit Number'))
        l['amount_goods'] = amt
        if l['qty'] and l['amount'] is not None:
            l['price'] = round(l['amount'] / l['qty'], 4)      # unit price incl. the packing charge
        d['lines'].append(l)
        cnt[l['handling_unit']] += 1
    hu = [h for h in cnt if h]
    if hu:
        d['handling_units'] = d['boxes'] = len(hu)             # one handling unit = one package
    goods = round(sum(l['amount_goods'] or 0 for l in d['lines']), 2)
    if d.get('goods_value') is not None and abs(goods - d['goods_value']) > 0.05:
        d['notes'].append('line Total Amount adds up to %s, Value Of Goods %s' % ('{:,.2f}'.format(goods), '{:,.2f}'.format(d['goods_value'])))
    d['sales_invoices'] = list(dict.fromkeys(l['sales_inv'] for l in d['lines'] if l['sales_inv']))
    d['notes'].append('item value = Total Amount including charges (packing charge %s added to the goods %s)'
                      % ('{:,.2f}'.format(round(sum(l['amount'] or 0 for l in d['lines']) - goods, 2)), '{:,.2f}'.format(goods)))
    return [_finish(d, C)]

# ======================================================================== Clasquin - Millesima wine & spirits (air, Bordeaux)
def _fr_num(s):
    """'1201,27' / '9752,88' / '159.54' -> float."""
    s = (s or '').strip().replace(' ', '')
    if re.fullmatch(r'\d+,\d+', s):
        s = s.replace(',', '.')
    return _num(s)


def _liquor_line(name, appellation, eu_hs, abv, liters, bottle_l, bottles, doc_country):
    import liquor_rules as L
    return dict(kind=L.kind_of(name, appellation, eu_hs), colour=L.colour_of(appellation, name), abv=abv, bottle_l=bottle_l,
                bottles=bottles, liters=liters, name=name, appellation=appellation, eu_hs=eu_hs,
                country_hint=L.country_hint(' '.join((name, appellation))), doc_country=doc_country)


def liquor_desc(l):
    """Permit description of a liquor line: NAME - RED WINE - 8 BTL X 0.75L - 12% VOL."""
    q = l.get('liquor') or {}
    kind = {'WINE': '%s WINE' % (q.get('colour') or '').replace('ROSE', 'ROSE'), 'SPARKLING': 'SPARKLING WINE',
            'FORTIFIED': 'FORTIFIED WINE'}.get(q.get('kind'), q.get('kind') or '')
    parts = [clean(q.get('name') or l.get('desc') or '').upper(), clean(kind).upper()]
    if q.get('bottles') and q.get('bottle_l'):
        parts.append('%d BTL X %gL' % (q['bottles'], q['bottle_l']))
    if q.get('abv') is not None:
        parts.append('%g%% VOL' % q['abv'])
    return ' - '.join(p for p in parts if p)


def millesima_docs(path, pages, C):
    """Clasquin import of MILLESIMA (Bordeaux) wines & spirits - one PDF: house AWB (page 1) + invoice 'DAE Line | GOODS |
    QUANTITY | UNIT PRICE | Total' (each wine: name (FRANCE), appellation - colour, 'code-Degree 12 , Liter 6 , Weight
    14.4 Kg , HSCode 22042112') + excise / CIVC document. The HSCode printed is the EU code - the Singapore HS + product
    code come from the liquor rules (liquor_rules.py)."""
    import permit_engine as E
    docs = []
    first_inv = next((i for i, p in enumerate(pages) if 'DAE Line' in p['text']), len(pages))
    awb_pages = [p for p in pages[:first_inv] if re.search(r'HAWB\s*No|Air\s*Waybill|Shipper.s Name', p['text'], re.I)]
    if awb_pages:
        a = E.parse_awb(path, awb_pages)
        t = '\n'.join(p['text'] for p in awb_pages)
        a['reader'] = 'Clasquin house AWB'
        if not a.get('mawb'):
            m = re.search(r'\b(\d{3})-(\d{8})\b', t) or re.search(r'^\s*(\d{3})[A-Z]{3}\s?(\d{8})\b', t, re.M)
            if m:
                a['mawb'] = m.group(1) + m.group(2)
        if a.get('mawb'):
            a.setdefault('cargo', {})['mawb'] = a['mawb']
        m = re.search(r'TERMS\s*:\s*([A-Z]{3})\b', t)
        if m:
            a['awb_terms'] = m.group(1)
        if not a.get('issued_by'):
            a['issued_by'] = 'CLASQUIN'
        a['consignee'] = a.get('consignee') or 'CLASQUIN SINGAPORE PTE LTD'
        docs.append(a)
    inv_pages = [p for p in pages if 'DAE Line' in p['text']]
    t = '\n'.join(p['text'] for p in inv_pages)
    d = _inv(path, 'Millesima invoice', 'MILLESIMA')
    d.update(title='INVOICE', ship_from_country='FR', ship_to_country='SG', direction='IMPORT',
             invoice_to='CLASQUIN SINGAPORE PTE LTD', ship_to='CLASQUIN SINGAPORE PTE LTD')
    m = re.search(r'Invoice\s*N[°o]\s*(\S+)', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'\bDate\s+(\d{2})/(\d{2})/(\d{4})', t)
    d['invoice_date'] = _dmy(*m.groups()) if m else ''
    m = re.search(r'TOTAL EX\.?\s*VAT\s+([A-Z]{3})\s+([\d ,.]+)', t)
    d['currency'] = m.group(1) if m else ''
    d['total'] = _fr_num(m.group(2)) if m else None
    m = re.search(r'TOTAL GROSS WEIGHT\s+([\d.,]+)\s*Kg', t, re.I)
    d['total_gw'] = _fr_num(m.group(1)) if m else None
    m = re.search(r'Status\s+(DUTY PAID|[A-Z ]+)', t)
    d['duty_status'] = m.group(1).strip() if m else ''
    if awb_pages:
        d['incoterm'] = docs[0].get('awb_terms', '')
        d['hawb'] = docs[0].get('hawb', '')
    LN = re.compile(r'^(\d{1,3})\s+(.+?)\s+(\d+)\s+((?:\d+(?=BTL))?(?:\d+\s*CL\s+)?(?:BTL CASE|BTL|MAGNUM|JEROBOAM|BOUTEILLE|BOTTLE|'
                    r'CASE)\b[^\d\n]*?(?:\(?\s*[\d.]+\s*L\)?)?)\s+(\d[\d ]*,\d+)\s+(\d[\d,]*\.\d{2})\s*$')
    DT = re.compile(r'^(\S+?)-Degree\s*([\d.]+)\s*°?\s*,\s*Liter\s*([\d.]+)\s*,\s*Weight\s*([\d.]+)\s*Kg\s*,\s*HSCode\s*(\d+)', re.I)
    cur, extra = None, []
    def close():
        if cur is None:
            return
        rest = []
        for x in extra:                     # name wrapped: "Quarts 2023 (FRANCE)" belongs to the name
            mc = re.search(r'\(\s*([A-Z ]{3,})\s*\)\s*$', x)
            if mc and ' - ' not in x:
                cur['name'] = clean(cur['name'] + ' ' + x[:mc.start()])
                cur['country'] = cur.get('country') or mc.group(1).strip()
            else:
                rest.append(x)
        cur['_app'] = next((x for x in rest if ' - ' in x), rest[0] if rest else '')
    for raw in t.splitlines():
        ln = raw.strip()
        m = LN.match(ln)
        if m:
            close()
            name = m.group(2)
            ctry = re.search(r'\(\s*([A-Z ]{3,})\s*\)\s*$', name)
            cur = dict(line=m.group(1), name=clean(re.sub(r'\(\s*[A-Z ]{3,}\s*\)\s*$', '', name)), qty=int(m.group(3)),
                       unit=m.group(4), price=_fr_num(m.group(5)), amount=_fr_num(m.group(6)),
                       country=ctry.group(1).strip() if ctry else '')
            extra = []
            d['lines'].append(cur)
            continue
        if cur is None:
            continue
        m = DT.match(ln)
        if m:
            cur.update(code=m.group(1), abv=float(m.group(2)), liters=float(m.group(3)), weight=float(m.group(4)), eu_hs=m.group(5))
            close()
            cur = None
            continue
        if re.match(r'(PAGE \d|TOTAL|Domiciliation|S\.A\. au capital)', ln):
            continue
        if re.fullmatch(r'\(\s*([A-Z ]{3,})\s*\)', ln):         # "(FRANCE)" wrapped to its own line
            cur['country'] = cur.get('country') or ln.strip('() ')
            continue
        extra.append(ln)
    out = []
    for x in d['lines']:
        if 'liters' not in x:
            d['notes'].append('CHECK line %s %s: alcohol degree / litres line not read' % (x['line'], x['name']))
        mb = re.search(r'([\d.]+)\s*L\)?\s*$', x['unit'] or '')
        bottle = float(mb.group(1)) if mb else None
        liters = x.get('liters')
        bottles = int(round(liters / bottle)) if liters and bottle else x['qty']
        if liters and bottle and abs(bottles * bottle - liters) > 0.01:
            d['notes'].append('CHECK line %s %s: %s L is not a whole number of %s L bottles' % (x['line'], x['name'], liters, bottle))
        l = _line(line=x['line'], item=x.get('code', ''), desc=x['name'], qty=float(bottles), uom='BTL', price=x['price'],
                  amount=x['amount'], hs='', coo=C.code(x['country']) if x.get('country') else '', gw=x.get('weight'))
        l['ltr'] = liters
        l['eu_hs'] = x.get('eu_hs', '')
        l['liquor'] = _liquor_line(x['name'], x.get('_app', ''), x.get('eu_hs', ''), x.get('abv'), liters, bottle, bottles,
                                   x.get('country', ''))
        l['desc'] = liquor_desc(l)
        out.append(l)
    d['lines'] = out
    # the invoice summary: Wine 96 L / Champagne 9.3 L / Whiskey 0.7 L
    summ = {k.upper(): _fr_num(v) for k, v in re.findall(r'^(Wine|Champagne|Whiske?y|Spirits?|Sparkling)\s+([\d.,]+)\s*L\b', t, re.M | re.I)}
    if summ:
        grp = {'WINE': 0.0, 'CHAMPAGNE': 0.0, 'WHISKEY': 0.0}
        for l in out:
            k = l['liquor']['kind']
            grp['CHAMPAGNE' if k == 'SPARKLING' else 'WHISKEY' if k == 'WHISKY' else 'WINE'] += l['ltr'] or 0
        for k, v in summ.items():
            k2 = 'WHISKEY' if k.startswith('WHISK') else k
            if k2 in grp and v is not None and abs(grp[k2] - v) > 0.01:
                d['notes'].append('MISMATCH invoice summary %s %s L, lines add up to %s L' % (k, v, round(grp[k2], 3)))
        d['liquor_summary'] = summ
    if d['total'] is not None and abs(sum(l['amount'] or 0 for l in out) - d['total']) > 0.05:
        d['notes'].append('MISMATCH line amounts %.2f <> invoice total %.2f' % (sum(l['amount'] or 0 for l in out), d['total']))
    d = _finish(d, C)
    d['notes'] = [n for n in d['notes'] if not re.match(r'no (HS / tariff code|net weight)', n)]
    docs.append(d)
    return docs


def clasquin_detail_xls(path, rows, C):
    """Clasquin 'detail_hscode' Excel (Line | PickId | Supplier | Supplier Invoice number | Commodity Code | Description |
    Unit Value | Total Value | Currency | Net KG/Litres | Country of Origin | Total Qnty | Bottle Size in Litres | Alcohol
    Percentage ...). Same invoice as the Millesima PDF: used to check it (or as the invoice when the PDF is missing)."""
    import readers as R
    hdr = [str(x or '').strip() for x in rows[0]]
    ix = {h: i for i, h in enumerate(hdr)}
    g = lambda r, k: r[ix[k]] if k in ix and ix[k] < len(r) else None
    d = _inv(path, 'Clasquin detail (HS codes)', '')
    d.update(title='INVOICE DETAIL', ship_to_country='SG', direction='IMPORT', detail_copy=True)
    for r in rows[1:]:
        if not isinstance(g(r, 'Line'), (int, float)):
            continue
        num = lambda v: _fr_num(str(v)) if v not in (None, '') else None
        name = clean(str(g(r, 'Description') or ''))
        liters, bottle, abv = num(g(r, 'Net KG/Litres')), num(g(r, 'Bottle Size in Litres')), num(g(r, 'Alcohol Percentage'))
        bottles = num(g(r, 'Total Qnty'))
        ctry = str(g(r, 'Country of Origin') or '').strip()
        l = _line(line=str(int(g(r, 'Line'))), item=re.sub(r'-\d+$', '', str(g(r, 'PickId') or '')), desc=name, qty=bottles,
                  uom='BTL', price=num(g(r, 'Unit Value')), amount=num(g(r, 'Total Value')), hs='',
                  coo=C.code(ctry) if ctry else '')
        l['currency'] = str(g(r, 'Currency') or '')
        l['ltr'] = liters
        l['eu_hs'] = str(g(r, 'Commodity Code') or '')
        l['liquor'] = _liquor_line(name, '', l['eu_hs'], abv, liters, bottle, int(bottles or 0), ctry)
        l['liquor']['name'] = clean(re.sub(r'\s*\((?:b|r|rs|rose|)\)\s*-?\s*', ' ', name, flags=re.I))   # colour marks out
        l['desc'] = liquor_desc(l)
        d['lines'].append(l)
        d['invoice_no'] = d['invoice_no'] or str(g(r, 'Supplier Invoice number') or '')
        d['currency'] = d['currency'] or l['currency']
        if not d.get('invoice_to'):
            d['invoice_to'] = d['ship_to'] = re.sub(r'^SG\s*/\s*', '', str(g(r, 'Supplier') or '')).upper()
    d['total'] = round(sum(l['amount'] or 0 for l in d['lines']), 2)
    d = _finish(d, C)
    d['notes'] = [n for n in d['notes'] if not re.match(r'no (HS / tariff code|net weight)', n)]
    return [d]



EXCEL_READERS = (
    ('CLASQUIN DETAIL HS', lambda t: 'PickId' in t and 'Supplier Invoice number' in t and 'Alcohol Percentage' in t, clasquin_detail_xls),
    ('RS COMPONENTS INVOICE', lambda t: 'RS Commercial Invoice' in t and 'CCCN Number' in t and 'Sales Invoice Number' in t, rs_invoice_xls),
    ('SHARKNINJA PACKING LIST', lambda t: 'SN APAC Holding' in t and re.search(r'PACKING\s+LIST', t) and 'Item Code' in t, lambda p, r, C: sharkninja_pl_xls(p, r, C)),
    ('SHANGHAI JIANLAN PACKING LIST', lambda t: 'Shanghai Jianlan Electric' in t and 'PACKING LIST' in t and 'NET WEIGHT' in t.upper(), lambda p, r, C: jianlan_pl_xls(p, r, C)),
    ('VERTIV CUSTOMS INVOICE', lambda t: re.search(r'Vertiv Tech', t, re.I) and re.search(r'CUSTOMS INVOICE|COMMERCIAL INVOICE', t) and 'Part No' in t, lambda p, r, C: vertiv_invoice_xls(p, r, C)),
    ('VERTIV PACKING LIST', lambda t: re.search(r'Vertiv Tech', t, re.I) and 'PACKING LIST' in t and 'Package No' in t, lambda p, r, C: vertiv_pl_xls(p, r, C)),
    ('VF PACKING LIST', lambda t: 'PACKAGE DETAIL' in t and 'Main Line #' in t and 'PTS Ref Number' in t, vf_packing_xls),
)


def read_excel(path, rows, countries=None, log=None):
    """Known Excel layouts (e.g. VF / Infor Nexus packing list) -> [document]; None when not known."""
    text = '\n'.join(' | '.join(str(v) for v in r if v not in (None, '')) for r in rows[:400])
    for name, test, fn in EXCEL_READERS:
        try:
            ok = test(text)
        except Exception:
            ok = False
        if ok:
            docs = fn(path, rows, countries or Countries())
            for d in docs:
                d['reader'] = d.get('reader') or name
            if log:
                log('   %s: %s layout, %d line(s)' % (os.path.basename(path), name, sum(len(d.get('lines', [])) for d in docs)))
            return docs
    return None


# ======================================================================== Sandvik SMC Logistics (export from the Singapore hub)
SANDVIK_TOOL_HS = ('8207', '8202', '8203', '8204', '8205', '8206', '8466')
SANDVIK_TOOL_WORDS = re.compile(r'\b(BIT|BITS|BUTTON|DRILL ROD|ROD|RODS|SHANK|SHANK ADAPTER|COUPLING SLEEVE|TUBE ROD|PICK|PICKS|'
                                r'CUTTER|REAMER|INSERT|DTH|HAMMER|DRILL STEEL|TOOL|TOOLS|CHISEL)\b')


def sandvik_type(desc, hs):
    """TOOLS or PARTS for the Sandvik export description (rock tools / drilling tools = TOOLS, everything else = PARTS)."""
    hs = re.sub(r'\D', '', hs or '')
    if hs[:4] in SANDVIK_TOOL_HS:
        return 'TOOLS'
    if not hs and SANDVIK_TOOL_WORDS.search((desc or '').upper()):
        return 'TOOLS'
    return 'PARTS'


def sandvik_invoice(path, pages, C):
    """Sandvik Mining & Construction Logistics INVOICE (goods delivered from the Singapore hub) - EXPORT permit.
    Item description on the permit = 'TOOLS OR PARTS INV NO : ( <invoice> )' (see resolve)."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Sandvik invoice', 'SANDVIK SMC LOGISTICS LTD')
    d['title'] = 'INVOICE'
    m = re.search(r'Invoice Number\s+Order Number.*\n\s*(\d{6,})', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice Date\s+Customer Number\s*\n\s*(\d{4}-\d{2}-\d{2})', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'Consignee\s+Buyer\s*\n(.+)', t)
    if m:
        line = clean(m.group(1))
        half = len(line) // 2
        d['ship_to'] = d['invoice_to'] = (line[:half].strip() if line[:half].strip() == line[half:].strip() else line).upper()
    m = re.search(r'Country of Destination\s*\n.*?\b([A-Z][a-z]+(?: [A-Z][a-z]+)*)\s*$', t, re.M)
    d['ship_to_country'] = C.code(m.group(1)) if m else ''
    m = re.search(r'Country of Departure\s*\n.*?\b([A-Z][a-z]+)\s*$', t, re.M)
    d['ship_from_country'] = C.code(m.group(1)) if m else 'SG'
    d['direction'] = 'EXPORT'           # Sandvik from the Singapore hub: export permits only
    m = re.search(r'Terms of Delivery.*\n.*?\b(EXW|FCA|FOB|CFR|CIF|CPT|CIP|DAP|DPU|DDP)\b', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'Dispatch Method.*\n\s*(Sea|Air|Road|Rail|Courier)\b', t, re.I)
    d['transport'] = m.group(1).title() if m else ''
    m = re.search(r'Carrier\s+Despatch Note.*\n\s*(.+?(?:PTE\.? LTD|LTD|LIMITED))', t, re.I)
    d['carrier'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'^\s*(?:Line Item code.*\n)?\s*([A-Z]{3})\s+KG\s+KG\s+([A-Z]{3})\s*$', t, re.M)
    d['currency'] = m.group(2) if m else ''
    LN = re.compile(r'^\s*(\d{1,4})\s+([A-Z0-9][A-Z0-9-]{4,})\s+(' + N + r')\s+([A-Z]{1,4})\s+(' + N + r')\s+(' + N + r')\s+(' + N
                    + r')\s+(' + N + r')\s*$')
    rows = t.splitlines()
    cur = None
    for i, raw in enumerate(rows):
        m = LN.match(raw)
        if m:
            cur = _line(line=m.group(1), item=m.group(2), qty=_num(m.group(3)), uom=m.group(4), price=_num(m.group(5)),
                        nw=_num(m.group(6)), gw=_num(m.group(7)), amount=_num(m.group(8)), currency=d['currency'])
            cur['desc'] = clean(rows[i + 1]) if i + 1 < len(rows) else ''
            d['lines'].append(cur)
            continue
        if not cur:
            continue
        mc = re.search(r'Country of Origin\s+([A-Z]{2})\b', raw)
        if mc:
            cur['coo'] = mc.group(1)
        ml = re.search(r'Load ID\s+(\d+)', raw)
        if ml:
            d['delivery'] = ml.group(1)
        md = re.search(r'Dispatch note\s*/\s*(\d+)', raw)
        if md:
            cur['order'] = md.group(1)
        mh = re.search(r'Export HS Code\s+(\d{6,10})', raw)
        if mh:
            cur['hs'] = _hs8(mh.group(1))
        if re.match(r'\s*Total Net Weight', raw):
            cur = None
    m = re.search(r'Total Net Weight\s+(' + N + r')\s*KG', t)
    d['total_nw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Gross Weight\s+(' + N + r')\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Amount\s+(' + N + r')', t)
    d['total'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Packages\s+(\d+)', t)
    if m:
        d['boxes'], d['package_unit'] = int(m.group(1)), 'PKG'
    d['desc_rule'] = 'SANDVIK'          # permit description: TOOLS OR PARTS INV NO : ( invoice )
    nohs = [l['item'] for l in d['lines'] if not l['hs']]
    if nohs:
        d['notes'].append('no Export HS Code on the invoice for %s - type it on the Item tab (Save to product master '
                          'remembers it)' % ', '.join(nohs))
    d['notes'].append('Exporter: goods delivered from Sandvik SMC Logistics Ltd., Singapore hub - check the exporter name / UEN '
                      '(typed once, remembered)')
    return [_finish(d, C)]


def sandvik_pl(path, pages, C):
    """Sandvik 'Packing List' (dispatch note): packages with weight, then items per package (item, qty, weight, COO)."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Sandvik packing list')
    m = re.search(r'List number Order number Page\s*\n\s*(\d+)', t)
    if m:
        d['delivery'] = m.group(1)
    pk = re.findall(r'^\s*(\d+)\s+Generic\s+(\w+).*\n\s*Weight\s+(' + N + r')\s*KG', t, re.M)
    if pk:
        d['packages'] = len(pk)
        kinds = {x[1].upper() for x in pk}
        d['package_unit'] = 'CTN' if kinds == {'CARTON'} else 'PAT' if kinds == {'PALLET'} else 'PKG'
        d['total_gw'] = round(sum(_num(x[2]) or 0 for x in pk), 3)
    m = re.search(r'Total Packs:\s*(\d+)', t)
    if m:
        d['packages'] = int(m.group(1))
    agg = collections.OrderedDict()
    rows = t.splitlines()
    for i, raw in enumerate(rows):
        m = re.match(r'^\s*([A-Z]{2}\d{6,})\s+(.+?)\s+(' + N + r')\s*([A-Z]{1,3})\s+(' + N + r')\s*$', raw)
        if not m:
            continue
        a = agg.setdefault(m.group(1), dict(desc=clean(m.group(2)), qty=0.0, nw=0.0, uom=m.group(4), coo=''))
        a['qty'] += _num(m.group(3)) or 0
        a['nw'] += _num(m.group(5)) or 0
        nxt = rows[i + 1] if i + 1 < len(rows) else ''
        mc = re.search(r'Country of origin:\s*([A-Z]{2})', nxt)
        if mc:
            a['coo'] = mc.group(1)
    for k, a in agg.items():
        d['lines'].append(_line(item=k, desc=a['desc'], qty=a['qty'], uom=a['uom'], nw=round(a['nw'], 3), coo=a['coo']))
    return [d]


# ======================================================================== Sandvik TAX INVOICE layout (built in, was a plugin)
# INV_2120450 / INV_2123424: Tarrif + COO + weights per line; packing list (dispatch note); hub pickup "Bill of Lading"
def svt_is_invoice(t):
    return bool(re.search(r'TAX INVOICE', t) and re.search(r'Sandvik Mining (?:&|and) Construction', t)
                and re.search(r'\bTarrif\b', t) and re.search(r'Invoice number', t))


def svt_is_pl(t):
    return bool(re.search(r'^\s*Packing List\s*$', t, re.M) and re.search(r'Sandvik Mining (?:&|and) Construction Logistics', t)
                and re.search(r'Package includes \d+ part', t) and re.search(r'List number', t))


def svt_is_carrier_bol(t):
    return bool(re.search(r'^\s*Bill of Lading\s*$', t, re.M) and re.search(r'Ship From:', t) and re.search(r'Waybill ref:', t)
                and re.search(r'Sandvik Mining (?:&|and) Construction Logistics', t) and re.search(r'Load ID', t))


# ------------------------------------------------------------------ Sandvik warehouse "Bill of Lading" (pickup by the carrier)
def svt_carrier_bol(path, pages, C):
    """CNT... 'Bill of Lading' from the Sandvik hub to the forwarder: packages + weight only (the ocean B/L / AWB of the
    forwarder wins for every transport value - this one has the lowest priority)."""
    from export_docs import _awb
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'Sandvik carrier bill of lading', 1)
    cg = {}
    m = re.search(r'Carrier Service:\s*(OCEAN|SEA|AIR)', t, re.I)
    if m:
        cg['mode'] = 'Air' if m.group(1).upper() == 'AIR' else 'Sea'
    m = re.search(r'Waybill ref:\s*(\S+)', t)
    d['fwd_job'] = m.group(1) if m else ''
    m = re.search(r'Load ID\s+(\d+)', t)
    d['load_id'] = m.group(1) if m else ''
    m = re.search(r'^\s*(\d+)\s+(' + N + r')\s+[\d.]+\s*$', t, re.M)          # total row: "4 1165.50 1.367"
    if m:
        d['pieces'], d['gw'] = int(m.group(1)), _num(m.group(2))
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = d['pieces'], 'PKG', d['gw']
    d['notes'].append('Sandvik pickup bill of lading %s (load %s): %s packages, %s KG - the forwarder B/L / AWB wins'
                      % (d['fwd_job'], d['load_id'], d['pieces'], d['gw']))
    d['cargo'] = cg
    return d


# ------------------------------------------------------------------ TAX INVOICE
SVT_FOOT = re.compile(r'^(Sandvik Mining & Construction|Logistics Limited|Domicile:|VAT No:|\S.*Dockline|Dublin\b|This order is accepted)')


def svt_body(page_text):
    """Page text without the repeated page header (up to the contact person line) and the address footer."""
    rows = page_text.splitlines()
    start = 0
    for i, r in enumerate(rows):
        if r.startswith('Contact person'):
            start = i + 2                     # the next line is the person + customer reference
            break
    out = []
    for r in rows[start:]:
        if SVT_FOOT.match(r.strip()):
            break
        out.append(r)
    return out


def svt_invoice(path, pages, C):
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Sandvik tax invoice', 'SANDVIK SMC LOGISTICS LTD')
    d['title'] = 'TAX INVOICE'
    d['ship_from_country'], d['direction'] = 'SG', 'EXPORT'
    m = re.search(r'Invoice number.*\n\s*(\d{6,})', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Issue date.*\n\s*(\d{2}\.\d{2}\.\d{4})', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'^Consignee\b.*\n(.+)', t, re.M)
    if m:
        line = clean(m.group(1))
        half = len(line) // 2
        a, b = line[:half].strip(), line[half:].strip()
        d['ship_to'] = d['invoice_to'] = (a if a == b else line).upper()
    m = re.search(r'Country of destination\s*\n.*?([A-Z][A-Za-z]+(?: [A-Z][A-Za-z]+)*)\s*$', t, re.M)
    d['ship_to_country'] = (C.code(m.group(1)) or '') if m else ''
    m = re.search(r'Terms\s*of delivery.*\n\s*(EXW|FCA|FOB|CFR|CIF|CPT|CIP|DAP|DPU|DDP)\b\s*(.*)', t)
    if m:
        d['incoterm'], d['place'] = m.group(1), clean(m.group(2))
    m = re.search(r'Dispatch method\s+Carrier\s*\n\s*(Sea|Air|Road|Rail|Courier)\b\s*(.*)', t, re.I)
    if m:
        d['transport'] = m.group(1).title()
        d['carrier'] = clean(m.group(2)).upper()
    m = re.search(r'Description\s+([A-Z]{3})\s+([A-Z]{3})\s*$', t, re.M)
    d['currency'] = m.group(2) if m else ''
    m = re.search(r'Terms of payment.*\n\s*(.+)', t)
    d['payment_terms'] = clean(m.group(1)) if m else ''

    LN = re.compile(r'^\s*(\d{1,4})\s+([A-Z0-9][A-Z0-9-]{4,})\s+(' + N + r')\s+([A-Z]{1,4})\s+(' + N + r')\s+(' + N + r')\s*$')
    rows = []
    for p in pages:
        rows += svt_body(p['text'])
    cur, state = None, ''
    for raw in rows:
        r = raw.strip()
        m = LN.match(r)
        if m and not r.startswith('0.00'):
            cur = _line(line=m.group(1), item=m.group(2), qty=_num(m.group(3)), uom=m.group(4), price=_num(m.group(5)),
                        amount=_num(m.group(6)), currency=d['currency'])
            cur['desc_lines'] = []
            d['lines'].append(cur)
            state = 'desc'
            continue
        if not cur:
            continue
        if re.match(r'Total Net Weight', r):
            cur, state = None, ''
            continue
        mh = re.search(r'Tarr?iff?\s+(\d{6,10})', r)
        if mh:
            cur['hs'] = _hs8(mh.group(1))
            state = 'detail'
        mc = re.search(r'Country of Origin\s+([A-Z]{2})\b', r)
        if mc:
            cur['coo'] = mc.group(1)
        mw = re.search(r'Net weight\s+(' + N + r')\s*KG', r, re.I)
        if mw:
            cur['nw'] = _num(mw.group(1))
        mg = re.search(r'Gross Weight\s+(' + N + r')\s*KG', r, re.I)
        if mg:
            cur['gw'] = _num(mg.group(1))
        md = re.search(r'Dispatch note\s+(\d+)\s*/\s*(\d+)', r)
        if md:
            cur['order'] = md.group(2)
        ml = re.search(r'Load ID\s+(\d+)', r)
        if ml:
            d['delivery'] = ml.group(1)
        mp = re.search(r'Package Number\s+([\d ,]+?)\s+Packaging', r)
        if mp:
            cur['boxes'] = [x for x in re.split(r'[\s,]+', mp.group(1)) if x]
        if state == 'desc' and r and not re.match(r'MA reference', r):
            cur['desc_lines'].append(clean(r))
        elif re.match(r'MA reference', r):
            state = 'detail'
    for l in d['lines']:
        l['desc'] = ' '.join(l.pop('desc_lines', [])).strip()
    m = re.search(r'^\s*Total\s+(' + N + r')\s*$', t, re.M) or re.search(r'Goods total\s+(' + N + r')\s*$', t, re.M)
    d['total'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Net Weight\s+(' + N + r')\s*KG', t)
    d['total_nw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Gross Weight\s*\n?\s*(' + N + r')\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Packages\s+(\d+)', t)
    if m:
        d['boxes'], d['package_unit'] = int(m.group(1)), 'PKG'
    d['desc_rule'] = 'SANDVIK'                # permit description: TOOLS OR PARTS INV NO : ( invoice )
    for k, lab in (('nw', 'net'), ('gw', 'gross')):
        s = round(sum(l.get(k) or 0 for l in d['lines']), 3)
        tot = d.get('total_' + k)
        if tot is not None and abs(s - tot) > 0.01:
            d['notes'].append('%s weight of the lines %s KG but the invoice total is %s KG - check' % (lab, s, tot))
    nohs = [l['item'] for l in d['lines'] if not l['hs']]
    if nohs:
        d['notes'].append('no Tarrif (HS) on the invoice for %s - type it on the Item tab' % ', '.join(nohs))
    d['notes'].append('Exporter: goods delivered from Sandvik SMC Logistics Ltd., Singapore hub - check the exporter UEN '
                      '(typed once, remembered); HAWB / booking not on the invoice - type it')
    return _finish(d, C)


# ------------------------------------------------------------------ PACKING LIST (dispatch note)
def svt_packing_list(path, pages, C):
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Sandvik packing list')
    m = re.search(r'List number.*\n\s*(\d+)', t)
    if m:
        d['delivery'] = m.group(1)
    m = re.search(r'Country of destination\s*\n.*?([A-Z][a-z]+(?: [A-Z][a-z]+)*)\s*$', t, re.M)
    d['dest'] = (C.code(m.group(1)) or '') if m else ''
    pk = re.findall(r'^\s*(\d+)\s+(.+?)\s+[\d.]+\s*x\s*[\d.]+\s*x\s*[\d.]+\s*CM\s*\n\s*Weight\s+(' + N + r')\s*KG', t, re.M)
    if pk:
        kinds = set()
        for _, kind, _w in pk:
            k = kind.upper()
            kinds.add('PAT' if 'PALLET' in k else 'CTN' if 'CARTON' in k or 'BOX' in k else 'PKG')
        d['packages'] = len(pk)
        d['package_unit'] = kinds.pop() if len(kinds) == 1 else 'PKG'
        d['total_gw'] = round(sum(_num(x[2]) or 0 for x in pk), 3)
        d['package_list'] = ['%s %s %s KG' % (x[0], clean(x[1]), x[2]) for x in pk]
    m = re.search(r'Total Packs:\s*(\d+)', t)
    if m:
        d['packages'] = int(m.group(1))
    agg = collections.OrderedDict()
    rows = t.splitlines()
    IT = re.compile(r'^\s*([A-Z0-9]{2}[A-Z0-9-]{4,})\s+(.+?)\s+(' + N + r')\s*([A-Z]{1,3})\s+(' + N + r')\s*$')
    for i, raw in enumerate(rows):
        m = IT.match(raw)
        if not m or not re.search(r'\d', m.group(1)) or m.group(1).startswith('Package'):
            continue
        a = agg.setdefault(m.group(1), dict(desc=clean(m.group(2)).rstrip(','), qty=0.0, nw=0.0, uom=m.group(4), coo=''))
        a['qty'] += _num(m.group(3)) or 0
        a['nw'] += _num(m.group(5)) or 0
        nxt = rows[i + 1] if i + 1 < len(rows) else ''
        mc = re.search(r'Country of origin:\s*([A-Z]{2})', nxt)
        if mc:
            a['coo'] = mc.group(1)
    for k, a in agg.items():
        d['lines'].append(_line(item=k, desc=a['desc'], qty=a['qty'], uom=a['uom'], nw=round(a['nw'], 3), coo=a['coo']))
    d['total_nw'] = round(sum(a['nw'] for a in agg.values()), 3) or None
    if not d['lines']:
        d['notes'].append('no item rows read on the packing list')
    return d


def sandvik_tax(path, pages, C):
    t = '\n'.join(p['text'] for p in pages)
    if svt_is_invoice(t):
        return [svt_invoice(path, pages, C)]
    if svt_is_carrier_bol(t):
        return [svt_carrier_bol(path, pages, C)]
    return [svt_packing_list(path, pages, C)]


# ======================================================================== Vulcan Industrial Engineering (India) -> Sandvik SG (IMPORT)
def _vulcan_lines(page):
    """Words of one PDF page grouped into lines: [(top, [(x0, text), ...])] (the text layer has no space characters)."""
    import pdfplumber  # noqa: F401  (page is a pdfplumber page)
    rows = {}
    for w in page.extract_words(x_tolerance=1, keep_blank_chars=False):
        rows.setdefault(round(w['top']), []).append((w['x0'], w['text']))
    return [(t, sorted(v)) for t, v in sorted(rows.items())]


def _spc(s):
    """SPROCKET,19T,2"PITCH -> SPROCKET, 19T, 2"PITCH (the PDF has no space characters)."""
    return clean(re.sub(r',(?=\S)', ', ', s or '')).strip(' ,')


def vulcan_invoice(path, pages, C):
    """VULCAN INDUSTRIAL ENGINEERING CO. PVT. LTD. (Gujarat, India) EXPORT INVOICE to Sandvik Mining & Construction
    Logistics (C/O SMC S.E.A., Singapore) - an IMPORT permit. One invoice per page (a PDF can hold several):
    PO, gross / net weight, boxes, line item, description (desc column), HSN code, qty, UOM, unit rate, amount, part no."""
    import pdfplumber
    docs = []
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            t = pg.extract_text(x_tolerance=1) or ''
            if not re.search(r'EXPORT\s*INVOICE', t[:200]):
                continue
            d = _inv(path, 'Vulcan invoice', 'VULCAN INDUSTRIAL ENGINEERING CO. PVT. LTD.')
            d.update(title='EXPORT INVOICE', ship_from_country='IN', ship_to_country='SG', direction='IMPORT')
            m = re.search(r'Invoice\s*No\s*:\s*(\d+)', t)
            d['invoice_no'] = m.group(1) if m else ''
            m = re.search(r'Invoice\s*Date\s*:\s*(\d{2}\.\d{2}\.\d{4})', t)
            d['invoice_date'] = _date_any(m.group(1)) if m else ''
            m = re.search(r'Delivery\s*No\s*:\s*(\d+)', t)
            d['delivery'] = m.group(1) if m else ''
            m = re.search(r'PO\s*No\s*(\d+)', t)
            d['order_no'] = d['customer_po'] = m.group(1) if m else ''
            m = re.search(r'Consignee\s*(.+?),?\s*\n', t)
            d['ship_to'] = d['invoice_to'] = 'SANDVIK MINING & CONSTRUCTION LOGISTICS' if 'Sandvik' in t else ''
            m = re.search(r'Terms\s*of\s*Delivery\s*:\s*(Ex\s*Works|FCA|FOB|CFR|CIF|CPT|CIP|DAP|DDP)', t, re.I)
            d['incoterm'] = 'EXW' if m and m.group(1).replace(' ', '').upper() == 'EXWORKS' else (m.group(1).upper() if m else '')
            m = re.search(r'Gross\s*Wt\.?\s*(' + N + r')\s+(' + N + r')\s*No\s*of\s*Box\s*(\d+)', t)
            if m:
                d['total_gw'], d['total_nw'] = _num(m.group(1)), _num(m.group(2))
                d['boxes'], d['package_unit'] = int(m.group(3)), 'BOX'
            m = re.search(r'Total\s*Amount\s+(' + N + r')', t)
            d['total'] = _num(m.group(1)) if m else None
            m = re.search(r'TOTAL\s*Amount.*\n.*?\b([A-Z]{3})\s+([A-Z]{3})\s*$', t, re.M)
            d['currency'] = m.group(2) if m else 'USD'
            m = re.search(r'Origin\s*of\s*Goods', t)
            coo = 'IN' if m and re.search(r'\bIndia\b', t[m.end():m.end() + 200]) else ''
            d['port_loading'] = 'NHAVA SHEVA' if re.search(r'Nhava\s*sheva', t, re.I) else ''
            # items from word positions: line no x<45, description x 60-205, UOM x 340-385 on the next line
            lines = _vulcan_lines(pg)
            cur = None
            for top, ws in lines:
                txt = ' '.join(w for _, w in ws)
                m = re.match(r'^(\d+/\d+)\s+(.+?)\s+HSN/SAC\s*Code\s*:?\s*:?(\d{6,8})\s+(' + N + r')\s+(' + N + r')\s+(' + N + r')$', txt)
                if m:
                    cur = _line(line=m.group(1), desc=m.group(2), hs=_hs8(m.group(3)), qty=_num(m.group(4)), price=_num(m.group(5)),
                                amount=_num(m.group(6)), coo=coo, currency=d['currency'], order=d['order_no'],
                                delivery=d['delivery'])
                    cur['_more'] = True
                    d['lines'].append(cur)
                    continue
                if not cur:
                    continue
                if re.match(r'Declaration|Total', txt):
                    cur = None
                    continue
                mp = re.search(r'Part\s*No\.?\s*:\s*([A-Z0-9-]+)', txt)
                if mp:
                    cur['item'] = mp.group(1)
                    cur['_more'] = False
                for x0, w in ws:
                    if 340 <= x0 <= 390 and re.fullmatch(r'[A-Z]{2,3}', w):
                        cur['uom'] = w
                if cur.get('_more') and not mp:
                    more = ' '.join(w for x0, w in ws if 60 <= x0 < 205)
                    if more:
                        cur['desc'] += ' ' + more
            for l in d['lines']:
                l.pop('_more', None)
                l['desc'] = _spc(l['desc'])
                l['item'] = l['item'] or l['line']
                l['uom'] = l['uom'] or 'EA'
            if d['lines'] and len(d['lines']) == 1 and d.get('total_nw'):
                d['lines'][0]['nw'], d['lines'][0]['gw'] = d['total_nw'], d['total_gw']
            d['ocr'] = False
            docs.append(_finish(d, C))
    return docs


def vulcan_pl(path, pages, C):
    """Vulcan PACKING LIST (one per delivery): delivery no, invoice no, packs, item with qty, net / gross kg."""
    import pdfplumber
    import readers as R
    out = []
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            t = pg.extract_text(x_tolerance=1) or ''
            if not re.search(r'PACKING\s*LIST', t[:200]):
                continue
            d = R.new_pl(path, 'Vulcan packing list')
            m = re.search(r'Delivery\s*No\s*:\s*(\d+)', t)
            d['delivery'] = m.group(1) if m else ''
            m = re.search(r'Invoice\s*No\s*:\s*(\d+)', t)
            if m:
                d['invoice_refs'] = [m.group(1)]
            m = re.search(r'Total\s*Net\s*&?\s*Gross\s*Weight\s+(' + N + r')\s+(' + N + r')', t)
            if m:
                d['total_nw'], d['total_gw'] = _num(m.group(1)), _num(m.group(2))
            m = re.search(r'in\s*Inch\s*\n\s*(\d+)\s*\n', t)
            d['packages'], d['package_unit'] = (int(m.group(1)), 'BOX') if m else (None, '')
            for mm in re.finditer(r'^(\d+/\d+)\s+(\d{3}\s?\d{7}|\d{6,})\s+(\d+)\s+(.+?)\s+(' + N + r')\s+(' + N + r')\s+(' + N + r')\s+\d+\s+\d+\s+\d+\s*$', t, re.M):
                d['lines'].append(_line(line=mm.group(1), desc=_spc(mm.group(4)), qty=_num(mm.group(5)), nw=_num(mm.group(6)),
                                        gw=_num(mm.group(7)), uom='EA'))
            mp = re.search(r'Part\s*No\.?\s*:\s*([A-Z0-9-]+)', t)
            if mp and len(d['lines']) == 1:
                d['lines'][0]['item'] = mp.group(1)
            out.append(d)
    return out


# ======================================================================== DOW Chemical Pacific (Singapore) - EXPORT
def dow_invoice(path, pages, C):
    """DOW CHEMICAL PACIFIC (SINGAPORE) PRIVATE LIMITED 'Export Invoice': invoice no, date, customer order, Our Reference,
    ship-to, delivery terms, currency (contract currency, not the SGD GST equivalent); per item: qty + packaging, material,
    price per price unit, amount, description, gross / net weight, delivery note, COO, export tariff no., shipment no."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'DOW export invoice', 'DOW CHEMICAL PACIFIC (SINGAPORE) PRIVATE LIMITED')
    d.update(title='EXPORT INVOICE', ship_from_country='SG', direction='EXPORT')
    m = re.search(r'Export Invoice\s+(\d{6,})', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'\bDate\s+(\d{1,2})([A-Z]{3})(\d{4})', t)
    if m and m.group(2) in MONTHS:
        d['invoice_date'] = _dmy(m.group(1), MONTHS[m.group(2)], m.group(3))
    m = re.search(r'Co\.?\s*Reg\.?\s*No\.?\s*(\d{8,9}[A-Z])', t)
    d['uen'] = m.group(1) if m else ''
    m = re.search(r'Customer Order Number\s+Our Reference\s*\n\s*(\S+)\s+(\S+)', t)
    if m:
        d['customer_po'], d['our_ref'] = m.group(1), m.group(2)
    m = re.search(r'Delivery Terms:.*\n(?:.*\n){0,2}?\s*(EXW|FCA|FOB|CFR|CIF|CPT|CIP|DAP|DPU|DDP)\s+(.+)', t)
    if m:
        d['incoterm'], d['place'] = m.group(1), clean(m.group(2))
    m = re.search(r'Due date\s+Currency\s*\n.*?\b(US Dollar|Euro|Singapore Dollar|Japanese Yen|[A-Z]{3})\s*$', t, re.M)
    cur = {'US DOLLAR': 'USD', 'EURO': 'EUR', 'SINGAPORE DOLLAR': 'SGD', 'JAPANESE YEN': 'JPY'}.get(m.group(1).upper(), m.group(1)) if m else ''
    d['currency'] = cur or 'USD'
    m = re.search(r'Ship to\s*\n(?:.*\n)?(.*(?:CORP|LTD|LIMITED|INC|CO)\.?)\s*$', t, re.M)
    if m:
        d['ship_to'] = clean(m.group(1)).upper()
    m = re.search(r'Freight/Border-Destination:\s*(.+)', t)
    d['ship_to_country'] = (C.code(clean(m.group(1))) or '') if m else ''
    m = re.search(r'Mode of Transport.*\n(?:.*\n){0,3}?\s*(Sea|Air|Road|Truck)\b', t)
    d['transport'] = m.group(1).title() if m else ''
    m = re.search(r'Port of Discharge:\s*(.+)', t)
    d['port_dest'] = clean(re.sub(r'\s*Port$', '', m.group(1).strip())).upper() if m else ''
    rows = t.splitlines()
    cur_l = None
    for i, raw in enumerate(rows):
        r = raw.strip()
        m = re.match(r'^(\d{1,4}0)\s+(' + N + r')\s+([A-Za-z]+)\s+(' + N + r')$', r)          # 10 12.000 Bag 0.240
        if m and i + 1 < len(rows):
            m2 = re.match(r'^(\d{6,})\s+([A-Z]{1,4})\s+(' + N + r')\s+(' + N + r')$', rows[i + 1].strip())   # material, MT, price, amount
            if m2:
                cur_l = _line(line=m.group(1), item=m2.group(1).lstrip('0') or m2.group(1), qty=_num(m.group(2)),
                              uom=m.group(3).upper(), amount=_num(m2.group(4)), currency=d['currency'])
                cur_l['price'] = round(cur_l['amount'] / cur_l['qty'], 4) if cur_l['qty'] else None
                cur_l['price_note'] = '%s %s per %s' % (m2.group(3), d['currency'], m2.group(2))
                cur_l['desc_lines'] = []
                d['lines'].append(cur_l)
                continue
        if not cur_l:
            continue
        if re.match(r'^(Subtotal|GST %|Total\b)', r):
            cur_l = None
            continue
        if re.match(r'^[A-Z]{3}/\d*[A-Z]{1,4}$', r) or re.match(r'^\d{6,}\s+[A-Z]{1,4}\s', r):
            continue
        mg = re.match(r'Gross Weight\s*:\s*(' + N + r')\s*KG', r, re.I)
        if mg:
            cur_l['gw'] = _num(mg.group(1))
            cur_l['_done'] = True
            continue
        mn = re.match(r'Net Weight\s*:\s*(' + N + r')\s*KG', r, re.I)
        if mn:
            cur_l['nw'] = _num(mn.group(1))
            continue
        mc = re.match(r'Country Of Origin:\s*(.+)', r, re.I)
        if mc:
            cur_l['coo'] = C.code(clean(mc.group(1))) or ''
            continue
        mh = re.match(r'Customs tariff nbr.*?:\s*(\d{6,10})', r, re.I)
        if mh:
            cur_l['hs'] = _hs8(mh.group(1))
            continue
        md = re.match(r'Delivery note:\s*(\d+)', r, re.I)
        if md:
            cur_l['delivery'] = cur_l['order'] = md.group(1)
            continue
        ms = re.match(r'Shipment No\.?:\s*(\S+)', r, re.I)
        if ms:
            d['shipment_no'] = ms.group(1)
            continue
        mb = re.match(r'No\. of (\w+)\s*:\s*(\d+)', r, re.I)
        if mb:
            continue
        if not cur_l.get('_done') and not re.match(r'Customer Material Number|K-\d|VEHICLE', r, re.I):
            cur_l['desc_lines'].append(clean(r.replace('™', '').replace('(tm)', '')))
    for l in d['lines']:
        l['desc'] = ' '.join(l.pop('desc_lines', [])).strip()
        l.pop('_done', None)
    m = re.search(r'^\s*Total\s+(' + N + r')\s*$', t, re.M)
    d['total'] = _num(m.group(1)) if m else None
    d['total_nw'] = round(sum(l.get('nw') or 0 for l in d['lines']), 3) or None
    d['total_gw'] = round(sum(l.get('gw') or 0 for l in d['lines']), 3) or None
    m = re.search(r'Shipping Marks and Additional Information\s*\n(.+)', t)
    d['marks'] = clean(m.group(1)) if m else ''
    return [_finish(d, C)]


def dow_pl(path, pages, C):
    """DOW 'Consolidated Packing List' (shipment number, delivery / item, qty + packaging, gross / net, COO).
    Some copies are printed without space characters (text runs together) - the patterns allow that."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'DOW packing list')
    m = re.search(r'Shipment\s*Number\s*\n.*?(\d{8,})\s*$', t, re.M)
    d['shipment_no'] = m.group(1) if m else ''
    m = re.search(r'Document\s*Date\s.*\n.*?(\d{2})/(\d{2})/(\d{4})', t)
    d['doc_date'] = '%s/%s/%s' % (m.group(2), m.group(1), m.group(3)) if m else ''
    IT = re.compile(r'^\s*(?:\S+\s+)?(\d{6,})\s*/\s*(\d+)\s+(' + N + r')\s*([A-Z]+)\s*/\s*(\S+)\s+(\d{6,})\s*(.+)$', re.M)
    for mm in IT.finditer(t):
        seg = t[mm.end():mm.end() + 700]
        g = re.search(r'\s(?:Gross\s*Weight)\s*(' + N + r')\s*KG', seg)
        n = re.search(r'\s(?:Net\s*Weight)\s*(' + N + r')\s*KG', seg)
        c = re.search(r'Country\s*of\s*Origin\s*:\s*([A-Za-z ]+)', seg)
        desc = clean(mm.group(7).replace('™', ''))
        d['lines'].append(_line(line=mm.group(2), item=mm.group(6).lstrip('0'), desc=desc,
                                qty=_num(mm.group(3)), uom=mm.group(4), lots=[mm.group(5)], delivery=mm.group(1),
                                gw=_num(g.group(1)) if g else None, nw=_num(n.group(1)) if n else None,
                                coo=(C.code(clean(c.group(1))) or '') if c else ''))
        d['delivery'] = mm.group(1)
    m = re.search(r'Total\s*Gross\s*weight\s*:\s*(' + N + r')\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total\s*Net\s*weight\s*:\s*(' + N + r')\s*KG', t)
    d['total_nw'] = _num(m.group(1)) if m else None
    m = re.search(r'Piece\s*Count\s*:\s*(' + N + r')\s*([A-Z]+)', t)
    if m:
        d['pieces'], d['pieces_unit'] = _num(m.group(1)), m.group(2)
    return [d]


def dow_booking(path, pages, C):
    """DOW 'Consolidated Booking Advice' (shipping instruction): shipment date (= departure), ETA, terms, transport mode,
    ports. Lowest priority - the forwarder's B/L wins."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'DOW booking advice', 1)
    cg = {'mode': 'Sea'} if re.search(r'LCL|FCL|Container', t, re.I) else {}
    m = re.search(r'Shipment\s*Date\s*\(MM/DD/YYYY\).*\n\s*(\d{2})/(\d{2})/(\d{4})', t)
    if m:
        cg['etd'] = '%s/%s/%s' % (m.group(2), m.group(1), m.group(3))
    m = re.search(r'ETA\s*([A-Z]+)\s*:\s*([A-Z]{3})\.?\s*(\d{1,2}),?\s*(\d{4})', t)
    if m and m.group(2) in MONTHS:
        cg['eta'] = _dmy(m.group(3), MONTHS[m.group(2)], m.group(4))
        cg['discharge_name'] = m.group(1)
    if re.search(r'LCL', t):
        cg['cargo_lcl'] = True
    d['notes'].append('DOW booking advice: departure (shipment date) %s, ETA %s - the B/L wins'
                      % (cg.get('etd', '?'), cg.get('eta', '?')))
    d['cargo'] = cg
    return [d]


# ======================================================================== Panasonic (Japan) Invoice & Packing List
PANA_UNITS = {'PIECE': 'PCS', 'UNIT': 'UNIT', 'PACK': 'PACK', 'PAIR': 'PR', 'CAN': 'CAN', 'SET': 'SET',
              'METER': 'MTR', 'ROLL': 'ROL', 'BOX': 'BOX', 'BOTTLE': 'BTL', 'SHEET': 'SHT'}
PANA_HEAD = re.compile(r'^\(?\s*(\d{5})(?:\s*-\s*(\d{5}))?\s*\)?\s+(.+)$')     # case no (or range) + rest
PANA_ONE = {'i', 'I', 'l', 'j', '|', '!', 'if', '1i', 'li', 'Il', 'il', 'iI', ']', '['}
PANA_OCR_NUM = str.maketrans('iIl|!jtOoSs$', '111111100555')


def _pana_numlike(t):
    return t in PANA_ONE or bool(re.fullmatch(r'[\d,.]*\d[\d,.]*[:;.]?', t)) or bool(re.fullmatch(r'[\diIl|!jtOo$,.]+', t) and re.search(r'\d', t))


def _pana_val(t):
    if t in PANA_ONE:
        return 1.0
    v = t.translate(PANA_OCR_NUM).strip(':;.')
    return _num(v) if re.fullmatch(r'\d[\d,]*(?:\.\d+)?', v) else None


def _pana_item(ln):
    """Item line of the Panasonic invoice -> dict(case, part, qty, price, amount, nw, gw, fixed) or None.
    '60305 * N510064042AA 1 1 17,500 17,500 4.200 4.700 0.031000' / '(60309) * N510056967AA 4 4 6,430 25,720'.
    OCR misreads of 1 ('i', 'I', 'j', 'if') and of single digits are repaired with qty x price = amount."""
    m = PANA_HEAD.match(ln)
    if not m:
        return None
    toks = m.group(3).split()
    while toks and len(toks[0]) <= 2 and not toks[0].isdigit():       # mark: * f / -f -/ + A
        toks.pop(0)
    if not toks or not re.search(r'[A-Z0-9]{4}', toks[0]):
        return None
    part = [toks.pop(0)]
    while toks and not _pana_numlike(toks[0]):
        part.append(toks.pop(0))
    toks = [t for t in toks if re.search(r'[0-9A-Za-z|!$\]\[]', t)]
    if len(toks) < 3:
        return None
    run = []                                       # trailing weights: NW GW M3 (3 decimals / 6 decimals)
    while toks and len(run) < 3:
        t = toks[-1]
        if re.fullmatch(r'\d+\.\d{3,}[:;.]?|\d[.,]\d{6}[:;.]?', t) or (run and re.fullmatch(r'\d{1,3},\d{3}', t)):
            run.insert(0, toks.pop())
        else:
            break
    ints = [_pana_val(t) for t in toks][-4:]
    if len(ints) < 3:
        return None
    if len(ints) == 3:
        ints = [None] + ints
    ppk, qty, price, amt = ints
    fixed = False
    cands = [q for q in (qty, ppk) if q]
    q = next((q for q in cands if price and amt and abs(q * price - amt) < 0.5), None)
    if q is None:
        if price and amt and abs(amt / price - round(amt / price)) < 1e-6 and (not cands or (qty != ppk)):
            q = round(amt / price)
        elif cands and price:
            q = cands[0]
            amt = round(q * price, 2)
        elif cands and amt:
            q = cands[0]
            price = round(amt / q, 4)
        else:
            q = cands[0] if cands else None
        fixed = True
    nw = gw = None
    if len(run) >= 2:
        nw, gw = _pana_kg(run[0]), _pana_kg(run[1])
    pn = ''.join(part).strip('.,:;')
    for _ in range(4):                             # OCR: O / i / l next to digits are 0 / 1 (N6i0014970AF, MTKBO00020AA)
        pn = re.sub(r'(?<=\d)[O]|[O](?=\d)', '0', re.sub(r'(?<=\d)[il|](?=\d)', '1', pn))
    return dict(case=m.group(1), rng=bool(m.group(2)), last=m.group(2) or m.group(1), part=pn,
                qty=q, price=price, amount=amt, nw=nw, gw=gw, fixed=fixed)


PANA_STOP = re.compile(r'^(?:SPARE PARTS|COMPONENT\s+(?:ASSY|PLACEMENT)|GA\d{4}|4500\d{6}|FE\d|F[EB]S?\d{4}|=\s*TO BE CONTINUED|REMARKS|'
                       r'(?:PIECES|UNITS|PACKS|PAIRS|CANS|SETS)\b|\d{4}\b|\(?\d{5}\)?\s)|CARTON\s*BOX|CARTONBOX|INVOICE\s*&', re.I)


def _pana_kg(s):
    """Weight column: always 3 decimals, so OCR '4,200' is 4.200 kg."""
    if not s:
        return None
    s = s.strip('.,')
    if re.fullmatch(r'\d{1,3},\d{3}', s):
        s = s.replace(',', '.')
    return _num(s)


def panasonic_inv(path, pages, C):
    """PANASONIC OPERATIONAL EXCELLENCE (Japan) 'INVOICE & PACKING LIST' - usually a sideways scan (OCR).
    Item: case no, part no, P/Pkg, total qty, unit price, amount; net / gross weight per case (first item of the case);
    unit from 'PIECE PER PIECE' / 'UNIT PER UNIT' headers; 'MADE IN ...' under the item, else Japan."""
    body = [p for p in pages if re.search(r'INVOICE\s*&\s*PACKING', p['text'][:300], re.I)] or pages
    t = '\n'.join(p['text'] for p in body)
    p1 = body[0]['text']
    d = _inv(path, 'Panasonic invoice & packing list', 'PANASONIC OPERATIONAL EXCELLENCE CO., LTD.')
    d.update(title='INVOICE & PACKING LIST', ship_from_country='JP', ship_to_country='SG', direction='IMPORT',
             ocr=any(p.get('ocr') for p in pages))
    m = re.search(r'PANASONIC\s+OPERATIONAL\s+EXCELLENCE\s+CO\.?,?\s*LTD\.?', t)
    if m:
        d['supplier'] = d['invoice_from'] = d['ship_from'] = 'PANASONIC OPERATIONAL EXCELLENCE CO., LTD.'
    # invoice no: page 1 / every page head / case marks - most common reading
    cands = re.findall(r'\b([A-Z0-9]{1,4}P\d{6}[A-Z0-9]{2})\b', '\n'.join(p['text'][:400] for p in pages) + '\n' +
                       '\n'.join(re.findall(r'CASE MARK.*', '\n'.join(p['text'] for p in pages))))
    first = re.search(r'Invoice No\.?[\s\S]{0,120}?\b([A-Z0-9]{1,4}P\d{6}[A-Z0-9]{2})\b', p1)
    if first:
        d['invoice_no'] = first.group(1)
    elif cands:
        d['invoice_no'] = collections.Counter(cands).most_common(1)[0][0]
    if d['invoice_no'] and d['invoice_no'][0] in '1lD|' and d['invoice_no'][1:3] == 'JP':
        d['invoice_no'] = 'I' + d['invoice_no'][1:]
    m = re.search(r'Date\s*[:;]\s*([A-Z]{3})\.?\s*(\d{1,2})\s*[\'`’]?\s*(\d{2})\b', p1)
    if m:
        d['invoice_date'] = _dmy(m.group(2), MONTHS.get(m.group(1).upper(), 1), m.group(3))
    m = re.search(r'Order Placed\s+(.+?)\s+Invoice No', p1)
    if m:
        d['invoice_to'] = d['ship_to'] = clean(m.group(1)).upper()
    m = re.search(r'^\s*\d{8}\s+(\d+.*?SINGAPORE\s*\d{6})', p1, re.M)
    if m:
        d['ship_to_block'] = d['invoice_to_block'] = [d['ship_to'], clean(m.group(1))]
    m = re.search(r'Sailing on[\s\S]{0,120}?(\d{2})/(\d{2})/(\d{2})\b', p1)
    if m:
        d['ship_date'] = _dmy(m.group(2), m.group(1), m.group(3))
    m = re.search(r'Payment\s*[:;]\s*(.+)', p1)
    d['payment_terms'] = clean(m.group(1)) if m else ''
    m = re.search(r'\b(FOB|CIF|CFR|EXW|FCA|DAP|DDP|CPT|CIP)\s+JAPAN\b', t)
    d['incoterm'] = m.group(1) if m else 'FOB'
    m = re.search(r'\b(JPY|USD|EUR|SGD)\b', t)
    d['currency'] = m.group(1) if m else 'JPY'
    m = re.search(r'Order No\.?[\s\S]{0,80}?\b(4500\d{6})\b', p1)
    d['order_no'] = m.group(1) if m else ''

    # maker printed under an item, left of the description (part no column, x ~140 pt; description ~168 pt) = brand
    makers = set()
    for p in body:
        for w in p.get('words') or []:
            tx = re.sub(r'[\s.,:;]+$', '', w['text'].strip())
            if 122 <= w['x0'] <= 158 and re.fullmatch(r'[A-Z][A-Z&\- ]{2,30}', tx) and not re.search(
                    r'CARTON|SPARE|COMPONENT|PIECE|UNIT|PACK|PAIR|REMARK|INVOICE|MADE IN|PAYMENT|SHIPPED|CONSIGN|MESSRS', tx):
                makers.add(tx)
    unit, lines, cur, case_w, made, rng = 'PIECE', [], None, {}, 0, None
    gtot = None
    for raw in t.splitlines():
        ln = raw.strip().lstrip(';:,.\'"“” ').strip()
        if not ln:
            continue
        mu = re.search(r'\b(PIECE|UNIT|PACK|PAIR|CAN|SET|METER|ROLL|BOX|BOTTLE|SHEET)\s+PER\b', ln)
        if mu:
            unit = mu.group(1)
        mt = re.search(r'(\d+)\.?\s*CARTON\s*BOX\b.*?[¥Y]\s*([\d,]{4,})\s+([\d.,]+)\s+([\d.,]+)', ln)
        if mt:
            amt = _num(mt.group(2))
            if gtot is None or (amt or 0) > gtot[1]:
                gtot = (int(mt.group(1)), amt, _pana_kg(mt.group(3)), _num(mt.group(4)))
        it = _pana_item(ln)
        if it:
            cur = dict(case=it['case'], part=it['part'], qty=it['qty'], unit=unit, price=it['price'], amount=it['amount'],
                       desc=[], coo='', fixed=it['fixed'], brand='',
                       boxes=[str(n) for n in range(int(it['case']), int(it['last']) + 1)] if int(it['last']) - int(it['case']) < 50 else [it['case']])
            if it['nw'] is not None:
                case_w[it['case']] = (it['nw'], it['gw'])
            rng = it['case'] if it['rng'] else None
            lines.append(cur)
            continue
        mr = rng and re.match(r'^(\d+)\s+(\d+[.,]\d{3})\s+(\S+)\s+\d[.,]\d+', ln)
        rng = None
        if mr:                                    # case range 60298-60300: next line = cases / NW / GW / M3 of the range
            case_w[cur['case']] = (_pana_kg(mr.group(2)), _pana_val(mr.group(3)))
            continue
        if not cur:
            continue
        mm = re.match(r'^MADE\s+IN\s+([A-Z ]+?)\s*$', ln)
        if mm:
            cur['coo'] = C.code(mm.group(1)) or mm.group(1)
            made += 1
            continue
        if re.fullmatch(r'[\dIil|!\]\[j.,;:\'" ]{1,5}', ln):              # number of cases under the case no
            continue
        mk = re.sub(r'[\s.,:;]+$', '', ln)
        if mk in makers and cur['desc']:
            cur['brand'] = mk
            continue
        if PANA_STOP.search(ln):              # next group / case / carton line / page end: description done
            cur = None
            continue
        cur['desc'].append(re.sub(r'\s+\d{4}$', '', ln))
    # net / gross weight: per case -> shared over the items of the case by value
    by_case = collections.OrderedDict()
    for l in lines:
        by_case.setdefault(l['case'], []).append(l)
    shared = 0
    for case, ls in by_case.items():
        nw, gw = case_w.get(case, (None, None))
        tot = sum(l['amount'] or 0 for l in ls) or 1
        acc_n = acc_g = 0.0
        for i, l in enumerate(ls):
            if nw is None:
                l['nw'] = l['gw'] = None
                continue
            if i == len(ls) - 1:
                l['nw'], l['gw'] = round(nw - acc_n, 3), round((gw or 0) - acc_g, 3) if gw else None
            else:
                l['nw'] = round(nw * (l['amount'] or 0) / tot, 3)
                l['gw'] = round(gw * (l['amount'] or 0) / tot, 3) if gw else None
                acc_n += l['nw']
                acc_g += l['gw'] or 0
        if len(ls) > 1:
            shared += 1
    nocoo = 0
    for i, l in enumerate(lines, 1):
        desc = ' '.join(re.sub(r'^[^A-Z0-9(]*[a-z\]\[|!]?\s+(?=[A-Z(])', '', x) for x in l['desc'])
        desc = re.sub(r'\s\d{4}\s*[.:;,]?(?=\s|$)', ' ', ' ' + desc + ' ')       # line no printed beside the text
        desc = re.sub(r'(?<=\()[iIl1](?=PACK)|(?<=\s)[iIlC]i?(?=PACK\s*=)', '1', desc)    # (IPACK = -> (1PACK =
        desc = clean(re.sub(r'\s[.:;,]+(?=\s|$)', ' ', desc)).strip(' ,;.:')
        if not l['coo']:
            l['coo'], nocoo = 'JP', nocoo + 1
        d['lines'].append(_line(line=str(i), item=l['part'], desc=desc or l['part'], qty=l['qty'], uom=PANA_UNITS.get(l['unit'], l['unit']),
                                price=l['price'], amount=l['amount'], coo=l['coo'], nw=l['nw'], gw=l['gw'],
                                currency=d['currency'], boxes=l['boxes'], pl=l['case'], brand=l['brand']))
    if gtot:
        d['boxes'], d['package_unit'], d['total'], d['total_nw'], d['total_gw'] = gtot[0], 'CTN', gtot[1], gtot[2], gtot[3]
    else:
        d['notes'].append('grand total line (cartons / amount / weights) not readable on the scan - item lines used')
    sn = round(sum(l['nw'] or 0 for l in d['lines']), 3)
    if d.get('total_nw') and abs(sn - d['total_nw']) > 0.01:
        d['notes'].append('net weight of the cases adds up to %s KG but the total is %s KG - check the case weights' % (fmt_kg(sn), fmt_kg(d['total_nw'])))
    sq = collections.Counter()
    for l in lines:
        sq[l['unit']] += l['qty'] or 0
    tq = {k[:-1]: _num(v) for k, v in re.findall(r'\b(PIECES|UNITS|PACKS|PAIRS|CANS|SETS)\s+([\d,]+)', t[t.rfind('REMARKS') - 400:t.rfind('REMARKS')] if 'REMARKS' in t else '')}
    diff = ['%s %s read / %s on the invoice' % (k, '{:,.0f}'.format(sq.get(k, 0)), '{:,.0f}'.format(v)) for k, v in tq.items() if v is not None and abs(sq.get(k, 0) - v) > 0.001]
    if diff:
        d['notes'].append('quantity totals differ: ' + '; '.join(diff))
    fx = [l['part'] for l in lines if l.get('fixed')]
    if fx:
        d['notes'].append('qty / unit price / amount not readable together on the scan for %d item(s) - repaired with '
                          'qty x unit price = amount, check: %s' % (len(fx), ', '.join(fx[:20])))
    if shared:
        d['notes'].append('%d case(s) hold several items - the case net / gross weight is shared over its items by value' % shared)
    if nocoo:
        d['notes'].append('%d item(s) without "MADE IN" - country of origin Japan used' % nocoo)
    d['notes_ocr'] = 'scanned Panasonic invoice read by OCR - check part no / qty / amounts' if d['ocr'] else ''
    mh = re.search(r'(?:KWE|HAWB)[ _-]*(\d{8,12})', os.path.basename(path), re.I)
    if mh and not d['hawb']:
        d['hawb'], d['hawb_from_name'] = mh.group(1), True
    d = _finish(d, C)
    out = [d]
    m = re.search(r'Shipped per\s+(AIR|SEA|OCEAN|VESSEL)[\s\S]{0,80}?\bfrom\s+([A-Z]+)', p1)
    if m:                                    # transport from the invoice; an air waybill / B/L added to the job wins
        a = _awb(path, 'Panasonic invoice (shipped per)', 0)
        a['hawb'] = d['hawb']
        a['cargo']['mode'] = 'Air' if m.group(1) == 'AIR' else 'Sea'
        iata = E.CITY_IATA.get(m.group(2)) if a['cargo']['mode'] == 'Air' else None
        if iata and iata in E.AIRPORTS:
            a['cargo']['loading_code'], a['cargo']['loading_name'] = E.AIRPORTS[iata]
            a['notes'].append('loading port %s from "Shipped per %s FREIGHT from %s" on the invoice - check against the air waybill'
                              % (E.AIRPORTS[iata][0], m.group(1), m.group(2)))
        out.append(a)
    return out


# ======================================================================== Taiyo Yuden (Singapore) tax invoice
def taiyo_yuden(path, pages, C):
    """TAIYO YUDEN (SINGAPORE) PTE. LTD. TAX INVOICE (export, e.g. to Kaifa Technology Malaysia by road).
    Group line "DESCRIPTION (HS No:85322400)" = description + HS of the item lines under it.
    Item line "PO-line  CUSTOMER-PART  QTY(PCS)  UNIT PRICE (per 1000 pcs)  AMOUNT", next line
    "<series> <maker part> Made In <country>" = maker part + COO. "Carton : 1-11" = 11 cartons. No weights."""
    d = _inv(path, 'Taiyo Yuden tax invoice', 'TAIYO YUDEN (SINGAPORE) PTE. LTD.')
    d['title'] = 'TAX INVOICE'
    t = '\n'.join(p['text'] for p in pages)
    m = re.search(r'Invoice No\s*:\s*(\d{6,12})', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice Date\s*:\s*(\d{1,2}/\d{1,2}/\d{4})', t)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'Ship Date\s*:\s*(\d{1,2}/\d{1,2}/\d{4})', t)
    d['ship_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'Sold To\s*:\s*(.+?)\s+Co\. Reg', t)
    sold = clean(m.group(1)).upper() if m else ''
    m = re.search(r'Ship To\s*:\s*(.+?)\s+HAWB\s*:', t)
    d['ship_to'] = clean(m.group(1)).upper() if m else sold
    d['invoice_to'] = sold or d['ship_to']
    m = re.search(r'Bill To\s*:\s*(.+)', t)
    if m:
        d['bill_to'] = clean(m.group(1)).upper()
    m = re.search(r'Ship To\s*:.*\n((?:.*\n){1,4}?)\s*Cust Code', t)
    blk = m.group(1) if m else ''
    d['ship_to_block'] = [clean(re.sub(r'\s+(?:Ship Mark|Remarks)\s*:.*$', '', x)) for x in blk.splitlines() if x.strip()]
    d['ship_to_country'] = C.find_in(' '.join(d['ship_to_block'])) or ('MY' if 'MALAYSIA' in blk.upper() else '')
    m = re.search(r'Term Point\s*:\s*(EXW|FCA|FAS|FOB|CFR|CNF|CIF|CPT|CIP|DAP|DPU|DDP)\b\s*(.*)', t)
    if m:
        d['incoterm'], d['place'] = m.group(1), clean(m.group(2))
    m = re.search(r'Payment Term\s*:\s*(.+)', t)
    d['payment_terms'] = clean(m.group(1)) if m else ''
    m = re.search(r'Ship Via\s*:\s*(.+)', t)
    d['carrier'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'\bMAWB\s*:\s*(\S+)', t)
    if m and not m.group(1).endswith(':'):
        d['mawb'] = m.group(1)
    m = re.search(r'\bHAWB\s*:\s*([A-Z0-9-]{5,})', t)
    d['hawb'] = m.group(1) if m else ''
    m = re.search(r'Carton\s*:\s*(\d+)\s*-\s*(\d+)', t)
    if m:
        d['boxes'], d['package_unit'] = int(m.group(2)) - int(m.group(1)) + 1, 'CTN'
    else:
        m = re.search(r'Carton\s*:\s*(\d+)\b', t)
        if m:
            d['boxes'], d['package_unit'] = 1, 'CTN'
    m = re.search(r'TOTAL\s*:\s*(' + N + r')\s*PCS\s+([A-Z]{3})\s+(' + N + r')', t)
    if m:
        d['currency'], d['total'] = m.group(2), _num(m.group(3))
    else:
        m = re.search(r'TOTAL AMOUNT DUE\s*:\s*([A-Z]{3})\s+(' + N + r')', t)
        if m:
            d['currency'], d['total'] = m.group(1), _num(m.group(2))
    m = re.search(r'Exch\. Rate 1 = (' + N + r') SGD', t)
    if m:
        d['exch_rate_sgd'] = _num(m.group(1))
    m = re.search(r'Unit Price\s*\n.*Per (\d+) Pcs', t)
    per = int(m.group(1)) if m else 1000
    GRP = re.compile(r'^(.+?)\s*\(HS\s*No\s*[:.]?\s*(\d{6,10})\)\s*$')
    IT = re.compile(r'^(\d{6,}-\d+)\s+(\S+)\s+(' + N + r')\s+(' + N + r')\s+(' + N + r')$')
    MK = re.compile(r'^(?:([A-Z]{1,4})\s+)?(\S+)\s+Made\s+In\s+(.+)$', re.I)
    gdesc, ghs, cur, unread = '', '', None, []
    for raw in t.splitlines():
        ln = raw.strip()
        if not ln:
            continue
        if re.match(r'^TOTAL\s*:', ln):
            cur = None
            break
        g = GRP.match(ln)
        if g:
            gdesc, ghs, cur = clean(g.group(1)).upper(), _hs8(g.group(2)), None
            continue
        m = IT.match(ln)
        if m:
            qty, uprice, amt = _num(m.group(3)), _num(m.group(4)), _num(m.group(5))
            if qty and uprice and amt is not None and abs(qty * uprice / per - amt) > max(0.02, amt * 0.001):
                d['notes'].append('line %s: %s x %s / %d is not %s - check' % (m.group(1), m.group(3), m.group(4), per, m.group(5)))
            cur = _line(item=m.group(2), desc=gdesc, hs=ghs, qty=qty, uom='PCE', amount=amt, currency=d.get('currency') or '',
                        price=round(amt / qty, 8) if qty and amt is not None else None)
            cur['po'] = m.group(1)
            cur['order'] = m.group(1)
            cur['cust_part'] = m.group(2)
            cur['price_per_%d' % per] = uprice
            d['lines'].append(cur)
            continue
        if re.match(r'^\d{6,}-\d+\s', ln):
            unread.append(ln)
            cur = None
            continue
        if cur is not None and not cur.get('maker_part'):
            k = MK.match(ln)
            if k:
                cur['series'], cur['maker_part'] = k.group(1) or '', k.group(2)
                cur['coo'] = C.code(clean(k.group(3))) or ''
                if cur['coo'] == '' :
                    d['notes'].append('country "%s" (line %s) not known - type the COO' % (k.group(3), cur['po']))
                cur['desc'] = clean('%s %s' % (cur['desc'], k.group(2))).upper()
    if unread:
        d['notes'].append('%d item line(s) not readable - check them on the invoice: %s' % (len(unread), ' | '.join(unread[:5])))
    d['notes'].append('unit price on the invoice is per %d pcs - the permit unit price is amount / quantity' % per)
    d['transport'] = 'Road' if d.get('ship_to_country') == 'MY' else ''
    if d['transport']:
        d['notes'].append('Singapore -> Malaysia: transport mode Road (change it on the Cargo page if not by truck)')
    d['direction'] = 'EXPORT'
    d = _finish(d, C)
    d['notes'] = [n if not n.startswith('no net weight on the invoice') else
                  'no weights on the invoice - add the Taiyo Yuden packing list (net + gross weight per carton)' for n in d['notes']]
    return d


def taiyo_yuden_pl(path, pages, C):
    """TAIYO YUDEN (SINGAPORE) PACKING LIST of one tax invoice: per carton "CTN PO-line DESCRIPTION (HS No:...)", then
    "CUSTOMER-PART QTY NET-KG", then "<series> <maker part> Made In <country>"; "Weight & Dimension" table = net / gross
    weight + M3 per carton. Lines are added up per PO line + part (= the invoice line) -> exact net / gross weight."""
    import readers as R
    d = R.new_pl(path, 'Taiyo Yuden packing list')
    t = '\n'.join(p['text'] for p in pages)
    m = re.search(r'Invoice No\s*:\s*(\d{6,12})', t)
    if m:
        d['invoice_refs'] = [m.group(1)]
    m = re.search(r'\bHAWB\s*:\s*([A-Z0-9-]{5,})', t)
    d['hawb'] = m.group(1) if m else ''
    d['shipper'] = 'TAIYO YUDEN (SINGAPORE) PTE. LTD.'
    m = re.search(r'Ship To\s*:\s*(.+?)\s+HAWB\s*:', t)
    d['consignee'] = clean(m.group(1)).upper() if m else ''
    CT = re.compile(r'^(\d{4})\s+(\d{6,}-\d+)\s+(.+?)\s*\(HS\s*No\s*[:.]?\s*(\d{6,10})\)\s*$')
    QL = re.compile(r'^(\S+)\s+(' + N + r')\s+(' + N + r')$')
    MK = re.compile(r'^(?:([A-Z]{1,4})\s+)?(\S+)\s+Made\s+In\s+(.+)$', re.I)
    WD = re.compile(r'^(\d{4})\s+(' + N + r')\s+(' + N + r')\s+[\d.]+\s*x\s*[\d.]+\s*x\s*[\d.]+\s+(' + N + r')$')
    rows, ctn_gw, cur = [], {}, None
    for raw in t.splitlines():
        ln = raw.strip()
        m = CT.match(ln)
        if m:
            cur = dict(ctn=m.group(1), order=m.group(2), desc=clean(m.group(3)).upper(), hs=_hs8(m.group(4)))
            continue
        m = WD.match(ln)
        if m:
            ctn_gw[m.group(1)] = (_num(m.group(2)), _num(m.group(3)), _num(m.group(4)))
            continue
        if cur and 'qty' not in cur:
            m = QL.match(ln)
            if m:
                cur.update(item=m.group(1), qty=_num(m.group(2)), nw=_num(m.group(3)))
                rows.append(cur)
                continue
        if cur and 'qty' in cur and 'coo' not in cur:
            m = MK.match(ln)
            if m:
                cur['maker_part'], cur['coo'] = m.group(2), C.code(clean(m.group(3))) or ''
    lines = collections.OrderedDict()
    for r in rows:
        k = (r['order'], r['item'])
        l = lines.get(k)
        if not l:
            l = lines[k] = _line(item=r['item'], desc=clean('%s %s' % (r['desc'], r.get('maker_part', ''))).upper(), hs=r['hs'],
                                 qty=0.0, uom='PCE', coo=r.get('coo', ''), order=r['order'], nw=0.0, gw=None)
            l['po'], l['cartons'] = r['order'], []
        l['qty'] += r['qty'] or 0
        l['nw'] = round(l['nw'] + (r['nw'] or 0), 3)
        l['cartons'].append(r['ctn'])
        g = ctn_gw.get(r['ctn'])
        if g and g[1] is not None:
            n_share = sum(1 for x in rows if x['ctn'] == r['ctn'])     # a carton with several lines: gross shared by net
            tot_n = sum(x['nw'] or 0 for x in rows if x['ctn'] == r['ctn'])
            l['gw'] = round((l['gw'] or 0) + (g[1] * ((r['nw'] or 0) / tot_n if tot_n else 1.0 / n_share)), 3)
    d['lines'] = list(lines.values())
    m = re.search(r'TOTAL\s+(\d+)\s+Cartons?\s+(' + N + r')\s*PCS\s+(' + N + r')\s*Kg', t, re.I)
    if m:
        d['packages'], d['package_unit'], d['total_nw'] = int(m.group(1)), 'CTN', _num(m.group(3))
    m = re.search(r'In Total\s*:\s*(' + N + r')\s+(' + N + r')\s+(' + N + r')', t)
    if m:
        d['total_nw'], d['total_gw'], d['cbm'] = _num(m.group(1)), _num(m.group(2)), _num(m.group(3))
    s_nw = round(sum(l['nw'] for l in d['lines']), 3)
    if d.get('total_nw') is not None and abs(s_nw - d['total_nw']) > 0.002:
        d['notes'].append('carton net weights add up to %s KG but the total is %s KG - check' % (s_nw, d['total_nw']))
    if d.get('packages') and len({r['ctn'] for r in rows}) != d['packages']:
        d['notes'].append('%d cartons read, packing list total %d - check' % (len({r['ctn'] for r in rows}), d['packages']))
    if not d['lines']:
        d['notes'].append('no carton lines could be read')
    return d


# ======================================================================== DuPont / Specialty Electronic Materials (BDP pack)
def _eu(s):
    """European number '4.267,66' -> 4267.66 ; '720,000' -> 720.0"""
    s = (s or '').strip()
    if not s or not re.search(r'\d', s):
        return None
    try:
        return float(s.replace('.', '').replace(',', '.'))
    except ValueError:
        return None


SEM_SKIP = re.compile(r'^(SPECIALTY ELECTRONIC MATERIALS|SWITZERLAND GMBH|GROSSMATTE|6014 LUZERN|SWITZERLAND$|VAT REGISTRATION|Date \d|'
                      r'Place of issue|Item Quantity and description|® ™|USD/1)', re.I)
SEM_CHG = re.compile(r'^(FREIGHT|INSURANCE|HANDLING(?: CHARGE)?|PACKING(?: CHARGE)?|SURCHARGE|[A-Z][A-Z ]{2,30}CHARGES?)\s+(-?[\d.]+,\d{2})$')
SEM_UOM = {'KG': 'KGM', 'L': 'LTR', 'TU': 'TU', 'CAN': 'CAN', 'PC': 'PCE', 'EA': 'PCE', 'ST': 'PCE', 'M': 'MTR'}


def sem_pack(path, pages, C):
    """DuPont pack from BDP (e.g. 1197631758 - Docs): transmittal letter + Specialty Electronic Materials Switzerland
    "LEGAL I/C INVOICE" (one or more, European numbers 4.267,66) + ONE sea waybill + Consolidated Packing List.
    Invoice line: "n  4,000 Drum" / "MATERIAL  720,000 KG  5,927306  4.267,66" / description / Gross + Net Weight /
    item charges (FREIGHT 20,48 -> added to the item value) / "Our Reference" = customer order number / COO."""
    out, inv_pages, swb_pages = [], collections.OrderedDict(), []
    for p in pages:
        t = p['text'] or ''
        m = re.search(r'LEGAL I/C INVOICE\s+(\d{6,12})', t)
        if m:
            inv_pages.setdefault(m.group(1), []).append(p)
        elif re.search(r'SEA WAYBILL NO\.', t) or re.search(r'B/L NO\.?:\s*ONEY', t[:400]):
            swb_pages.append(p)
    # B/L "NET WEIGHT : x KG / HS-NO 3403198000" per product -> HS by net weight (fallback for items without HS)
    bl_hs = collections.defaultdict(set)
    swb_t = '\n'.join(p['text'] for p in swb_pages)
    for nw_, hs_ in re.findall(r'NET WEIGHT\s*:\s*([\d.,]+)\s*KG\s*\n\s*HS-NO\s*(\d{6,10})', swb_t):
        bl_hs[round(float(nw_.replace(',', '')), 3)].add(hs_)
    for no, pgs in inv_pages.items():
        t = '\n'.join(p['text'] for p in pgs)
        d = _inv(path, 'DuPont SEM legal invoice', 'SPECIALTY ELECTRONIC MATERIALS SWITZERLAND GMBH')
        d.update(invoice_no=no, ship_from_country='CH', title='LEGAL I/C INVOICE', direction='IMPORT')
        m = re.search(r'^Date\s+(\d{1,2})([A-Z]{3})(\d{4})', t, re.M)
        if m and m.group(2) in MONTHS:
            d['invoice_date'] = _dmy(m.group(1), MONTHS[m.group(2)], m.group(3))
        m = re.search(r'Customer Order Number.*\n\s*(\d{6,12})', t)
        d['customer_po'] = m.group(1) if m else ''
        m = re.search(r'Mail / Invoice to\s+Ship to\s*\n(.+)', t)
        nm = 'DUPONT DIVERSIFIED INDUSTRIALS SINGAPORE PTE. LTD.' if re.search(r'DUPONT DIVERSIFIED INDUSTRIALS', t) else (clean(m.group(1)).upper() if m else '')
        d['ship_to'] = d['invoice_to'] = nm
        d['ship_to_country'] = 'SG' if re.search(r'SINGAPORE\s+\d{6}', t) else ''
        m = re.search(r'VAT REG\. NR\.\s*(\d{8,9}[A-Z])', t)
        d['uen'] = m.group(1) if m else ''
        m = re.search(r'^(EXW|FCA|FAS|FOB|CFR|CIF|CPT|CIP|DAP|DPU|DDP)\s+(.+)$', t, re.M)
        if m:
            d['incoterm'], d['place'] = m.group(1), clean(m.group(2))
        m = re.search(r'\n(Sea|Air|Road|Rail)\s', t)
        d['transport'] = m.group(1).title() if m else ''
        m = re.search(r'Payment Terms/Due date\s+Currency\s*\n.*?\s(US Dollar|Euro|Singapore Dollar|[A-Z]{3})\s*$', t, re.M)
        d['currency'] = {'US DOLLAR': 'USD', 'EURO': 'EUR', 'SINGAPORE DOLLAR': 'SGD'}.get((m.group(1) if m else '').upper(), (m.group(1) if m else '') or 'USD')
        m = re.search(r'^Total\s+([\d.]+,\d{2})\s*$', t, re.M)
        d['total'] = _eu(m.group(1)) if m else None
        m = re.search(r'Shipment No\.:\s*(\d+)', t)
        d['shipment_no'] = m.group(1) if m else ''
        m = re.search(r'Freight Value\s+[A-Z]{3}\s+([\d.]+,\d{2})', t)
        d['freight_value'] = _eu(m.group(1)) if m else None
        cur, state, unread, chg_total = None, '', [], 0.0
        lines = [ln.strip() for ln in t.splitlines()]
        start = next((i for i, ln in enumerate(lines) if ln.startswith('Item Quantity and description')), 0)
        for ln in lines[start:]:
            if not ln or SEM_SKIP.match(ln):
                continue
            if re.match(r'^(Total Gross Weight|Subtotal|VAT %|Exempte|F\.O\.B\.|Freight/Border)', ln):
                cur = None
                continue
            m = re.match(r'^(\d{1,3})\s+([\d.]+,\d{3})\s+([A-Za-z]+)$', ln)
            if m:
                cur = dict(pos=m.group(1), pk=_eu(m.group(2)), pku=m.group(3).upper(), desc=[], charges=[])
                state = 'mat'
                continue
            if cur is None:
                continue
            m = re.match(r'^(\d{6,})\s+([\d.]+,\d{3})\s+([A-Z]{1,4})\s+([\d.]+,\d+)\s+(-?[\d.]+,\d{2})$', ln)
            if state == 'mat' and m:
                cur.update(material=m.group(1), qty=_eu(m.group(2)), uom=m.group(3), uprice=_eu(m.group(4)), value=_eu(m.group(5)))
                l = _line(item=m.group(1), desc='', qty=cur['qty'], uom=SEM_UOM.get(m.group(3), m.group(3)), amount=cur['value'],
                          currency=d['currency'], price=cur['uprice'])
                l['line'], l['material'], l['packages'], l['package_unit'] = cur['pos'], m.group(1), cur['pk'], cur['pku']
                l['goods_value'] = cur['value']
                cur['l'] = l
                d['lines'].append(l)
                state = 'desc'
                continue
            if 'l' not in cur:
                unread.append(ln)
                cur = None
                continue
            l = cur['l']
            m = re.match(r'^Gross Weight\s*:\s*([\d.,]+)\s*KG', ln)
            if m:
                l['gw'], state = _eu(m.group(1)), 'w'
                continue
            m = re.match(r'^Net Weight\s*:\s*([\d.,]+)\s*KG', ln)
            if m:
                l['nw'] = _eu(m.group(1))
                continue
            m = re.match(r'^Product Name:\s*(\S+)', ln)
            if m:
                l['ean'] = m.group(1)
                continue
            m = SEM_CHG.match(ln)
            if m:                                  # item charge (FREIGHT ...) -> part of the item value
                v = _eu(m.group(2))
                cur['charges'].append((m.group(1).title(), v))
                l['amount'] = round((l['amount'] or 0) + v, 2)
                l['charges'] = round((l.get('charges') or 0) + v, 2)
                l['charges_label'] = ', '.join(dict.fromkeys(x[0] for x in cur['charges']))
                chg_total += v
                continue
            m = re.match(r'^(-?[\d.]+,\d{2})$', ln)
            if m and cur['charges']:
                tot = _eu(m.group(1))
                if tot is not None and abs(tot - l['amount']) > 0.011:
                    d['notes'].append('line %s: value %s + charges = %s, invoice line total %s - check'
                                      % (l['line'], '{:,.2f}'.format(l['goods_value']), '{:,.2f}'.format(l['amount']), '{:,.2f}'.format(tot)))
                continue
            m = re.match(r'^Our Reference\s+(\d+)', ln)
            if m:
                l['our_ref'] = l['order'] = m.group(1)
                continue
            m = re.match(r'^Delivery note:\s*(\d+)', ln)
            if m:
                l['delivery'] = m.group(1)
                continue
            m = re.match(r'^Country Of Origin:\s*(.+?)(?:\s+\d+\s+PL)?$', ln)
            if m:
                if not l['coo']:
                    l['coo'] = C.code(clean(m.group(1))) or ''
                    if not l['coo']:
                        d['notes'].append('line %s: country "%s" not known - type the COO' % (l['line'], m.group(1)))
                continue
            if re.match(r'^(CUSTOMS INVOICE|Shipment No\.)', ln):
                continue
            if state == 'desc':
                cur['desc'].append(ln)
                l['desc'] = clean(re.sub(r'[®™]', '', ' '.join(cur['desc']))).upper()
        if unread:
            d['notes'].append('%d line(s) not read - check: %s' % (len(unread), ' | '.join(unread[:4])))
        for l in d['lines']:
            if l.get('qty'):
                l['price'] = round(l['amount'] / l['qty'], 6)
            if l.get('nw') is not None and bl_hs.get(round(l['nw'], 3)) and len(bl_hs[round(l['nw'], 3)]) == 1:
                l['bl_hs'] = next(iter(bl_hs[round(l['nw'], 3)]))
        d['total_gw'] = round(sum(l.get('gw') or 0 for l in d['lines']), 3) or None
        d['total_nw'] = round(sum(l.get('nw') or 0 for l in d['lines']), 3) or None
        d['boxes'] = int(sum(l.get('packages') or 0 for l in d['lines'])) or None
        units = {l.get('package_unit') for l in d['lines'] if l.get('package_unit')}
        d['package_unit'] = {'DRUM': 'DRM', 'CARTON': 'CTN', 'PAIL': 'PAL', 'BOX': 'BOX'}.get(next(iter(units)), 'PKG') if len(units) == 1 else 'PKG'
        if chg_total:
            d['notes'].append('item charges (%s) added to the item values: %s %s' % (
                ', '.join(dict.fromkeys(l['charges_label'] for l in d['lines'] if l.get('charges_label'))), d['currency'],
                '{:,.2f}'.format(chg_total)))
        d = _finish(d, C)
        d['notes'] = [n for n in d['notes'] if not n.startswith('no HS / tariff code on the invoice')]   # HS from the vendor list
        out.append(d)
    if swb_pages:
        try:
            r_ = one_swb(path, swb_pages, C)
            out += [x for x in (r_ if isinstance(r_, list) else [r_]) if x]
        except Exception as e:
            if out:
                out[0]['notes'].append('sea waybill pages not read: %s' % e)
    return out


# ======================================================================== Sandvik MCL invoice INTO Singapore (e.g. from Australia DC)
def sandvik_import_invoice(path, pages, C):
    """Sandvik Mining and Construction Logistics Ltd (Ireland) INVOICE for goods shipped TO Sandvik SMC Logistics
    Singapore from another warehouse (e.g. Pinkenba, Australia) - IMPORT. Line: "1 BG01249609 DOUBLE-ACTING 1 EA 2278.18
    168.000 166.000 2278.18" + rest of the description, then "Export HS code ... Import HS Code ... Country of Origin FI"."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Sandvik import invoice', 'SANDVIK MINING AND CONSTRUCTION LOGISTICS LTD')
    d['title'] = 'INVOICE'
    m = re.search(r'Invoice number\s+Order Number.*\n\s*(\d{6,})(?:\s+(\S+(?:\s*/\s*\S+)?))?', t, re.I)
    if m:
        d['invoice_no'], d['order_no'] = m.group(1), clean(m.group(2) or '')
    m = re.search(r'Invoice date\s+Customer Number.*\n\s*(\d{4}-\d{2}-\d{2})', t, re.I)
    d['invoice_date'] = _date_any(m.group(1)) if m else ''
    m = re.search(r'Consignee\s+Buyer\s*\n(.+)', t)
    if m:
        line = clean(m.group(1))
        half = len(line) // 2
        d['ship_to'] = d['invoice_to'] = (line[:half].strip() if line[:half].strip().upper() == line[half:].strip().upper() else line).upper()
    m = re.search(r'Country of destination\s*\n(.+)', t, re.I)
    d['ship_to_country'] = (C.find_in(m.group(1)) or ('SG' if 'SINGAPORE' in m.group(1).upper() else '')) if m else ''
    m = re.search(r'Country of Departure\s*\n(.+)', t, re.I)
    d['ship_from_country'] = (C.find_in(m.group(1).split()[-1]) or C.find_in(m.group(1)) or '') if m else ''
    m = re.search(r'Goods under this invoice delivered from:\s*\n(.+)\n(?:.*\n){0,3}?.*\b(Australia|New Zealand|Sweden|Finland|Germany|USA|United States|China|India|Chile|Peru|South Africa|Canada|Ireland)\b', t)
    if m:
        d['ship_from'] = clean(m.group(1)).upper()
        d['ship_from_country'] = d['ship_from_country'] or C.code(m.group(2)) or ''
    d['direction'] = 'IMPORT' if d.get('ship_to_country') == 'SG' else 'EXPORT'
    m = re.search(r'Terms of delivery\s+Incoterms\s*\n(.+)', t, re.I)
    mi = re.search(r'\b(EXW|FCA|FAS|FOB|CFR|CIF|CPT|CIP|DAP|DPU|DDP)\b\s*(.*)$', m.group(1)) if m else None
    if mi:
        d['incoterm'], d['place'] = mi.group(1), clean(mi.group(2))
    m = re.search(r'Despatch method.*\n\s*(Sea|Air|Road|Rail|Courier)\b', t, re.I)
    d['transport'] = m.group(1).title() if m else ''
    m = re.search(r'Carrier\s+MA reference.*\n\s*(.+?)(?:\s{2,}|\s+Sandvik|$)', t, re.M)
    d['carrier'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'^\s*([A-Z]{3})\s+KG\s+KG\s+([A-Z]{3})\s*$', t, re.M)
    d['currency'] = m.group(2) if m else ''
    LN = re.compile(r'^\s*(\d{1,4})\s+([A-Z0-9][A-Z0-9-]{4,})\s+(.*?)\s*(' + N + r')\s+([A-Z]{1,4})\s+(' + N + r')\s+(' + N + r')\s+('
                    + N + r')\s+(' + N + r')\s*$')
    rows = t.splitlines()
    cur, more = None, False
    for raw in rows:
        m = LN.match(raw)
        if m:
            cur = _line(line=m.group(1), item=m.group(2), desc=clean(m.group(3)).upper(), qty=_num(m.group(4)), uom=m.group(5),
                        price=_num(m.group(6)), nw=_num(m.group(7)), gw=_num(m.group(8)), amount=_num(m.group(9)), currency=d['currency'])
            d['lines'].append(cur)
            more = True
            continue
        if not cur:
            continue
        mh = re.search(r'Export HS code\s+(\d{6,10})?\s*Import HS Code\s+(\d{6,10})?\s*Country of Origin\s+([A-Z]{2})?', raw, re.I)
        if mh:
            more = False
            exp_hs, imp_hs = mh.group(1) or '', mh.group(2) or ''
            cur['hs'] = _hs8(imp_hs if d['direction'] == 'IMPORT' and imp_hs else exp_hs or imp_hs)
            cur['hs_export'] = exp_hs
            if mh.group(3):
                cur['coo'] = mh.group(3)
            continue
        if more and raw.strip() and not re.match(r'\s*(Customer PO|Packaging|Total|Line Item)', raw):
            cur['desc'] = clean('%s %s' % (cur['desc'], raw)).upper()
            continue
        mp = re.search(r'Dispatch Note\s+(\d+)\s*/\s*(\d+)(?:\s+Load ID\s+(\d+))?', raw, re.I)
        if mp:
            cur['po'], cur['order'] = mp.group(1), mp.group(2)
            if mp.group(3):
                d['delivery'] = mp.group(3)
            continue
        mk = re.search(r'Package Number\s+(\d+)', raw)
        if mk:
            cur['boxes'] = [mk.group(1)]
        if re.match(r'\s*Total Net Weight', raw):
            cur = None
    m = re.search(r'Total Net Weight\s+(' + N + r')\s*KG', t)
    d['total_nw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Gross Weight\s+(' + N + r')\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Amount\s+(' + N + r')', t)
    d['total'] = _num(m.group(1)) if m else None
    m = re.search(r'Total no of packages\s+(\d+)', t, re.I)
    if m:
        d['boxes'], d['package_unit'] = int(m.group(1)), 'PKG'
    if d.get('total_nw') and d.get('total_gw') and d['total_nw'] > d['total_gw']:
        d['notes'].append('net weight %s KG is MORE than the gross weight %s KG on the invoice - net weight as on the invoice, check'
                          % (d['total_nw'], d['total_gw']))
    d['no_rules'] = d['direction'] == 'IMPORT'     # Sandvik export rules (hub -> export permit) are not for this import
    return [_finish(d, C)]


def mainfreight_hbl(path, pages, C):
    """Mainfreight Air & Ocean express house bill of lading: document no. (HBL), CBR (carrier booking = ocean B/L), vessel /
    voyage, ports, packages, weight, container / seal / type, incoterm, shipped on board date."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'Mainfreight house B/L', 4)
    cg = {'mode': 'Sea'}
    m = re.search(r'BOOKING NUMBER\s+DOCUMENT NUMBER\s*\n(?:.*\n)?.*?CBR:\s*([A-Z0-9]{6,})\s+([A-Z]{2,4}\d{6,})', t)
    if m:
        d['mawb'], d['hawb'] = cg['mawb'], cg['hawb'] = m.group(1), m.group(2)
    m = re.search(r'MAIN VESSEL:\s*(.+?)\s*/\s*([A-Z0-9]+)', t)
    if m:
        cg['transport_id'], cg['flight'] = clean(m.group(1)).upper(), m.group(2)
    m = re.search(r'PORT OF LOADING\s+PORT OF DISCHARGE.*\n\s*([A-Z ]+),\s*[A-Z ]+?\s+([A-Z][A-Z ]+?)(?:\s{1,}[A-Z ]+)?$', t, re.M)
    if m:
        cg['loading_name'], cg['discharge_name'] = clean(m.group(1)).upper(), clean(m.group(2).split()[0]).upper()
    m = re.search(r'(\d+)\s+Package\(s\)\s+(' + N + r')\s*KG\s+(' + N + r')\s*M3', t)
    if m:
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = int(m.group(1)), 'PKG', _num(m.group(2))
        d['pieces'], d['gw'] = cg['packages_qty'], cg['gw_kg']
    conts = []
    for mc in re.finditer(r'^([A-Z]{4}\d{7})\s+(\S+)\s+(\d{2})([A-Z]+)\s+(' + N + r')\s+(' + N + r')\s+(\d+)\s+([A-Z]+)\s+(CFS|CY)/(CFS|CY)', t, re.M):
        lcl = mc.group(9) == 'CFS'
        conts.append(dict(no=mc.group(1), seal=mc.group(2), size=mc.group(3), size_text=('LCL' if lcl else 'FCL') + mc.group(3),
                          type=mc.group(3) + mc.group(4)[:2], tne=(_num(mc.group(5)) or 0) / 1000.0 or None))
    if conts:
        cg['containers'], cg['container_info'] = [c['no'] for c in conts], conts
    m = re.search(r'INCOTERM:\s*([A-Z]{3})', t)
    if m:
        d['awb_terms'] = m.group(1)
    m = re.search(r'SHIPPED ON BOARD\s*:\s*(\d{1,2}-[A-Za-z]{3}-\d{2,4})', t)
    if m:
        cg['etd'] = _dmon(m.group(1))
    m = re.search(r'SHIPPER/EXPORTER.*\n(.+?)(?:\s+CBR:|$)', t, re.M)
    d['shipper'] = clean(m.group(1)).upper() if m else ''
    d['issued_by'] = 'MAINFREIGHT AIR & OCEAN PTY LTD'
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]


# ======================================================================== Merck Healthcare KGaA billing document + packing list
def merck_kgaa_invoice(path, pages, C):
    """Merck Healthcare KGaA (Darmstadt) 'Invoice' / billing document to Merck Pte. Ltd. Singapore (IMPORT, e.g. ME
    permit). European numbers (23.995 PC, 219.314,30). Line "10 3008891501 CONCOR 2.5 23.995 PC 9,14 SGD / 1PC 219.314,30"
    + Batch / Country of origin / HS Code; totals: gross + net weight of the invoice (one line -> that line)."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Merck KGaA billing document', 'MERCK HEALTHCARE KGAA')
    d.update(title='INVOICE', ship_from_country='DE', invoice_from='MERCK HEALTHCARE KGAA')
    m = re.search(r'Billing document:\s*/\s*Date:\s*\n\s*(\d{6,12})\s*/\s*(\d{1,2})\.(\d{1,2})\.(\d{4})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _dmy(m.group(2), m.group(3), m.group(4))
    m = re.search(r'Sold to address:\s*\n(.+)', t)
    d['invoice_to'] = clean(m.group(1)).upper() if m else 'MERCK PTE. LTD.'
    m = re.search(r'Ship to:\s*\n(.+)', t)
    d['ship_to'] = clean(m.group(1)).upper() if m else d['invoice_to']
    d['ship_to_country'] = 'SG' if re.search(r'\bSINGAPORE\s+\d{6}', t) else ''
    d['direction'] = 'IMPORT' if d['ship_to_country'] == 'SG' else ''
    m = re.search(r'Incoterms:\s*([A-Z]{3})\b', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'Incoterms location:\s*(.+?)\s+Payment', t)
    d['place'] = clean(m.group(1)) if m else ''
    m = re.search(r'Currency:\s*([A-Z]{3})', t)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'Reference:\s*(\S+?),', t)
    d['customer_po'] = m.group(1) if m else ''
    m = re.search(r'Delivery n/date:\s*(\d+)', t)
    d['delivery'] = m.group(1) if m else ''
    m = re.search(r'Sales order:\s*(\d+)', t)
    d['order_no'] = m.group(1) if m else ''
    LN = re.compile(r'^(\d{1,4})\s+(\d{6,})\s+(.+?)\s+([\d.]+(?:,\d+)?)\s+([A-Z]{1,4})\s+([\d.]+,\d+)\s+([A-Z]{3})\s*/\s*([\d.]*)\s*([A-Z]{1,4})\s+(-?[\d.]+,\d{2})$')
    cur = None
    for raw in t.splitlines():
        ln = raw.strip()
        m = LN.match(ln)
        if m:
            per = _eu(m.group(8)) or 1.0
            qty, amt = _eu(m.group(4)), _eu(m.group(10))
            cur = _line(line=m.group(1), item=m.group(2), desc=clean(m.group(3)).upper(), qty=qty, uom=m.group(5),
                        price=round(amt / qty, 6) if qty else (_eu(m.group(6)) or 0) / per, amount=amt, currency=m.group(7),
                        order=d.get('delivery', ''))
            d['lines'].append(cur)
            continue
        if not cur:
            continue
        m = re.match(r'^Batch:\s*(\S+)', ln)
        if m:
            cur['lots'] = [m.group(1)]
            continue
        m = re.match(r'^Country of origin\s+([A-Z]{2})\b', ln)
        if m:
            cur['coo'] = m.group(1)
            continue
        m = re.match(r'^HS Code\s+(\d{6,10})', ln)
        if m:
            cur['hs'] = _hs8(m.group(1))
            continue
        if re.match(r'^(NET TOTAL PRICE|TOTAL PRICE)', ln):
            cur = None
    m = re.search(r'^TOTAL PRICE\s+(-?[\d.]+,\d{2})', t, re.M)
    d['total'] = _eu(m.group(1)) if m else None
    m = re.search(r'Total Gross Weight:\s*([\d.,]+)\s*KG', t)
    d['total_gw'] = _eu(m.group(1)) if m else None
    m = re.search(r'Total Net Weight:\s*([\d.,]+)\s*KG', t)
    d['total_nw'] = _eu(m.group(1)) if m else None
    if len(d['lines']) == 1:                   # one line: the invoice weights are that line's weights
        d['lines'][0]['nw'], d['lines'][0]['gw'] = d['total_nw'], d['total_gw']
    return [_finish(d, C)]


def merck_kgaa_pl(path, pages, C):
    """Merck Healthcare KGaA 'Packing List' (freight order): one row per pallet (delivery, pallet no., material,
    description, batch, qty, gross weight, volume) + Summary (material, total qty, NET weight, no. of pallets)."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Merck KGaA packing list')
    m = re.search(r'Freight order or Freight Booking Number\s+(\d+)', t)
    d['freight_order'] = m.group(1) if m else ''
    PAL = re.compile(r'^(?:(\d{10})\s+)?(\d{18,20})\s+(\d{6,})\s+(.+?)\s+([A-Z0-9]{5,8})\s+(\d+)\s+([A-Z]{1,3})\s+([\d.]+)\s*KG\s+([\d.]+)\s*M3')
    pal, deliv = [], ''
    for raw in t.splitlines():
        m = PAL.match(raw.strip())
        if m:
            deliv = m.group(1) or deliv
            pal.append(dict(delivery=deliv, pallet=m.group(2), item=m.group(3), batch=m.group(5), qty=float(m.group(6)),
                            uom=m.group(7), gw=float(m.group(8)), cbm=float(m.group(9))))
    summ = {}
    for m in re.finditer(r'(?:^|\n)(?:(.+)\n)?(\d{6,})\s+([A-Z0-9]{5,8})\s+(\d+)\s+([A-Z]{1,3})\s+([\d.]+)\s*KG\s+(\d+)\s*\n(.+)?', t):
        summ[m.group(2)] = dict(desc=clean('%s %s' % (m.group(1) or '', m.group(8) or '')).upper(), batch=m.group(3),
                                qty=float(m.group(4)), uom=m.group(5), nw=float(m.group(6)), pallets=int(m.group(7)))
    by = collections.OrderedDict()
    for x in pal:
        b = by.setdefault((x['delivery'], x['item']), dict(qty=0.0, gw=0.0, pallets=0, batch=x['batch'], uom=x['uom']))
        b['qty'] += x['qty']
        b['gw'] += x['gw']
        b['pallets'] += 1
    for (dl, item), b in by.items():
        s_ = summ.get(item, {})
        l = _line(item=item, desc=s_.get('desc', ''), qty=b['qty'], uom=b['uom'], order=dl, nw=s_.get('nw'), gw=round(b['gw'], 3))
        l['lots'] = [b['batch']]
        if s_ and abs((s_.get('qty') or 0) - b['qty']) > 0.001:
            d['notes'].append('material %s: pallets add up to %s, summary %s' % (item, b['qty'], s_.get('qty')))
        d['lines'].append(l)
    d['packages'], d['package_unit'] = (len(pal), 'PLT') if pal else (None, '')
    m = re.search(r'^Total\s+([\d.]+)\s+([\d.]+)\s*M3', t, re.M)
    d['total_gw'] = float(m.group(1)) if m else (round(sum(x['gw'] for x in pal), 3) or None)
    m = re.search(r'Total Net Weight\s+([\d.]+)\s*KG', t)
    d['total_nw'] = float(m.group(1)) if m else None
    d['invoice_refs'] = []
    if not d['lines']:
        d['notes'].append('no pallet lines could be read')
    return d


# ======================================================================== Organon (DGF OCEAN import)
def _organon_mon(s):
    m = re.match(r'(\d{1,2})-([A-Z]{3})-(\d{4})', s or '')
    return _dmy(m.group(1), MONTHS.get(m.group(2), 1), m.group(3)) if m and m.group(2) in MONTHS else ''


def organon_invoice(path, pages, C):
    """Organon International Services GmbH (Lucerne) 'Invoice' - one or more lines:
    "AERIUS 5MG 10TAB MYS 1007705" (or the item no. on its own line) + "[Tariff Number Exp: 3004900000] D143576
    0840164517781 20,539 EA 15.58 0.00 319,997.62 0.00" + Tariff Number Imp / Country of Origin.
    Document N° = the delivery = the packing list number (weights from the packing list)."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Organon invoice', 'ORGANON INTERNATIONAL SERVICES GMBH')
    d.update(title='INVOICE', ship_from_country='', direction='IMPORT', ship_to_country='SG')
    m = re.search(r'(?:Shipping Site|Destination Port|Incoterm.*?)\s(\d{8,12})\s+(\d{1,2}-[A-Z]{3}-\d{4})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _organon_mon(m.group(2))
    m = re.search(r'\b(\d{8,12})\s+(\d{6,10})\s+(\d{4})\s+\d+\s*/\s*\d+\s*$', t, re.M)
    if m:
        d['customer_po'], d['delivery'] = m.group(1), m.group(2).lstrip('0')
    m = re.search(r'Incoterm\s+([A-Z]{3})\b', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'Sold to.*\n(.+)', t)
    if m:
        s = clean(m.group(1)).upper()
        w = s.split()
        half = len(w) // 2                     # "ORGANON MALAYSIA SDN BHD ORGANON MALAYSIA SDN. BHD." -> first column
        d['invoice_to'] = ' '.join(w[:half]) if half and w[0] == w[half] else s
    m = re.search(r'^([A-Z]{3})\s+\1\s+\1\s*$', t, re.M)
    d['currency'] = m.group(1) if m else ''
    NL = re.compile(r'^(?:Tariff Number Exp:\s*(\d{6,12})\s+)?(\S+)\s+(\d{12,14})\s+([\d,]+(?:\.\d+)?)\s+([A-Z]{2,3})\s+([\d,]+\.\d+)\s+'
                    r'([\d.,]+)\s+([\d,]+\.\d{2})\s+[\d.]+\s*$')
    prev, cur, desc_buf = '', None, []
    started = False
    for raw in t.splitlines():
        ln = raw.strip()
        if not started:
            if d['currency'] and re.match(r'^%s\s+%s\s+%s$' % ((d['currency'],) * 3), ln):
                started = True
            continue
        if ln.startswith('****SHIPPING') or ln.startswith('Subtotal'):
            started = False
            cur = None
            continue
        m = NL.match(ln)
        if m:
            txt = ' '.join(desc_buf)
            mi = re.match(r'^(.*?)\s*\b(\d{6,8})$', txt)
            desc, item = (mi.group(1), mi.group(2)) if mi else (txt, '')
            qty, amt = _num(m.group(4)), _num(m.group(8))
            cur = _line(item=item, desc=clean(desc).upper(), qty=qty, uom=m.group(5), price=_num(m.group(6)), amount=amt,
                        currency=d['currency'], order=d.get('delivery', ''))
            cur['lots'] = [m.group(2)]
            if m.group(1):
                cur['hs'] = _hs8(m.group(1))
            d['lines'].append(cur)
            desc_buf = []
            continue
        m = re.match(r'^Tariff Number Exp:\s*(\d{6,12})', ln)
        if m and cur is not None:
            cur['hs'] = cur['hs'] or _hs8(m.group(1))
            continue
        if ln.startswith('Tariff Number Imp'):
            continue
        m = re.match(r'^Country of Origin:\s*([A-Z]{2})\b', ln)
        if m and cur is not None:
            cur['coo'] = m.group(1)
            continue
        if re.match(r'^(Mfg Date|Exp\. Date|Description|Unit Price)', ln):
            continue
        if cur is not None and cur.get('coo'):
            cur = None
        if cur is None:
            desc_buf.append(ln)
    m = re.search(r'Subtotal\s+([\d,]+(?:\.\d+)?)\s', t)
    d['total'] = _num(m.group(1)) if m else None
    if d['lines'] and not any(l['item'] for l in d['lines']):
        d['notes'].append('item numbers not found')
    return [_finish(d, C)]


def organon_pl(path, pages, C):
    """Organon packing lists: (a) KN Geel 'Packing List' (delivery 0802381241, Total Gross / Net Weight, Kg, rows
    "1007705 AERIUS 5MG 10TAB MYS D143576 01.06.2028 56.16 4320"); (b) N.V. Organon Oss 'Packing List' (Delivery #,
    pallets with Gross / Net Wgt, Material #, Qt2,400, Total Count Pallets 5 ... Gross Weight:820.000 KG Net Weight:747.500 KG).
    One line per material (whole delivery net / gross weight split by qty)."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Organon packing list')
    by = collections.OrderedDict()
    m = re.search(r'Delivery\s*#:\s*(\d{6,})', t) or re.search(r'Reference Organon:\s*(\d{6,})', t) or \
        re.search(r'Order number:.*\n\s*(0?8\d{8})\b', t) or \
        re.search(r'Packing\s*List.*\n.*?\b(0?8\d{8})\b', t)
    d['delivery'] = m.group(1).lstrip('0') if m else ''
    m = re.search(r'Total Gross Weight,?\s*Kg:?\s*([\d.,]+)', t, re.I) or re.search(r'Gross Weight:\s*([\d.,]+)\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Net Weight,?\s*Kg:?\s*([\d.,]+)', t, re.I) or re.search(r'Net Weight:\s*([\d.,]+)\s*KG', t)
    d['total_nw'] = _num(m.group(1)) if m else None
    m = re.search(r"(\d+)\s+HU'?s\s+([\d.,]+)\s*KG\s+([\d.,]+)\s*KG", t)       # DHL Supply Chain (Belgium) layout
    if m and not d['total_gw']:
        d['total_gw'], d['total_nw'] = _num(m.group(2)), _num(m.group(3))
    m = re.search(r'Number of Pallets:\s*(\d+)', t) or re.search(r'Total Count Pallets\s*(\d+)', t) or re.search(r"(\d+)\s+HU'?s\s", t)
    d['packages'], d['package_unit'] = (int(m.group(1)), 'PLT') if m else (None, '')
    for m in re.finditer(r'^(\d{6,8})\s+(.+?)\s+([A-Z]\d{5,})\s+\d{2}\.\d{2}\.\d{4}\s+([\d.]+)\s+(\d+)\s*$', t, re.M):
        b = by.setdefault(m.group(1), dict(desc=clean(m.group(2)).upper(), qty=0.0, batch=m.group(3)))
        b['qty'] += float(m.group(5))
    for m in re.finditer(r'Material\s*#\s*(\d{6,8})\s*\n(.+?)\s+GTIN.*\n\s*Batch\s+(\S+).*?Qt\s*([\d,]+)', t):
        b = by.setdefault(m.group(1), dict(desc=clean(m.group(2)).upper(), qty=0.0, batch=m.group(3)))
        b['qty'] += _num(m.group(4)) or 0
    for m in re.finditer(r'^\d{6}\s+(\d{6,8})\s+\S+\s+(.+?)\s+[A-Z]{3}-\d{4}\s+([\d,]+)\s*EA', t, re.M):
        b = by.setdefault(m.group(1), dict(desc=clean(m.group(2)).upper(), qty=0.0, batch=''))
        b['qty'] += _num(m.group(3)) or 0
    tq = sum(b['qty'] for b in by.values())
    for item, b in by.items():
        sh = b['qty'] / tq if tq else 1.0 / len(by)
        l = _line(item=item, desc=b['desc'], qty=b['qty'], uom='EA', order=d['delivery'],
                  nw=round(d['total_nw'] * sh, 3) if d['total_nw'] else None, gw=round(d['total_gw'] * sh, 3) if d['total_gw'] else None)
        l['lots'] = [b['batch']]
        d['lines'].append(l)
    if not d['lines']:
        d['notes'].append('no material lines could be read')
    return d


# ======================================================================== CommScope EMEA (Ireland) -> CommScope Singapore
def commscope_invoice(path, pages, C):
    """CommScope 'COMMERCIAL INVOICE' (EMEA Ireland, Asia Suzhou, Connectivity Suzhou): "990574891 000010 760259259
    24,000 FT 501.89000 MF 12,045.47" (order, PO line, material, qty, UOM, unit price per UOM, net amount) + material
    description [+ long description after the HTS line] + "Country Of Origin: CN HTS Code : 8544421100";
    totals row "Freight VAT|TAX Amount Material Total Surcharge Invoice Total Inv. Currency"; Shipment Number = the
    packaging summary of the same shipment (weights)."""
    t = '\n'.join(p['text'] for p in pages)
    m = re.search(r'Exporter\s*\n\s*(.+?)\s+(\d{9,12})\s*$', t, re.M)
    exporter = clean(m.group(1)).upper() if m else 'COMMSCOPE'
    d = _inv(path, 'CommScope invoice', exporter)
    cc = 'IE' if re.search(r'IRELAND', t[:600]) else 'CN' if re.search(r'CHINA|SUZHOU', t[:800], re.I) else ''
    d.update(title='COMMERCIAL INVOICE', ship_from_country=cc, direction='IMPORT', ship_to_country='SG')
    m = m or re.search(r'Commercial Invoice Number[\s\S]{0,160}?\b(\d{9,12})\b', t)
    d['invoice_no'] = m.group(m.lastindex) if m else ''
    m = re.search(r'Commercial Invoice Date[\s\S]{0,160}?\b(\d{1,2})\s*([A-Z]{3})\s*(\d{4})\b', t)
    if m and m.group(2) in MONTHS:
        d['invoice_date'] = _dmy(m.group(1), MONTHS[m.group(2)], m.group(3))
    m = re.search(r'Ship To\s*:\s*\d*.*\n(.+)', t)
    d['invoice_to'] = d['ship_to'] = 'COMMSCOPE SOLUTIONS SINGAPORE PTE LTD' if re.search(r'COMMSCOPE SOLUTIONS SINGAPORE', t) else (clean(m.group(1)).upper() if m else '')
    m = re.search(r'Incoterms.*\n([A-Z]{3})\b', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'Shipment Number\s*\n.*?(\d{6,})\s*$', t, re.M)
    d['shipment'] = d['delivery'] = m.group(1) if m else ''
    LN = re.compile(r'^(\d{6,12})\s+(\d{3,6})\s+(\S+)\s+([\d,]+(?:\.\d+)?)\s+([A-Z]{1,4})\s+([\d,]+\.\d+)\s+([A-Z]{1,3})\s+([\d,]+\.\d{2})\s*$')
    L = [x.strip() for x in t.splitlines()]
    cur, after = None, False
    for i, ln in enumerate(L):
        m = LN.match(ln)
        if m:
            cur = _line(item=m.group(3), desc=clean(L[i + 1]).upper() if i + 1 < len(L) else '', qty=_num(m.group(4)),
                        uom=m.group(5), amount=_num(m.group(8)), order='%s/%d' % (m.group(1), int(m.group(2))))
            cur['price'] = round(cur['amount'] / cur['qty'], 6) if cur['qty'] else None
            cur['po'] = m.group(1)
            d['lines'].append(cur)
            after = False
            continue
        if cur is None:
            continue
        mc = re.search(r'Country Of Origin:\s*([A-Z]{2})\b', ln)
        if mc:
            cur['coo'] = mc.group(1)
        mh = re.search(r'HTS Code\s*:\s*(\d{6,12})', ln)
        if mh:
            cur['hs'] = _hs8(mh.group(1))
            cur['hts'] = mh.group(1)
            after = True
            continue
        if after and ln and not re.match(r'^(This tariff|Freight|Purchase Order|Signature|Terms)', ln) and not cur.get('long_desc'):
            cur['long_desc'] = clean(ln)                 # long description printed after the HTS line
        if re.match(r'^(This tariff|Freight)', ln):
            after = False
    m = re.search(r'Inv\. Currency\s*\n\s*([\d,.]+)\s+[\d,.]+\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([A-Z]{3})', t)
    if m:
        d['total'], d['currency'] = _num(m.group(4)), m.group(5)
        fr = _num(m.group(1)) or 0
        if fr:
            d['freight'] = fr
    else:
        m = re.search(r'\b([A-Z]{3})\s*\n\s*(?:Signature|Terms and Conditions)', t)
        d['currency'] = m.group(1) if m else ''
    for l in d['lines']:
        l['currency'] = d['currency']
        if d['shipment']:                      # order line of this shipment (the same PO line can ship twice)
            l['order'] = '%s@%s' % (l['order'], d['shipment'])
    return [_finish(d, C)]


def commscope_pl(path, pages, C):
    """CommScope 'PACKAGING SUMMARY' (also printed without spaces: "100.000PC", "44.000/34.000KG", "**TOTALS**").
    Handling units (pallet / carton): "57905512 PLT NPC06UVDB-BL004F 1900.000 EA 8544421100 EAR99 350.535/317.466 LB ..." +
    "990585735/10 159.000/144.000 KG ..."; a mixed pallet lists one row per order line / material and the pallet weight
    once -> split by quantity. One line per order line + material (KG); TOTALS gross / net KG; Shipment No(s)."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'CommScope packaging summary')
    m = re.search(r'Shipment\s*No\(s\):.*\n\s*(\d{6,})', t)
    d['shipment'] = d['delivery'] = m.group(1) if m else ''
    HU = re.compile(r'^(\d{7,9})\s+(PLT|CTN|BOX|PKG|CRT|DRM|SKD|BDL|REL|CAS)\b(.*)$')
    PART = r'(\S+)\s+([\d,]+\.\d+)\s*([A-Z]{1,3})\s+(\d{8,12})\b'
    hus, cur = [], None
    for raw in t.splitlines():
        ln = raw.strip()
        m = HU.match(ln)
        if m:
            cur = dict(parts=[], gw=None, nw=None, pending=None)
            hus.append(cur)
            mp = re.match(r'\s*' + PART, m.group(3))
            if mp:
                cur['pending'] = dict(item=mp.group(1), qty=_num(mp.group(2)), uom=mp.group(3), hs=mp.group(4))
            continue
        if re.match(r'^\*\*\s*TOTALS', ln):
            cur = None
            continue
        if cur is None:
            continue
        mo = re.match(r'^(\d{6,})/(\d+)\s*(.*)$', ln)
        if not mo:
            continue
        rest = mo.group(3)
        mp = re.match(PART, rest)
        if mp:
            cur['parts'].append(dict(order='%s/%d' % (mo.group(1), int(mo.group(2))), item=mp.group(1), qty=_num(mp.group(2)), uom=mp.group(3), hs=mp.group(4)))
        elif cur.get('pending'):
            cur['parts'].append(dict(order='%s/%d' % (mo.group(1), int(mo.group(2))), **cur.pop('pending')))
            cur['pending'] = None
        mk = re.search(r'([\d,]+\.\d+)\s*/\s*([\d,]+\.\d+)\s*KG', rest)
        if mk:
            cur['gw'], cur['nw'] = _num(mk.group(1)), _num(mk.group(2))
    by = collections.OrderedDict()
    mixed = 0
    for hu in hus:
        if not hu['parts']:
            continue
        tq = sum(x['qty'] or 0 for x in hu['parts']) or len(hu['parts'])
        if len(hu['parts']) > 1:
            mixed += 1
        for x in hu['parts']:
            sh = (x['qty'] or 0) / tq if tq else 1.0 / len(hu['parts'])
            b = by.setdefault((x['order'], x['item']), dict(qty=0.0, gw=0.0, nw=0.0, uom=x['uom'], hs=x['hs'], has_w=False))
            b['qty'] += x['qty'] or 0
            if hu['gw'] is not None:
                b['gw'] += hu['gw'] * sh
                b['nw'] += hu['nw'] * sh
                b['has_w'] = True
    for (order, item), b in by.items():
        l = _line(item=item, qty=b['qty'], uom=b['uom'], order='%s@%s' % (order, d['shipment']) if d['shipment'] else order, nw=round(b['nw'], 3) if b['has_w'] else None,
                  gw=round(b['gw'], 3) if b['has_w'] else None)
        l['hs'] = _hs8(b['hs'])
        d['lines'].append(l)
    if mixed:
        d['notes'].append('%d mixed pallet(s) / carton(s): the package weight is split by quantity between its lines' % mixed)
    m = re.search(r'TOTALS[\s\S]*?([\d.,]+)\s*/\s*([\d.,]+)\s*KG', t)
    if m:
        d['total_gw'], d['total_nw'] = _num(m.group(1)), _num(m.group(2))
    m = re.search(r'Total\s*Number\s*of\s*Packages:\s*(\d+)\s*\((\d+)\s*([A-Z]+)', t)
    if m:
        d['packages'], d['package_unit'] = int(m.group(1)), 'PLT' if m.group(3).startswith('PAL') else 'CTN' if m.group(3).startswith('CART') else 'PKG'
    return d


# ======================================================================== Halliburton (Australia -> Singapore cross dock) + Crane Worldwide
def halliburton_invoice(path, pages, C):
    """Halliburton 'COMMERCIAL INVOICE (Consolidated)': "9033.00 US EAR99 NLR D162C6Q444 9509787875 1.00 EA 22,676.89
    22,676.89" (HS schedule no., origin, ECCN, licence, material, reference, qty, unit, price, value) + description;
    "EXW Total 22,676.89 USD"; charges rows (freight / insurance / other) added when not 0."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Halliburton invoice', 'HALLIBURTON')
    m = re.search(r'Shipper of Record.*\n(.+?),?\s+Halliburton', t)
    d['invoice_from'] = d['supplier'] = clean(m.group(1)).rstrip(',').upper() if m else 'HALLIBURTON'
    d.update(title='COMMERCIAL INVOICE', ship_from_country='AU' if re.search(r'AUSTRALIA', t[:800]) else '')
    m = re.search(r'^\s*(\d{6,8})\s*\n\s*COMMERCIAL INVOICE', t, re.M)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})', t)
    if m:
        d['invoice_date'] = _dmy(m.group(2), m.group(1), m.group(3))      # US date mm/dd/yyyy
    m = re.search(r'Intermediate Consignee.*\n.*?(Halliburton [A-Z][\w ]*?(?:Pte\.? Ltd|PTE\.? LTD))', t, re.I)
    d['invoice_to'] = d['ship_to'] = clean(m.group(1)).upper() if m else ''
    d['ship_to_country'] = 'SG' if re.search(r'Intermediate Consignee[\s\S]{0,300}SINGAPORE', t) else ''
    d['direction'] = 'IMPORT'
    m = re.search(r'Incoterms\s*\n\s*([A-Z]{3})\b', t)
    d['incoterm'] = m.group(1) if m else ''
    LN = re.compile(r'^(\d{4}\.\d{2}(?:\.\d+)?)\s+([A-Z]{2})\s+(\S+)\s+(\S+)\s+(\S+)\s+(\d{6,})\s+([\d,]+\.\d+)\s+([A-Z]{1,4})\s+([\d,]+\.\d+)\s+([\d,]+\.\d{2})\s*$')
    lines = [x.strip() for x in t.splitlines()]
    for i, ln in enumerate(lines):
        m = LN.match(ln)
        if not m:
            continue
        l = _line(item=m.group(5), desc=clean(lines[i + 1]).upper() if i + 1 < len(lines) else '', qty=_num(m.group(7)),
                  uom=m.group(8), price=_num(m.group(9)), amount=_num(m.group(10)), coo=m.group(2), order=m.group(6))
        l['hs'] = re.sub(r'\D', '', m.group(1))[:8]
        l['eccn'] = m.group(3)
        for x in lines[i + 2:i + 6]:
            if re.match(r'^(\*{5,}|FCC|Subtotal|\d{4}\.\d{2})', x):
                break
            if re.match(r'^\d{6,}$', x):
                l['box'] = x
            elif not re.match(r'^AMI Doc', x):
                l['desc2'] = clean(x).upper()
        d['lines'].append(l)
    m = re.search(r'^[A-Z]{3}\s+Total\s+([\d,]+\.\d{2})\s+([A-Z]{3})', t, re.M)
    if m:
        d['total'], d['currency'] = _num(m.group(1)), m.group(2)
    else:
        m = re.search(r'Subtotal\s+([\d,]+\.\d{2})\s+([A-Z]{3})', t)
        d['total'], d['currency'] = (_num(m.group(1)), m.group(2)) if m else (None, 'USD')
    for lab, key in (('International Freight Charges', 'freight'), ('Insurance', 'insurance'),
                     (r'Other Misc\. Charges[^\d]*', 'charges'), ('Domestic Freight Charges', 'dom_freight')):
        m = re.search(r'^' + lab + r'\s+([\d,]+\.\d{2})\s+[A-Z]{3}', t, re.M)
        v = _num(m.group(1)) if m else 0
        if v:
            d[key] = v
    m = re.search(r'Subtotal\s+([\d,]+\.\d{2})', t)
    if m and d.get('total') and _num(m.group(1)) != d['total']:
        d['goods_value'] = _num(m.group(1))
        d['total'] = d['goods_value']          # item lines = goods value; charges kept apart
    return [_finish(d, C)]


def halliburton_manifest(path, pages, C):
    """Halliburton 'Shipment Manifest / Packing List': box rows + material rows "D162C6Q444 1.00 EA 4 05.00 460.00 ..."
    (net weight printed with a gap: 4 05.00 = 405.00), totals "Total Gross (Loaded) Weight : 460.00 KG"."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Halliburton shipment manifest')
    m = re.search(r'Shipment:\s*(\d{6,})', t)
    d['shipment'] = m.group(1) if m else ''
    d['invoice_refs'] = [d['shipment']] if d['shipment'] else []
    seen = set()
    for m in re.finditer(r'^([A-Z0-9][A-Z0-9-]{4,})\s+([\d,]+\.\d+)\s+([A-Z]{1,4})\s+(\d[\d ]*\.\d{2})\s+([\d,]+\.\d{2})\s+(.+?)\s*$', t, re.M):
        k = (m.group(1), m.group(2))
        if k in seen:
            continue
        seen.add(k)
        d['lines'].append(_line(item=m.group(1), desc=clean(m.group(6)).upper(), qty=_num(m.group(2)), uom=m.group(3),
                                nw=_num(m.group(4).replace(' ', '')), gw=_num(m.group(5))))
    m = re.search(r'Total Gross \(Loaded\) Weight\s*:\s*([\d,.]+)\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    d['total_nw'] = round(sum(l['nw'] or 0 for l in d['lines']), 3) or None
    m = re.search(r'Total Number of Pieces[^:]*:\s*(\d+)', t)
    boxes = set(re.findall(r'^\d\s+(\d{8,})\s+[\d.]+\s*x', t, re.M))
    d['packages'] = int(m.group(1)) if m else (len(boxes) or None)
    d['package_unit'] = 'PKG'
    return d


def crane_bl(path, pages, C):
    """Crane Worldwide Logistics house B/L / notice of arrival - HBL number (MLCWPERH19044182), MOBL, vessel, weights."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'Crane Worldwide B/L / notice of arrival', 2)
    cg = {'mode': 'Sea'}
    m = re.search(r'B/L No\.\s*\n.*?\b([A-Z]{4}[A-Z0-9]{6,})\s*$', t, re.M) or re.search(r'HOBL\s+([A-Z0-9]{8,})', t) or \
        re.search(r'B/L NUMBER\s+([A-Z0-9]{8,})', t)
    if m:
        d['hawb'] = cg['hawb'] = m.group(1)
    m = re.search(r'MOBL\s+([A-Z0-9]{8,})', t)
    if m:
        d['mawb'] = cg['mawb'] = m.group(1)
    m = re.search(r'(\d+)\s+(?:Piece|PACKAGE)\(S\)\s+([\d,.]+)\s+([\d,.]+)', t, re.I)
    if m:
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = int(m.group(1)), 'PKG', _num(m.group(2))
        d['pieces'], d['gw'] = cg['packages_qty'], cg['gw_kg']
    m = re.search(r'Container/Seal No\s+([A-Z]{4}\d{7})', t)
    if m:
        cg['containers'] = [m.group(1)]
    m = re.search(r'Consignee\s*\n(.+)', t)
    d['consignee'] = clean(re.split(r'\s+Port of|\s+Place of', m.group(1))[0]).upper() if m else ''
    d['issued_by'] = 'CRANE WORLDWIDE LOGISTICS'
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]


# ======================================================================== Ningbo suppliers -> CommScope Singapore
def _cn_date(s):
    m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', s or '')
    if m:
        return _dmy(m.group(3), m.group(2), m.group(1))
    m = re.search(r'(\d{1,2})-([A-Za-z]{3})-(\d{2,4})', s or '')
    if m and m.group(2).upper() in MONTHS:
        return _dmy(m.group(1), MONTHS[m.group(2).upper()], m.group(3))
    return ''


def ningbo_fengmei(path, pages, C):
    """Ningbo Fengmei Precision Technology pack: China export declaration (商品编号 = HS 3926901000) + INVOICE
    ("1 8005079246 760249251 7200 USD 0.52400 USD 3,772.80": no, PO, part, qty, price, amount) + Packing List
    ("1 8005079246 760249251 7200 338.40 475.60 plt(...)(288Cartons)", "PACK IN 2PALLETS") + contract."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Ningbo Fengmei invoice pack', 'NINGBO FENGMEI PRECISION TECHNOLOGY CO., LTD')
    d.update(title='INVOICE', ship_from_country='CN', direction='IMPORT', ship_to_country='SG')
    m = re.search(r'INVOICE NO[：:]\s*([A-Z0-9-]{5,})', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'DATE OF INVOICE[：:]\s*(\S+)', t)
    d['invoice_date'] = _cn_date(m.group(1)) if m else ''
    m = re.search(r'SHIP TO[：:]\s*(.+?PTE\.? LTD\.?)', t)
    d['invoice_to'] = d['ship_to'] = clean(m.group(1)).upper() if m else ''
    d['incoterm'] = 'EXW' if re.search(r'EX-?WORKS', t) else ''
    hs = {}
    for m in re.finditer(r'^(\d{2})\s+(\d{10})\s+(\d[\d,]*)\s', t, re.M):           # declaration rows: item no, HS, qty
        hs[int(m.group(1))] = _hs8(m.group(2))
    lines = [x.strip() for x in t.splitlines()]
    inv_part = t.split('Packing List')[0]
    pl = {}
    for m in re.finditer(r'^(\d{1,3})\s+(\d{6,})\s+(\S+)\s+([\d,]+)\s+([\d,.]+)\s+([\d,.]+)\s', t, re.M):
        pl[(m.group(2), m.group(3))] = (_num(m.group(5)), _num(m.group(6)))
    for i, ln in enumerate(lines):
        m = re.match(r'^(\d{1,3})\s+(\d{6,})\s+(\S+)\s+([\d,]+)\s+([A-Z]{3})\s+([\d,.]+)\s+[A-Z]{3}\s+([\d,]+\.\d{2})$', ln)
        if not m or ln not in inv_part:
            continue
        desc = ' '.join(x for x in (lines[i - 1] if i else '', lines[i + 1] if i + 1 < len(lines) else '')
                        if x and not re.match(r'^(Item No|MADE IN|EX-WORKS|TOTAL)', x))
        nw, gw = pl.get((m.group(2), m.group(3)), (None, None))
        l = _line(line=m.group(1), item=m.group(3), desc=clean(desc).upper(), qty=_num(m.group(4)), uom='PCS',
                  price=_num(m.group(6)), amount=_num(m.group(7)), currency=m.group(5), order=m.group(2), nw=nw, gw=gw,
                  coo='CN' if re.search(r'MADE IN CHINA|ORGINAL COUNTRY: CHINA|原产国', t) else '')
        l['hs'] = hs.get(int(m.group(1)), '') or (list(hs.values())[0] if len(hs) == 1 else '')
        d['currency'] = m.group(5)
        d['lines'].append(l)
    m = re.search(r'TOTAL[：:]\s*[\d,]+\s*PCS\s+([A-Z]{3})\s+([\d,]+\.\d{2})', t)
    d['total'] = _num(m.group(2)) if m else None
    m = re.search(r'TOTAL[：:]\s*[\d,]+\s+([\d,.]+)\s+([\d,.]+)\s+Total', t)
    if m:
        d['total_nw'], d['total_gw'] = _num(m.group(1)), _num(m.group(2))
    m = re.search(r'PACK IN\s*(\d+)\s*PALLETS?', t, re.I)
    if m:
        d['boxes'], d['package_unit'] = int(m.group(1)), 'PLT'
    d['hs_from'] = 'China export declaration'
    return [_finish(d, C)]


def ningbo_safewell(path, pages, C):
    """Ningbo Safewell Xinghuo Imp&Exp: COMMERCIAL INVOICE "760259386 200 1200 960.00 US$3.400 US$4,080.00" (part, cartons,
    pcs, N.W., price, amount) under the description + PACKING LIST "760259386 200 1200 960.00 1370.00 7.56" (G.W.)."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Ningbo Safewell invoice', 'NINGBO SAFEWELL XINGHUO IMP&EXP CO.,LTD.')
    d.update(title='COMMERCIAL INVOICE', ship_from_country='CN', direction='IMPORT', ship_to_country='SG', currency='USD')
    m = re.search(r'INVOICE NO\.?:\s*([A-Z0-9-]{5,})', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'^DATE:\s*(\S+)', t, re.M)
    d['invoice_date'] = _cn_date(m.group(1)) if m else ''
    m = re.search(r'PO NO\.?:\s*(\S+)', t)
    d['customer_po'] = m.group(1) if m else ''
    m = re.search(r'^TO:\s*\n(.+)', t, re.M)
    d['invoice_to'] = d['ship_to'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'INCO TERMS:\s*([A-Z]{3})', t)
    d['incoterm'] = m.group(1) if m else ''
    inv, pl = t.split('PACKING LIST')[0], ''.join(t.split('PACKING LIST')[1:])
    gw = {m.group(1): _num(m.group(5)) for m in re.finditer(r'^(\S+)\s+(\d+)\s+([\d,]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d.]+)\s*$', pl, re.M)}
    lines = [x.strip() for x in inv.splitlines()]
    for i, ln in enumerate(lines):
        m = re.match(r'^(\S+)\s+(\d+)\s+([\d,]+)\s+([\d,.]+)\s+US\$\s*([\d,.]+)\s+US\$\s*([\d,]+\.\d{2})$', ln)
        if not m:
            continue
        l = _line(item=m.group(1), desc=clean(lines[i - 1]).upper() if i else '', qty=_num(m.group(3)), uom='PCS',
                  price=_num(m.group(5)), amount=_num(m.group(6)), nw=_num(m.group(4)), gw=gw.get(m.group(1)),
                  order=d['customer_po'], coo='CN' if re.search(r'COO:\s*CHINA', t) else '')
        l['cartons'] = int(m.group(2))
        d['lines'].append(l)
    m = re.search(r'TOTAL:.*?US\$\s*([\d,]+\.\d{2})', inv)
    d['total'] = _num(m.group(1)) if m else None
    m = re.search(r'TOTAL:\s*\**\s*\d+\s+[\d,]+\s+([\d,.]+)\s+([\d,.]+)', pl)
    if m:
        d['total_nw'], d['total_gw'] = _num(m.group(1)), _num(m.group(2))
    m = re.search(r'(\d+)\s*CTNS\s*=\s*(\d+)\s*PLTS', t)
    if m:
        d['boxes'], d['package_unit'] = int(m.group(2)), 'PLT'
    return [_finish(d, C)]


# ======================================================================== SharkNinja (SN APAC Holding HK) -> SharkNinja Singapore
def sharkninja_invoice(path, pages, C):
    """SN APAC Holding HK 'COMMERCIAL INVOICE': "Item Code TF200SMWH 804 $72.0634 $57,939.00" + Product / Long Description,
    Brand, HS Code (often blank), PO No.; "TOTAL $57,939.00"; currency from "Amount (SGD)"."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'SharkNinja invoice', 'SN APAC HOLDING HK LTD.')
    d.update(title='COMMERCIAL INVOICE', ship_from_country='HK', direction='IMPORT', ship_to_country='SG')
    m = re.search(r'Invoice No\.?:\s*(\S+)', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Date of Invoice:\s*(\S+)', t)
    d['invoice_date'] = _cn_date(m.group(1)) if m else ''
    m = re.search(r'MESSRS:\s*(.+?PTE\.? LTD\.?)', t)
    d['invoice_to'] = d['ship_to'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'Amount\s*\(([A-Z]{3})\)', t) or re.search(r'Unit Price\s*\(([A-Z]{3})\)', t)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'Incoterms\s+([A-Z]{3})\b', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'B/L No\.?:\s*(\S+)', t)
    d['hawb'] = m.group(1) if m else ''
    cur, pend = None, None
    AMT = r'(?:[A-Z]{0,3}\$\s*|[A-Z]{3}\s+)?'
    ITEM = re.compile(r'^Item\s*Code\s*:?\s*(\S+)\s+(?:(.*?)\s+)?([\d,]+(?:\.\d+)?)\s*(?:PCS|PC|EA|SET|UNITS?)?\s+' + AMT + r'([\d,]*\.?\d+)\s+' + AMT + r'([\d,]+\.\d{2})$', re.I)
    LAB = re.compile(r'^(Product\s*Desc(?:ription)?|Long\s*Desc(?:ription)?|Item\s*Desc(?:ription)?|Description|Brand(?:\s*Name)?|HS\s*Code|PO\s*No\.?)\s*:?\s*(.*)$', re.I)
    def put(k, v):
        k = k.upper().replace(' ', '')
        v = clean(v)
        if not v:
            return
        if k.startswith('LONGDESC'):
            cur['desc'], cur['_long'] = v.upper(), True
        elif 'DESC' in k:
            if not cur.get('_long'):
                cur['desc'] = v.upper()
        elif k.startswith('BRAND'):
            cur['brand'] = v.upper()
        elif k.startswith('HSCODE'):
            if re.search(r'\d{6}', v):
                cur['hs'] = _hs8(v)
        elif k.startswith('PONO'):
            cur['order'] = v
    for raw in t.splitlines():
        ln = raw.strip()
        if not ln:
            continue
        m = ITEM.match(ln)
        if m:
            cur, pend = _line(item=m.group(1), qty=_num(m.group(3)), uom='PCS', price=_num(m.group(4)), amount=_num(m.group(5))), None
            if m.group(2) and re.search(r'[A-Za-z]{3}', m.group(2)):
                cur['desc'] = clean(m.group(2)).upper()        # description printed on the item line
            d['lines'].append(cur)
            continue
        if cur is None:
            continue
        m = LAB.match(ln)
        if m:
            pend = None
            if clean(m.group(2)):
                put(m.group(1), m.group(2))
            else:
                pend = m.group(1)          # value on the next line
            continue
        if pend:
            if not re.match(r'^(TOTAL|Marks|Page\b|Remarks?|MADE IN)', ln, re.I):
                put(pend, ln)
            pend = None
    for l in d['lines']:
        l.pop('_long', None)
        if not l.get('desc'):
            d['notes'].append('CHECK item %s: no description on the invoice - type it on the Item tab' % l.get('item'))
    if re.search(r'MADE IN CHINA', t):
        for l in d['lines']:
            l['coo'] = l['coo'] or 'CN'
    m = re.search(r'^TOTAL\s+\$?([\d,]+\.\d{2})', t, re.M)
    d['total'] = _num(m.group(1)) if m else None
    return [_finish(d, C)]


def sharkninja_pl_xls(path, rows, C):
    """SN APAC Holding 'PACKING LIST' Excel: "Item Code: | TF200SMWH | container | qty | cartons | gross | net | cbm"."""
    import readers as R
    d = R.new_pl(path, 'SharkNinja packing list')
    for r in rows:
        v = [x for x in r if x not in (None, '')]
        if not v:
            continue
        if str(v[0]).startswith('Invoice No') or (len(v) > 3 and str(v[2]).startswith('Invoice No')):
            d['invoice_refs'] = [str(v[-1])]
        if str(v[0]).strip().rstrip(':') == 'Item Code' and len(v) >= 7:
            l = _line(item=str(v[1]).strip(), qty=fnum(v[3]), uom='PCS', gw=fnum(v[5]), nw=fnum(v[6]))
            l['container'] = str(v[2])
            d['lines'].append(l)
            if re.match(r'[A-Z]{4}\d{7}$', str(v[2])):
                d['containers'] = list(dict.fromkeys(d['containers'] + [str(v[2])]))
        if str(v[0]).strip().upper() == 'TOTAL' and len(v) >= 5:
            d['packages'], d['package_unit'] = int(fnum(v[2]) or 0) or None, 'CTN'
            d['total_gw'], d['total_nw'] = fnum(v[3]), fnum(v[4])
        if str(v[0]).strip() == 'B/L No.' and len(v) > 1:
            d['hawb'] = str(v[1]).strip()
    return [d]


# ======================================================================== Shanghai Jianlan Electric -> Sandvik Mining & Construction Logistics
CN_WORDS = {'传感器': 'SENSOR', '线束': 'WIRING HARNESS', '电缆': 'CABLE', '喇叭': 'HORN', '开关': 'SWITCH', '接头': 'CONNECTOR',
            '螺栓': 'BOLT', '螺母': 'NUT', '垫圈': 'WASHER', '轴承': 'BEARING', '密封件': 'SEAL', '过滤器': 'FILTER'}


def cn_desc(s):
    """Chinese part names -> English (customs description must be English)."""
    s = s or ''
    for k, v in CN_WORDS.items():
        s = s.replace(k, ' %s ' % v)
    return clean(s).upper()


def jianlan_invoice(path, pages, C):
    """Shanghai Jianlan Electric 'COMMERCIAL INVOICE': "1 China 8544421100 HORN BG01405624B 4500667892 10 70 288.40 20188.00"
    (no., COO, HS, description, part no., PO no., PO line, qty, unit price, amount, CNY)."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Shanghai Jianlan invoice', 'SHANGHAI JIANLAN ELECTRIC SCIENCE DEVELOPMENT CO., LTD')
    d.update(title='COMMERCIAL INVOICE', ship_from_country='CN', direction='IMPORT', ship_to_country='SG')
    m = re.search(r'\b(JL\d{6,})\s+([A-Za-z]{3})\.?(\d{1,2})\.(\d{4})', t)
    if m:
        d['invoice_no'] = m.group(1)
        d['invoice_date'] = _dmy(m.group(3), MONTHS.get(m.group(2).upper(), 1), m.group(4))
    d['invoice_to'] = d['ship_to'] = 'SANDVIK MINING & CONSTRUCTION LOGISTICS' if re.search(r'Sandvik Mining', t) else ''
    m = re.search(r'\((CNY|USD|EUR)\)', t)
    d['currency'] = m.group(1) if m else 'CNY'
    m = re.search(r'\b(FOB|EXW|FCA|CIF|CFR|DAP)\b\s+\d+\s*Days', t)
    d['incoterm'] = m.group(1) if m else ''
    LN = re.compile(r'^(\d{1,3})\s+([A-Za-z][A-Za-z ]*?)\s+(\d{8,10})\s+(.+?)\s+(\S+)\s+(\d{8,12})\s+(\d{1,5})\s+([\d,]+)\s+([\d,.]+)\s+([\d,]+\.\d{2})$')
    for raw in t.splitlines():
        m = LN.match(raw.strip())
        if not m:
            continue
        coo = C.code(m.group(2)) if hasattr(C, 'code') else ''
        l = _line(line=m.group(1), item=m.group(5), desc=cn_desc(m.group(4)), qty=_num(m.group(8)), uom='PCS',
                  price=_num(m.group(9)), amount=_num(m.group(10)), order=m.group(6), coo=coo or ('CN' if m.group(2).upper() == 'CHINA' else ''))
        l['hs'] = _hs8(m.group(3))
        l['po_line'] = m.group(7)
        d['lines'].append(l)
    m = re.search(r'TOTAL:\s*[￥¥$]?\s*([\d,]+\.\d{2})', t)
    d['total'] = _num(m.group(1)) if m else None
    return [_finish(d, C)]


def jianlan_pl_xls(path, rows, C):
    """Shanghai Jianlan 'PACKING LIST' Excel: rows DESCRIPTION | PARTS NO. | PO NO. | ORDER QTY | DELIVERED QTY | packages |
    GROSS | NET; one line per PO + part (cartons added up)."""
    import readers as R
    d = R.new_pl(path, 'Shanghai Jianlan packing list')
    hi = next((i for i, r in enumerate(rows) if any('PARTS NO' in str(x or '').upper() for x in r)
               and any('NET WEIGHT' in str(x or '').upper() for x in r)), None)
    if hi is None:
        d['notes'].append('packing list heading not found')
        return [d]
    H = [re.sub(r'\s+', ' ', str(x or '')).upper() for x in rows[hi]]
    col = lambda *ws: next((i for i, h in enumerate(H) if any(w in h for w in ws)), None)
    cd, cp, co, cq, ck, cg, cn = col('DESCRIPTION'), col('PARTS NO'), col('PO NO'), col('DELIVERED'), col('PACKAGES'), col('GROSS'), col('NET')
    by = collections.OrderedDict()
    for r in rows[hi + 1:]:
        g = lambda c: r[c] if c is not None and c < len(r) else None
        if any(str(x or '').strip().upper().startswith('TOTAL') for x in r):
            vals = [x for x in r if isinstance(x, (int, float))]
            if len(vals) >= 4:
                d['packages'], d['package_unit'] = int(vals[1]), 'CTN'
                d['total_gw'], d['total_nw'] = float(vals[2]), float(vals[3])
            break
        part = str(g(cp) or '').strip()
        if not part:
            continue
        k = (str(g(co) or '').strip(), part)
        b = by.setdefault(k, dict(desc=cn_desc(str(g(cd) or '')), qty=0.0, gw=0.0, nw=0.0, pk=0))
        b['qty'] += fnum(g(cq)) or 0
        b['gw'] += fnum(g(cg)) or 0
        b['nw'] += fnum(g(cn)) or 0
        b['pk'] += int(fnum(g(ck)) or 0)
    for (po, part), b in by.items():
        d['lines'].append(_line(item=part, desc=b['desc'], qty=b['qty'], uom='PCS', order=po, nw=round(b['nw'], 3), gw=round(b['gw'], 3)))
    for r in rows:
        for x in r:
            m = re.search(r'INVOICE NO\.?\s*\n?\s*(\S+)', str(x or ''))
            if m:
                d['invoice_refs'] = [m.group(1)]
    return [d]


# ======================================================================== Allison (Wuxi) Drive Systems -> Sandvik Mining & Construction Logistics
def allison_invoice(path, pages, C):
    """Allison (Wuxi) 'Declaration Invoice': "1 BEARING CUP 4500647362 247570 04696103 75 16.68 1251.00 #N/A" (no., description,
    customer PO, Dana PN, customer PN, qty, price, amount, HS code or #N/A); "COUNTRY OF ORIGIN: CHINA"."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Allison (Wuxi) invoice', 'ALLISON (WUXI) DRIVE SYSTEMS CO., LTD.')
    d.update(title='INVOICE', ship_from_country='CN', direction='IMPORT', ship_to_country='SG')
    m = re.search(r'Invoice No\.?:\s*(\S+)', t)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Invoice Date:\s*(\S+)', t)
    d['invoice_date'] = _cn_date(m.group(1)) if m else ''
    m = re.search(r'Ship to:\s*(.+)', t)
    d['invoice_to'] = d['ship_to'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'Transportation Terms:\s*([A-Z]{3})\b', t) or re.search(r'Transportation Terms:\s*(Ex\s*work)', t, re.I)
    d['incoterm'] = ('EXW' if m and m.group(1).upper().startswith('EX') else m.group(1)) if m else ''
    m = re.search(r'Amount\s*\n?\s*\(([A-Z]{3})\)', t) or re.search(r'\(([A-Z]{3})\)', t)
    d['currency'] = m.group(1) if m else 'USD'
    coo = 'CN' if re.search(r'COUNTRY OF ORIGIN\s*[：:]\s*CHINA', t) else ''
    LN = re.compile(r'^(\d{1,3})\s+(.+?)\s+(\d{10})\s+(\S+)\s+(\d{6,10})\s+([\d,]+)\s+([\d,.]+)\s+([\d,]+\.\d{2})\s+(\S+)\s*$')
    for raw in t.splitlines():
        m = LN.match(raw.strip())
        if not m:
            continue
        l = _line(line=m.group(1), item=m.group(5), desc=clean(m.group(2)).upper(), qty=_num(m.group(6)), uom='PCS',
                  price=_num(m.group(7)), amount=_num(m.group(8)), order=m.group(3), coo=coo)
        l['maker_part'] = m.group(4)
        l['hs'] = _hs8(m.group(9)) if re.match(r'\d{6,}', m.group(9)) else ''
        d['lines'].append(l)
    m = re.search(r'^Total:\s*[\d,]+\s+([\d,]+\.\d{2})', t, re.M)
    d['total'] = _num(m.group(1)) if m else None
    d['no_rules'] = True                       # import into Singapore: no Sandvik export rules
    return [_finish(d, C)]


def allison_pl(path, pages, C):
    """Allison (Wuxi) 'Packing List': "[1 120*80*56CM 0.54 CBM] GEAR 2118510 04699775 2 56.00 [183] 4500724420"
    (description, Dana PN, customer PN, qty, net KG, pallet gross KG, customer PO); TOTAL: 3 PALLETS ... 338.26 429.00."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Allison (Wuxi) packing list')
    m = re.search(r'Invoice No\.?:\s*(\S+)', t)
    d['invoice_refs'] = [m.group(1)] if m else []
    by = collections.OrderedDict()
    for raw in t.splitlines():
        ln = re.sub(r'^\d{1,2}\s+\d+\*\d+\*\d+\s*CM\s+[\d.]+\s*CBM\s+', '', raw.strip())
        m = re.match(r'^(.+?)\s+(\S+)\s+(\d{8})\s+([\d,]+)\s+([\d.]+)(?:\s+([\d.]+))?\s+(\d{10})$', ln)
        if not m:
            continue
        b = by.setdefault((m.group(7), m.group(3)), dict(desc=clean(m.group(1)).upper(), qty=0.0, nw=0.0))
        b['qty'] += _num(m.group(4)) or 0
        b['nw'] += _num(m.group(5)) or 0
    for (po, pn), b in by.items():
        d['lines'].append(_line(item=pn, desc=b['desc'], qty=b['qty'], uom='PCS', order=po, nw=round(b['nw'], 3)))
    m = re.search(r'TOTAL\s*[：:]\s*(\d+)\s*PALLETS?\s+[\d.]+\s*CBM\s+[\d,]+\s+([\d.]+)\s+([\d.]+)', t)
    if m:
        d['packages'], d['package_unit'], d['total_nw'], d['total_gw'] = int(m.group(1)), 'PLT', _num(m.group(2)), _num(m.group(3))
    return d


# ======================================================================== Vertiv Tech (China) -> Vertiv Singapore (Excel) + Sinotrans B/L / arrival notice
def _xls_label(rows, label):
    """Value next to a label cell ("Invoice No.:" -> the next filled cell on that row)."""
    for r in rows:
        for i, x in enumerate(r):
            if isinstance(x, str) and x.strip().rstrip(':：').strip().upper() == label.upper():
                for y in r[i + 1:]:
                    if y not in (None, '') and str(y).strip():
                        return y
    return None


def _xls_date(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime('%d/%m/%Y')
    s = str(v or '')
    m = re.match(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', s)
    if m:
        return _dmy(m.group(3), m.group(2), m.group(1))
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    return _dmy(m.group(2), m.group(1), m.group(3)) if m else ''         # Vertiv packing list: mm/dd/yyyy


def vertiv_invoice_xls(path, rows, C):
    """Vertiv Tech 'CUSTOMS INVOICE' Excel: Item No | Part No | [HS] | Description | Quantity (EA) | Unit Price (CNY) |
    Total Amount (CNY) | [HS]; Trade Terms; Country of origin."""
    d = _inv(path, 'Vertiv customs invoice', '')
    sup = str(rows[0][0] if rows and rows[0] else '').strip() or 'VERTIV TECH CO.,LTD.'
    d.update(supplier=sup.upper(), invoice_from=sup.upper(), ship_from=sup.upper(), title='CUSTOMS INVOICE',
             ship_from_country='CN', direction='IMPORT', ship_to_country='SG')
    d['invoice_no'] = str(_xls_label(rows, 'Invoice No.') or _xls_label(rows, 'Invoice No') or '').strip()
    d['invoice_date'] = _xls_date(_xls_label(rows, 'Date'))
    d['customer_po'] = str(_xls_label(rows, 'P.O No.') or '').strip()
    bt = _xls_label(rows, 'BILL TO')
    d['invoice_to'] = d['ship_to'] = re.split(r'\s+FORMERLY\b|\s+\d', str(bt or '').upper())[0].strip()
    txt = ' '.join(str(x) for r in rows for x in r if isinstance(x, str))
    m = re.search(r'Trade Terms:\s*([A-Z]{3})', txt)
    d['incoterm'] = m.group(1) if m else ''
    coo = 'CN' if re.search(r'Country of origin:\s*China', txt, re.I) else ''
    hi = next((i for i, r in enumerate(rows) if any(str(x or '').strip().upper().startswith('PART NO') for x in r)
               and any('AMOUNT' in str(x or '').upper() for x in r)), None)
    if hi is None:
        d['notes'].append('item heading not found')
        return [_finish(d, C)]
    H = [re.sub(r'\s+', ' ', str(x or '')).strip().upper() for x in rows[hi]]
    col = lambda f: next((i for i, h in enumerate(H) if f(h)), None)
    cp, cd, cq = col(lambda h: h.startswith('PART NO')), col(lambda h: h.startswith('DESCRIPTION')), col(lambda h: h.startswith('QUANTITY'))
    cu, ca, ch = col(lambda h: h.startswith('UNIT PRICE')), col(lambda h: h.startswith('TOTAL AMOUNT')), col(lambda h: h.startswith('HS'))
    m = re.search(r'\(([A-Z]{3})\)', H[ca] if ca is not None else '')
    d['currency'] = m.group(1) if m else 'CNY'
    m = re.search(r'\(([A-Z]{2,4})\)', H[cq] if cq is not None else '')
    uom = m.group(1) if m else 'PCS'
    for r in rows[hi + 1:]:
        g = lambda c: r[c] if c is not None and c < len(r) else None
        first = str(next((x for x in r if x not in (None, '')), '') or '').strip().upper()
        if first.startswith('TOTAL'):
            nums = [x for x in r if isinstance(x, (int, float))]
            d['total'] = float(nums[1]) if len(nums) > 1 else None
            break
        if not g(cp) or g(ca) in (None, ''):
            continue
        l = _line(item=str(g(cp)).strip(), desc=clean(str(g(cd) or '')).upper(), qty=fnum(g(cq)), uom=uom, price=fnum(g(cu)),
                  amount=fnum(g(ca)), order=d['customer_po'], coo=coo)
        hs = re.sub(r'\D', '', str(int(g(ch)) if isinstance(g(ch), float) else g(ch) or ''))
        l['hs'] = (hs + '00')[:8] if len(hs) == 6 else hs[:8]
        d['lines'].append(l)
    return [_finish(d, C)]


def vertiv_pl_xls(path, rows, C):
    """Vertiv Tech 'PACKING LIST' Excel: Package No. | Part No. | Description | QTY | G.Wt. | Net.Wt. | Dim | Po No.
    Weights per package "227.00*2" (x packages); a package holding several parts has the weight on its first row only
    -> split by quantity. One line per part."""
    import readers as R
    d = R.new_pl(path, 'Vertiv packing list')
    inv = str(_xls_label(rows, 'Invoice No') or _xls_label(rows, 'Invoice No.') or '').strip()
    d['invoice_refs'] = [inv] if inv else []
    hi = next((i for i, r in enumerate(rows) if any(str(x or '').strip().upper().startswith('PACKAGE NO') for x in r)), None)
    if hi is None:
        d['notes'].append('heading not found')
        return [d]
    H = [re.sub(r'\s+', ' ', str(x or '')).strip().upper() for x in rows[hi]]
    col = lambda w: next((i for i, h in enumerate(H) if h.startswith(w)), None)
    cpk, cp, cd, cq, cg, cn, co = col('PACKAGE'), col('PART'), col('CONTENT'), col('QTY'), col('G.WT'), col('NET'), col('PO')

    def w(v):
        s = str(v or '').replace('\u202d', '').replace('\u202c', '').strip()
        m = re.match(r'([\d.]+)\s*(?:\*\s*(\d+))?', s)
        return (float(m.group(1)) * (int(m.group(2)) if m.group(2) else 1)) if m else None
    pk, by, last_po = [], collections.OrderedDict(), ''
    for r in rows[hi + 1:]:
        g = lambda c: r[c] if c is not None and c < len(r) else None
        first = str(next((x for x in r if x not in (None, '')), '') or '').strip().upper()
        if first.startswith('TOTAL'):
            nums = [x for x in r if isinstance(x, (int, float))]
            m = re.search(r'TOTAL\s+(\d+)\s*Package', first, re.I)
            d['packages'], d['package_unit'] = (int(m.group(1)), 'PKG') if m else (None, '')
            if len(nums) >= 3:
                d['total_gw'], d['total_nw'] = float(nums[1]), float(nums[2])
            break
        if not g(cp):
            continue
        if g(cg) not in (None, '') or not pk:
            pk.append(dict(gw=w(g(cg)) or 0, nw=w(g(cn)) or 0, parts=[]))
        part = str(g(cp)).strip()
        pk[-1]['parts'].append((part, fnum(g(cq)) or 0))
        po = str(g(co) or '').replace('\u202d', '').replace('\u202c', '').strip() or last_po   # merged cell: PO above
        last_po = po
        b = by.setdefault(part, dict(desc=clean(str(g(cd) or '')).upper(), qty=0.0, gw=0.0, nw=0.0, po=po))
        b['qty'] += fnum(g(cq)) or 0
    for p in pk:
        tq = sum(q for _, q in p['parts']) or len(p['parts'])
        for part, q in p['parts']:
            sh = (q / tq) if tq else 1.0 / len(p['parts'])
            by[part]['gw'] += p['gw'] * sh
            by[part]['nw'] += p['nw'] * sh
    for part, b in by.items():
        d['lines'].append(_line(item=part, desc=b['desc'], qty=b['qty'], uom='EA', order=b['po'], nw=round(b['nw'], 3), gw=round(b['gw'], 3)))
    if len(pk) > 1 and any(len(p['parts']) > 1 for p in pk):
        d['notes'].append('several parts in one package - package weight split by quantity')
    return [d]


def sinotrans_bl(path, pages, C):
    """Sinotrans house B/L (first line = B/L no. STL26067044) + attachment (INV# list, HS) / Sinotrans ARRIVAL NOTICE
    (House B/L, Ocean B/L, vessel / voyage, ETA, packages, weight, container)."""
    t = '\n'.join(p['text'] for p in pages)
    an = 'ARRIVAL NOTICE' in t[:600]
    d = _awb(path, 'Sinotrans arrival notice' if an else 'Sinotrans house B/L', 2 if an else 3)
    cg = {'mode': 'Sea'}
    m = re.search(r'House B/L\s*:\s*([A-Z0-9]{6,})', t) or re.search(r'B/L NO\.\s*([A-Z0-9]{6,})', t) or re.match(r'\s*([A-Z]{3}\d{6,})\s*\n', t)
    if m:
        d['hawb'] = cg['hawb'] = m.group(1)
    m = re.search(r'Ocean B/L\s*:\s*([A-Z0-9]{8,})', t)
    if m:
        d['mawb'] = cg['mawb'] = m.group(1)
    m = re.search(r'Vessel\s*:\s*(.+?)\s*/\s*(\S+)', t) or re.search(r'^([A-Z][A-Z .]+?)/(\d{3,5}[A-Z]?)\s+[A-Z]', t, re.M)
    if m:
        cg['transport_id'], cg['flight'] = clean(m.group(1)).upper(), m.group(2)
    m = re.search(r'ETA\s*:\s*([A-Za-z]{3})\s+(\d{1,2}),\s*(\d{4})', t)
    if m and m.group(1).upper() in MONTHS:
        cg['eta'] = _dmy(m.group(2), MONTHS[m.group(1).upper()], m.group(3))
    m = re.search(r'SHIPPED ON BOARD\s*\n\s*(\d{4})-(\d{2})-(\d{2})', t)
    if m:
        cg['etd'] = _dmy(m.group(3), m.group(2), m.group(1))
    m = re.search(r'(\d+)\s+PACKAGES\s+([\d,.]+)\s*KGS', t) or re.search(r'Measurement:.*\n.*?(\d+)\s+([\d,.]+)\s*kg', t)
    if m:
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = int(m.group(1)), 'PKG', _num(m.group(2))
        d['pieces'], d['gw'] = cg['packages_qty'], cg['gw_kg']
    conts = list(dict.fromkeys(re.findall(r'\b([A-Z]{4}\d{7})\b', t)))
    if conts:
        cg['containers'] = conts
    m = re.search(r'\bHS:\s*(\d{6,10})', t)
    if m:
        d['awb_hs'] = m.group(1)
    refs = re.findall(r'\b([0-9A-Z]{8,}-[0-9A-Z]{1,3})\b', t)
    d['invoice_refs'] = list(dict.fromkeys(refs))
    m = re.search(r'Port\s*:\s*([A-Z ]+),', t) or re.search(r'PORT OF LOADING\s+([A-Z ]+?)\s*\n', t)
    if m:
        cg['loading_name'] = clean(m.group(1)).upper()
    d['issued_by'] = 'SINOTRANS (SG) PTE LTD'
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]


# ======================================================================== Maquet (Suzhou) / Getinge -> Getinge South East Asia (import)
def maquet_invoice(path, pages, C):
    """Maquet (Suzhou) 'Invoice': item rows "100 SU000900003 2 EA" + description + "703.84 1,407.67" (unit / total price)
    + "Commodity code 94029000" + "Country of origin CN"; Order no. (= packing list no.), "Item sum"."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Maquet (Suzhou) invoice', 'MAQUET (SUZHOU) CO., LTD.')
    d.update(title='INVOICE', ship_from_country='CN', direction='IMPORT', ship_to_country='SG')
    m = re.search(r'Maquet \(Suzhou\) Co\., Ltd\..*?\b(\d{8,12})\s+(\d{2})\.(\d{2})\.(\d{4})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _dmy(m.group(2), m.group(3), m.group(4))
    m = re.search(r'Customer:\s*\d+\s*\n(.+)', t)
    d['invoice_to'] = d['ship_to'] = clean(m.group(1)).upper() if m else ''
    m = re.search(r'Order no\.:\s*(\d+)', t)
    d['delivery'] = d['order_no'] = m.group(1) if m else ''
    m = re.search(r'Purchase order no\.:\s*(\S+)', t)
    d['customer_po'] = m.group(1) if m else ''
    m = re.search(r'Description\s+([A-Z]{3})\s+\1', t)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'Terms of delivery:\s*([A-Z]{3})', t)
    d['incoterm'] = m.group(1) if m else ''
    L = [x.strip() for x in t.splitlines()]
    cur = None
    for i, ln in enumerate(L):
        m = re.match(r'^(\d{2,5})\s+([A-Z0-9][A-Z0-9-]{5,})\s+([\d,.]+)\s+([A-Z]{1,4})$', ln)
        if m:
            cur = _line(line=m.group(1), item=m.group(2), qty=_num(m.group(3)), uom=m.group(4), order=d['delivery'])
            cur['desc'] = clean(L[i + 1]).upper() if i + 1 < len(L) else ''
            d['lines'].append(cur)
            continue
        if cur is None:
            continue
        m = re.match(r'^([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$', ln)
        if m and cur['amount'] is None:
            cur['price'], cur['amount'] = _num(m.group(1)), _num(m.group(2))
            continue
        m = re.match(r'^item net\s+([\d,]+\.\d{2})', ln)
        if m:
            cur['amount'] = _num(m.group(1))
            continue
        m = re.match(r'^Commodity code\s+(\d{6,10})', ln)
        if m:
            cur['hs'] = _hs8(m.group(1))
            continue
        m = re.match(r'^Country of origin\s+([A-Z]{2})\b', ln)
        if m:
            cur['coo'] = m.group(1)
    m = re.search(r'Item sum\s+([\d,]+\.\d{2})', t)
    d['total'] = _num(m.group(1)) if m else None
    return [_finish(d, C)]


def maquet_pl(path, pages, C):
    """Maquet (Suzhou) 'Packing List': order no., Quantity of packages, Net / Gross weight; package groups
    "1-2/3 1,630x755x1,110 373 488 SU000200038 MODUEVO ... 2" + material rows "SU000600073 ... 2" (weights per package
    group only -> net weight of the order split by the invoice value)."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Maquet (Suzhou) packing list')
    m = re.search(r'Order no\.:\s*(\d+)', t)
    d['delivery'] = m.group(1) if m else ''
    m = re.search(r'Quantity of packages:\s*(\d+)', t)
    d['packages'], d['package_unit'] = (int(m.group(1)), 'PKG') if m else (None, '')
    m = re.search(r'Net weight\s+([\d,.]+)\s*KG', t)
    d['total_nw'] = _num(m.group(1)) if m else None
    m = re.search(r'Gross weight\s+([\d,.]+)\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    by = collections.OrderedDict()
    for raw in t.splitlines():
        ln = re.sub(r'^\d+(?:-\d+)?/\d+\s+[\d,]+x[\d,]+x[\d,]+\s+[\d,]+\s+[\d,]+\s+', '', raw.strip())
        m = re.match(r'^([A-Z]{2}\d{9}|[A-Z0-9]{10,12})\s+(.+?)\s+(\d+)$', ln)
        if m:
            b = by.setdefault(m.group(1), dict(desc=clean(m.group(2)).upper(), qty=0.0))
            b['qty'] += float(m.group(3))
    for it, b in by.items():
        d['lines'].append(_line(item=it, desc=b['desc'], qty=b['qty'], uom='EA', order=d['delivery']))
    return d


# ======================================================================== Givaudan Singapore imports (Givaudan International CI, Barneveld packing list, Arnold Suhr)
def givaudan_ci(path, pages, C):
    """Givaudan International 'Customer Commercial Invoice': "52302539 / 10" order line, "1950621 94.000 KG 51.3700 4,828.78"
    (material, net qty, unit, price, amount) + description + "Custom Tariff number: 33011900" + batch rows with origin."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Givaudan customer commercial invoice', 'GIVAUDAN INTERNATIONAL SA')
    d.update(title='COMMERCIAL INVOICE', direction='IMPORT', ship_to_country='SG')
    m = re.search(r'Number\s+Date\s+Page\s*\n\s*(\d{6,})\s+(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _dmy(m.group(2), MONTHS.get(m.group(3).upper(), 1), m.group(4))
    d['invoice_to'] = d['ship_to'] = 'GIVAUDAN SINGAPORE PTE LTD' if 'GIVAUDAN SINGAPORE' in t else ''
    m = re.search(r'Incoterms:\s*([A-Z]{3})', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'^([A-Z]{3})\s+\1\s*$', t, re.M)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'Ship from\s*:\s*\n.*\n.*\n.*\n\s*(?:The\s+)?(\w+)', t)
    d['ship_from_country'] = (C.code(m.group(1)) if m and hasattr(C, 'code') else '') or ''
    L = [x.strip() for x in t.splitlines()]
    cur, order = None, ''
    for i, ln in enumerate(L):
        m = re.match(r'^(\d{6,})\s*/\s*(\d+)\s*$', ln)
        if m:
            order = m.group(1)
            continue
        m = re.match(r'^(\d{5,10})\s+([\d,]+\.\d+)\s+([A-Z]{1,3})\s+([\d,]+\.\d+)\s+([\d,]+\.\d{2})$', ln)
        if m:
            cur = _line(item=m.group(1), qty=_num(m.group(2)), uom=m.group(3), price=_num(m.group(4)), amount=_num(m.group(5)), order=order)
            cur['desc'] = clean(re.sub(r'\s+VAT\s+[\d.]+\s*%.*$', '', L[i + 1])).upper() if i + 1 < len(L) else ''
            if cur['uom'] == 'KG':
                cur['nw'] = cur['qty']
            d['lines'].append(cur)
            continue
        if cur is None:
            continue
        m = re.match(r'^Custom Tariff number:\s*(\d{6,10})', ln)
        if m:
            cur['hs'] = _hs8(m.group(1))
        ml = re.search(r'LA\s*=\s*([\d,.]+)\s*L\b', ln)          # "71.08 LA = 105.927 L" -> litres of the line
        if ml:
            cur['ltr'] = (cur.get('ltr') or 0) + (_num(ml.group(1)) or 0)
        m = re.match(r'^\|\s*(\S+)\s+[\d,.]+\s*[A-Z]{1,3}\s+([A-Za-z ]+?)\s*$', ln)
        if m and not cur['coo']:
            cur['coo'] = C.code(m.group(2)) if hasattr(C, 'code') else ''
            cur['lots'] = [m.group(1)]
    m = re.search(r'Total gross weight:\s*([\d,.]+)\s*KG', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    m = re.search(r'Total Amount\s+[A-Z]{3}\s+([\d,]+\.\d{2})', t)
    d['total'] = _num(m.group(1)) if m else None
    return [_finish(d, C)]


def givaudan_pl(path, pages, C):
    """Givaudan / Barneveld Flavours 'PACKING LIST' (one page per order, spaces often missing): "6825308 25.000 KG 25.000 25.000"
    (material, qty, net, total net) + "H.S.Code:"; "Piece 1 25.000 26.360 0.041" (pieces, net, gross, m3)."""
    import readers as R
    d = R.new_pl(path, 'Givaudan packing list')
    gw = nw = pk = 0.0
    for p in pages:
        t = p['text']
        m = re.search(r'Order\s*Number:\s*G?(\d{6,})', t)
        order = m.group(1) if m else ''
        if order:
            d['invoice_refs'].append(order)
        for m in re.finditer(r'^(\d{5,10})\s+([\d,.]+)\s*KG\s+([\d,.]+)\s+([\d,.]+)\s*\n(.+)', t, re.M):
            l = _line(item=m.group(1), qty=_num(m.group(2)), uom='KG', nw=_num(m.group(4)), order=order, desc=clean(m.group(5)).upper())
            d['lines'].append(l)
        m = re.search(r'^Piece\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)', t, re.M)
        if m:
            pk += int(m.group(1))
            nw += _num(m.group(2)) or 0
            gw += _num(m.group(3)) or 0
            if d['lines'] and d['lines'][-1].get('order') == order:
                d['lines'][-1]['gw'] = _num(m.group(3))
    d['packages'], d['package_unit'] = (int(pk), 'PKG') if pk else (None, '')
    d['total_gw'], d['total_nw'] = round(gw, 3) or None, round(nw, 3) or None
    return d


def arnold_suhr_invoice(path, pages, C):
    """Arnold Suhr Netherlands 'Invoice' to Givaudan Singapore: "1Caffeine Anhydrous 3182463 1,000.00 kg 8.30 8,300.00"
    (line, description, your article no. = Givaudan material, qty, unit, price, total) + "Tariff code: 29393000"."""
    t = '\n'.join(p['text'] for p in pages)
    d = _inv(path, 'Arnold Suhr invoice', 'ARNOLD SUHR NETHERLANDS B.V.')
    d.update(title='INVOICE', ship_from_country='NL', direction='IMPORT', ship_to_country='SG')
    m = re.search(r'Doc\. Number\s+Date\s+Page\s*\n\s*(\d{5,})\s+(\d{2})-(\d{2})-(\d{4})', t)
    if m:
        d['invoice_no'], d['invoice_date'] = m.group(1), _dmy(m.group(2), m.group(3), m.group(4))
    d['invoice_to'] = d['ship_to'] = 'GIVAUDAN SINGAPORE PTE LTD' if 'Givaudan Singapore' in t else ''
    m = re.search(r'Your reference\s*\n\s*(\d+)', t)
    d['customer_po'] = m.group(1) if m else ''
    m = re.search(r'Total\s+([A-Z]{3})\b', t)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'Delivery Terms\s+([A-Z]{3})\b', t)
    d['incoterm'] = m.group(1) if m else ''
    cur = None
    for raw in t.splitlines():
        ln = raw.strip()
        m = re.match(r'^(\d{1,3})\s*([A-Za-z].+?)\s+(\d{5,10})\s+([\d,]+\.\d+)\s+([a-zA-Z]{1,4})\s+([\d,]+\.\d+)\s+([\d,]+\.\d{2})$', ln)
        if m:
            cur = _line(line=m.group(1), item=m.group(3), desc=clean(m.group(2)).upper(), qty=_num(m.group(4)), uom=m.group(5).upper(),
                        price=_num(m.group(6)), amount=_num(m.group(7)), order=d['customer_po'])
            if cur['uom'] == 'KG':
                cur['nw'] = cur['qty']
            d['lines'].append(cur)
            continue
        m = re.match(r'^Tariff code:\s*(\d{6,10})', ln)
        if m and cur is not None:
            cur['hs'] = _hs8(m.group(1))
    m = re.search(r'^Total\s+[A-Z]{3}\s+([\d,]+\.\d{2})', t, re.M)
    d['total'] = _num(m.group(1)) if m else None
    return [_finish(d, C)]


def arnold_suhr_pl(path, pages, C):
    """Arnold Suhr 'Packing List': "1 Caffeine Anhydrous 3182463 20,00 kg carton 1,000.00 kg 1,100.00 kg" + batch rows;
    "Shipment gross weight"."""
    import readers as R
    t = '\n'.join(p['text'] for p in pages)
    d = R.new_pl(path, 'Arnold Suhr packing list')
    m = re.search(r'Doc\. Number\s+Date\s+Page\s*\n\s*(\d{5,})', t)
    d['delivery'] = m.group(1) if m else ''
    for m in re.finditer(r'^(\d{1,3})\s*([A-Za-z].+?)\s+(\d{5,10})\s+[\d,.]+\s*kg\s+(\w+)\s+([\d,]+\.\d+)\s*kg\s+([\d,]+\.\d+)\s*kg', t, re.M):
        d['lines'].append(_line(item=m.group(3), desc=clean(m.group(2)).upper(), qty=_num(m.group(5)), uom='KG', nw=_num(m.group(5)),
                                gw=_num(m.group(6))))
    m = re.search(r'Shipment gross weight:\s*([\d,.]+)\s*kg', t)
    d['total_gw'] = _num(m.group(1)) if m else None
    d['total_nw'] = round(sum(l['nw'] or 0 for l in d['lines']), 3) or None
    return d



# ======================================================================== carrier / forwarder charge documents - not for the permit
def not_for_permit(path, pages, C):
    """Ocean carrier freight / local charges tax invoice or carrier arrival notice (e.g. CMA CGM to the forwarder): no goods
    lines - shown as 'not used' instead of being read as a supplier invoice."""
    t = '\n'.join(p['text'] for p in pages)
    kind = 'carrier arrival notice' if re.search(r'ARRIVAL NOTICE', t[:400]) else 'carrier freight / local charges invoice'
    m = re.search(r'B/L-NO / DEST:\s*(\S+)|Bill of Lading:\s*(\S+)', t)
    ref = (' (ocean B/L %s)' % (m.group(1) or m.group(2))) if m else ''
    return [dict(kind='OTHER', file=os.path.basename(path), path=path, reader='not for the permit',
                 notes=['%s%s - not a goods document, not used for the permit' % (kind, ref)])]



# ======================================================================== Airbus / Satair (Airbus Asia-Pacific Spares Centre) sea shipments
def satair_pack(path, pages, C):
    """One PDF per house B/L (file name = HBL, e.g. 6520079471.PDF): Airbus 'Shipping Advice No/Proforma Invoice' pages
    (European numbers) + Expeditors sea waybill (SEAWAYBILL NUMBER = HBL, packages, gross KGS, vessel, HS).
    Item: "01 0208222269 D2901584400000 0,078 1,000 EA 662,14 662,14" (pos, order, part, kg net, qty, UoM, price, total USD)
    + "261EXU0030 [serial] 81089060900 FR F" (batch, EU tariff, origin) + description [+ date] [+ Dest. Cust Tarrif 88073000]."""
    t = '\n'.join(p['text'] for p in pages)
    fn = os.path.splitext(os.path.basename(path))[0]
    m = re.search(r'SEAWAYBILL NUMBER\s*\n.*?\b(\d{8,12})\s*$', t, re.M) or re.search(r'\b(65\d{8})\b', t)
    hbl = m.group(1) if m else (fn if re.fullmatch(r'[A-Z0-9]{8,14}', fn) else '')
    out = []
    d = _inv(path, 'Airbus shipping advice / proforma (Satair)', 'AIRBUS SAS')
    d.update(title='PROFORMA INVOICE', ship_from_country='DE', direction='IMPORT', ship_to_country='SG', currency='USD')
    m = re.search(r'Proforma Invoice\s+(\d{6,})', t)
    d['invoice_no'] = m.group(1) if m else fn
    m = re.search(r'MATERIAL, LOGISTICS & SUPPLIERS\s+(\d{2})\.(\d{2})\.(\d{4})', t)
    d['invoice_date'] = _dmy(m.group(1), m.group(2), m.group(3)) if m else ''
    d['invoice_to'] = d['ship_to'] = 'AIRBUS ASIA-PACIFIC SPARES CENTRE C/O SATAIR PTE LTD'
    m = re.search(r'Terms of Delivery\s+([A-Z]{3})', t)
    d['incoterm'] = m.group(1) if m else ''
    m = re.search(r'Total\s*(?:USD|([A-Z]{3}))', t)
    d['hawb'] = hbl
    L = [x.strip() for x in t.splitlines()]
    LN = re.compile(r'^(\d{2,3})\s+(\d{8,12})\s+(\S+)\s+([\d.]*\d,\d{3})\s+([\d.]*\d,\d{3})\s+([A-Z]{1,3})\s+([\d.]*\d,\d{2})\s+([\d.]*\d,\d{2})$')
    L2 = re.compile(r'^(\S+)(?:\s+(\S+))?\s+(\d{8,11})\s+([A-Z]{2})\s+[A-Z]$')
    for i, ln in enumerate(L):
        mm = LN.match(ln)
        if not mm:
            continue
        l = _line(line=str(len(d['lines']) + 1), item=mm.group(3), qty=_eu(mm.group(5)), uom=mm.group(6), price=_eu(mm.group(7)),
                  amount=_eu(mm.group(8)), order=mm.group(2), nw=_eu(mm.group(4)), currency='USD')
        l['hawb'] = hbl
        m2 = L2.match(L[i + 1]) if i + 1 < len(L) else None
        j = i + 2 if m2 else i + 1
        if m2:
            l['lots'] = [m2.group(1)]
            l['coo'] = m2.group(4)
            l['eu_hs'] = m2.group(3)[:8]
            if m2.group(2):
                l['serial'] = m2.group(2)
        desc = re.sub(r'\s+\d{2}\.\d{2}\.\d{4}$', '', L[j]) if j < len(L) else ''
        l['desc'] = clean(desc).upper()
        dest = ''
        for x in L[j + 1:j + 6]:
            if re.match(r'^\d{2,3}\s+\d{8,12}\s', x) or x.startswith('For Customs') or x.startswith('Total value'):
                break
            if re.fullmatch(r'\d{8}', x):
                dest = x
            me = re.match(r'ECCN No:\s*US\s+(\S+)', x)
            if me and me.group(1) not in ('NOT_LISTED', 'NOT_SUBJECT'):
                l['eccn'] = me.group(1)
        l['hs'] = dest or l.get('eu_hs', '')
        l['hs_note'] = 'Dest. Cust Tariff' if dest else 'EU tariff'
        d['lines'].append(l)
    m = re.search(r'Total value:\s*([\d.]*\d,\d{3})\s+[\d.,]+QM\s+([\d.]*\d,\d{2})', t)   # page 1 summary: GW, m3, value
    d['total_gw'], d['total'] = (_eu(m.group(1)), _eu(m.group(2))) if m else (None, None)
    m = re.search(r'Number of boxes\s+(\d+)', t)
    d['boxes'] = int(m.group(1)) if m else None
    d['total_nw'] = round(sum(l['nw'] or 0 for l in d['lines']), 3) or None
    d['no_rules'] = True
    d['consolidate_default'] = 'hawb_hs_coo'      # Satair: one item row per house B/L + HS code + origin
    out.append(_finish(d, C))
    # Expeditors sea waybill of this house B/L
    if re.search(r'Sea Waybill|SEA WAYBILL', t):
        a = _awb(path, 'Expeditors sea waybill', 1)
        cg = {'mode': 'Sea'}
        a['hawb'] = hbl
        m = re.search(r'(\d+)\s+(PCS|PKG|PLT|CAS|CTN)S?\s+AIRCRAFT PARTS.*?([\d,]+\.\d+)\s*K\s+([\d.]+)', t)
        if m:
            a['pieces'], a['gw'] = int(m.group(1)), _num(m.group(3))
            cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = a['pieces'], 'PKG', a['gw']
        m = re.search(r'EXPORTCARRIER.*\n\s*(.+?)\s*,\s*(\S+)\s+([A-Z ]+)$', t, re.M)
        if m:
            cg['transport_id'], cg['flight'], cg['loading_name'] = clean(m.group(1)).upper(), m.group(2), clean(m.group(3)).upper()
        a['consignee'] = ''
        a['issued_by'] = 'EXPEDITORS INTERNATIONAL'
        a['cargo'] = cg
        out.append(a)
    return out


def expeditors_an(path, pages, C):
    """Expeditors 'Arrival Notice': master B/L, house B/L list ("HOUSE BL:" block), vessel / voyage, pieces, gross weight,
    containers, ETA, place (PPZ)."""
    t = '\n'.join(p['text'] for p in pages)
    d = _awb(path, 'Expeditors arrival notice', 2)
    cg = {'mode': 'Sea'}
    m = re.search(r'Master Bill of Lading:\s*([A-Z0-9]{8,})', t)
    if m:
        d['mawb'] = cg['mawb'] = m.group(1)
    m = re.search(r'HOUSE BL:\s*\n([\s\S]+?)\n\s*(?:Thank you|$)', t)
    if m:
        d['house_list'] = list(dict.fromkeys(re.findall(r'\b([A-Z0-9]{8,14})\b', m.group(1))))
    m = re.search(r'House Bill of Lading:\s*([A-Z0-9]{8,})', t)
    if m and not d.get('house_list'):
        d['hawb'] = cg['hawb'] = m.group(1)
    m = re.search(r'Vessel:\s*(.+?)/(\S+)', t)
    if m:
        cg['transport_id'], cg['flight'] = clean(m.group(1)).upper(), m.group(2)
    m = re.search(r'Pieces:\s*(\d+)', t)
    if m:
        cg['packages_qty'], cg['packages_unit'] = int(m.group(1)), 'PKG'
    m = re.search(r'Gross Weight:\s*([\d.,]+)\s*K', t)
    if m:
        cg['gw_kg'] = _num(m.group(1))
    conts = re.findall(r'^([A-Z]{4}\d{7})\s+(\S+)\s+\d+\s+(\d{2})([A-Z]?)', t, re.M)
    if conts:
        cg['containers'] = [c[0] for c in conts]
    m = re.search(r'Estimated Date of Arrival:\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', t)
    if m and m.group(2)[:3].upper() in MONTHS:
        cg['eta'] = _dmy(m.group(1), MONTHS[m.group(2)[:3].upper()], m.group(3))
    m = re.search(r'From:\s*([A-Z ]+)\n\s*To:\s*SINGAPORE', t)
    if m:
        cg['loading_name'] = clean(m.group(1)).upper()
    m = re.search(r'Carrier:\s*(.+)', t)
    if m:
        cg['carrier_agent'] = clean(m.group(1)).upper().replace('TLD', 'LTD')
    d['issued_by'] = 'EXPEDITORS SINGAPORE'
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    return [d]


READERS = (
    ('MILLESIMA WINE INVOICE', lambda t: re.search(r'MILLESIMA', t, re.I) and 'DAE Line' in t and re.search(r'Degree', t), millesima_docs),
    ('MAQUET SUZHOU INVOICE', lambda t: re.search(r'Maquet \(Suzhou\)', t) and re.search(r'^Invoice\s*$', t[:200], re.M) and 'Commodity code' in t, maquet_invoice),
    ('MAQUET SUZHOU PACKING LIST', lambda t: re.search(r'Maquet \(Suzhou\)', t, re.I) and 'Packing List' in t[:200] and 'Quantity of packages' in t, maquet_pl),
    ('GIVAUDAN COMMERCIAL INVOICE', lambda t: 'Customer Commercial Invoice' in t[:200] and 'Givaudan International' in t, givaudan_ci),
    ('GIVAUDAN PACKING LIST', lambda t: 'PACKING LIST' in t[:200] and re.search(r'GIVAUDAN\s*SINGAPORE', t) and 'Barneveld' in t, givaudan_pl),
    ('ARNOLD SUHR INVOICE', lambda t: 'Arnold Suhr' in t and re.match(r'\s*Invoice', t) and 'Your article no' in t, arnold_suhr_invoice),
    ('ARNOLD SUHR PACKING LIST', lambda t: 'Arnold Suhr' in t and re.match(r'\s*Packing List', t), arnold_suhr_pl),
    ('CARRIER CHARGES / ARRIVAL NOTICE', lambda t: re.search(r'CMA CGM|Ocean Network Express|MAERSK|Hapag|EVERGREEN|COSCO', t[:1500], re.I) and (
        re.search(r'TAX INVOICE', t[:400]) and re.search(r'OCEAN FREIGHT|Terminal Handling|Bunker', t) or
        re.search(r'^\s*(?:Page 1 of \d+\s*)?ARRIVAL NOTICE', t[:200], re.M) and 'B/L-NO' in t), not_for_permit),
    ('AIRBUS / SATAIR SHIPPING ADVICE', lambda t: 'Shipping Advice No/Proforma Invoice' in t and re.search(r'AIRBUS', t), satair_pack),
    ('EXPEDITORS ARRIVAL NOTICE', lambda t: re.match(r'\s*Arrival Notice', t) and 'EXPEDITORS' in t.upper() and 'Master Bill of Lading' in t, expeditors_an),
    ('UD TRUCKS', lambda t: 'UD TRUCKS SINGAPORE' in t and 'I N V O I C E' in t, ud_trucks),
    ('KALMAR', lambda t: re.search(r'Kalmar\s*Pte\.?\s*Ltd', t, re.I) and re.search(r'TAX\s*INVOICE', t), kalmar),
    ('GETINGE', lambda t: 'GETINGE' in t.upper() and re.search(r'Commercial\s+Invoice', t, re.I), getinge),
    ('CUMMINS', lambda t: 'CUMMINS ASIA PACIFIC' in t and 'HARMONIZED TARIFF' in t, cummins),
    ('L BEAUTY', lambda t: re.search(r'L Beauty Pte Ltd', t) and 'TAX INVOICE' in t, lbeauty),
    ('BMW', lambda t: 'Parts invoice' in t and 'BMW' in t and 'HU ID' in t, bmw),
    ('COMMSCOPE INVOICE', lambda t: re.search(r'Commscope', t, re.I) and 'COMMERCIAL INVOICE' in t[:200] and 'HTS Code' in t, commscope_invoice),
    ('COMMSCOPE PACKAGING SUMMARY', lambda t: 'PACKAGING SUMMARY' in t[:200] and re.search(r'CommScope', t, re.I), commscope_pl),
    ('HALLIBURTON INVOICE', lambda t: 'COMMERCIAL INVOICE' in t[:300] and re.search(r'Halliburton', t) and 'Shipper of Record' in t and 'Harmonized' in t, halliburton_invoice),
    ('HALLIBURTON MANIFEST', lambda t: 'Shipment Manifest / Packing List' in t and re.search(r'Halliburton', t, re.I), halliburton_manifest),
    ('CRANE WORLDWIDE B/L', lambda t: re.search(r'Crane Worldwide Logistics', t) and re.search(r'BILL OF LADING|NOTICE OF ARRIVAL', t[:600]), crane_bl),
    ('NINGBO FENGMEI', lambda t: re.search(r'Ningbo Fengmei Precision', t, re.I) and 'INVOICE' in t, ningbo_fengmei),
    ('NINGBO SAFEWELL', lambda t: 'NINGBO SAFEWELL' in t and 'COMMERCIAL INVOICE' in t, ningbo_safewell),
    ('SHARKNINJA INVOICE', lambda t: re.search(r'SN\s*APAC\s*Holding|SharkNinja', t, re.I) and re.search(r'COMMERCIAL\s+INVOICE', t, re.I)
     and re.search(r'Item\s*Code', t, re.I) and not re.search(r'PACKING\s+LIST', t[:300], re.I), sharkninja_invoice),
    ('SHANGHAI JIANLAN INVOICE', lambda t: 'Shanghai Jianlan Electric' in t and 'COMMERCIAL INVOICE' in t, jianlan_invoice),
    ('ALLISON WUXI INVOICE', lambda t: 'Allison (Wuxi)' in t and re.search(r'Declaration Invoice|Commercial Invoice', t, re.I) and 'Packing List' not in t[:500], allison_invoice),
    ('ALLISON WUXI PACKING LIST', lambda t: 'Allison (Wuxi)' in t and 'Packing List' in t[:500], allison_pl),
    ('SINOTRANS B/L / ARRIVAL NOTICE', lambda t: re.search(r'SINOTRANS', t, re.I) and (re.search(r'ARRIVAL NOTICE', t[:600]) or re.search(r'SHIPPED ON BOARD|BILL OF LADING', t)), sinotrans_bl),
    ('ORGANON INVOICE', lambda t: 'Organon International Services' in t and re.search(r'Tariff Number (?:Exp|Imp)', t), organon_invoice),
    ('ORGANON PACKING LIST', lambda t: 'Packing List' in t[:300] and re.search(r'Organon', t, re.I) and re.search(r"Total Net Weight|Net Weight:|HU'?s\s+[\d.]+\s*KG", t), organon_pl),
    ('MERCK KGAA INVOICE', lambda t: 'Merck Healthcare KGaA' in t and re.search(r'Billing document:', t), merck_kgaa_invoice),
    ('MERCK KGAA PACKING LIST', lambda t: re.match(r'\s*Packing List', t) and 'Merck Healthcare KGaA' in t and 'Freight order' in t, merck_kgaa_pl),
    ('SANDVIK IMPORT INVOICE', lambda t: re.search(r'Sandvik Mining and Construction Logistics', t, re.I) and re.search(r'Import HS Code', t) and re.search(r'^INVOICE', t.lstrip()), sandvik_import_invoice),
    ('MAINFREIGHT HOUSE B/L', lambda t: 'Mainfreight' in t and re.search(r'Express Bill of Lading|BILL OF LADING', t) and 'DOCUMENT NUMBER' in t, mainfreight_hbl),
    ('DUPONT SEM INVOICE PACK', lambda t: 'LEGAL I/C INVOICE' in t and 'SPECIALTY ELECTRONIC MATERIALS' in t, sem_pack),
    ('TAIYO YUDEN PACKING LIST', lambda t: re.search(r'TAIYO YUDEN \(SINGAPORE\)', t) and 'PACKING LIST' in t[:400] and 'Net Weight' in t, taiyo_yuden_pl),
    ('TAIYO YUDEN', lambda t: re.search(r'TAIYO YUDEN \(SINGAPORE\)', t) and 'TAX INVOICE' in t and re.search(r'Per 1000 Pcs|HS No', t), taiyo_yuden),
    ('PERKINS CUSTOMS ONLY INVOICE', lambda t: 'Customs Only Invoice' in t and 'Perkins Engines Company' in t and 'Invoice Detail' in t, perkins_cinv),
    ('PERKINS', lambda t: 'Perkins Engines' in t and 'TARIFF' in t, perkins),
    ('DHL SEA CUSTOMS INSTRUCTIONS', lambda t: re.search(r'Sea Freight (?:FCL|LCL) (?:Arrival|Departure)? ?Customs Instructions (?:Import|Export)', t) and 'HOUSE BILL OF LADING' in t, dhl_ci_sea),
    ('ONE ARRIVAL NOTICE', lambda t: 'ARRIVAL NOTICE' in t[:200] and 'Ocean Network Express' in t, one_an),
    ('ONE SEA WAYBILL', lambda t: re.search(r'SEA WAYBILL|BILL OF LADING', t[:400]) and re.search(r'Ocean Network Express', t) and re.search(r'\bONEY[A-Z0-9]{6,}', t), one_swb),
    ('LEGRAND', lambda t: 'LEGRAND' in t and 'Commodity' in t and 'Country of' in t, legrand),
    ('NURMI', lambda t: 'NURMI HYDRAULICS' in t.upper() and 'INVOICE' in t, nurmi),
    ('DHL CUSTOMS INSTRUCTIONS', lambda t: 'Customs Instructions' in t and 'BROKERAGE' in t and 'HOUSE BILL OF LADING' in t, dhl_ci),
    ('DSV ARRIVAL NOTICE', lambda t: re.search(r'Sea Freight (?:FCL|LCL) Arrival Notice', t) and 'myDSV' in t, dsv_an),
    ('SEA ARRIVAL NOTICE', lambda t: re.search(r'Sea Freight (?:FCL|LCL) Arrival Notice', t[:300] if 'Email Cover Sheet' not in t[:200] else t) and 'OCEAN BILL OF LADING' in t, dsv_an),
    ('DSV BILL OF LADING', lambda t: 'DSV Ocean Transport' in t and re.search(r'B/L No\.|Sea Waybill No\.', t), dsv_bl),
    ('EVEREST', lambda t: 'Shanghai Everest Fine Chemicals' in t and re.search(r'COMMERCIAL\s+INVOICE|PACKING\s+LIST', t), everest),
    ('VF INVOICE', lambda t: 'VF - INVOICE' in t and 'Infor Nexus' in t, vf_invoice),
    ('SANDVIK INVOICE', lambda t: re.search(r'Sandvik (?:Mining|SMC)', t) and 'Invoice Number' in t and 'Export HS Code' in t, sandvik_invoice),
    ('SANDVIK TAX INVOICE', lambda t: svt_is_invoice(t) or svt_is_pl(t) or svt_is_carrier_bol(t), sandvik_tax),
    ('VULCAN INVOICE', lambda t: re.search(r'VULCAN\s*IN\s*DUSTRIAL', t, re.I) and re.search(r'EXPORT\s*INVOICE', t[:300]), vulcan_invoice),
    ('VULCAN PACKING LIST', lambda t: re.search(r'VULCAN\s*IN\s*DUSTRIAL', t, re.I) and re.search(r'PACKING\s*LIST', t[:300]), vulcan_pl),
    ('DOW EXPORT INVOICE', lambda t: 'DOW CHEMICAL PACIFIC' in t and 'Export Invoice' in t and re.search(r'Customs tariff nbr', t), dow_invoice),
    ('DOW PACKING LIST', lambda t: re.search(r'Consolidated\s*Packing\s*List', t) and re.search(r'Dow\s*Chemical\s*Pacific', t, re.I), dow_pl),
    ('DOW BOOKING ADVICE', lambda t: re.search(r'Consolidated\s*Booking\s*Advice', t) and re.search(r'Dow\s*Chemical\s*Pacific', t, re.I), dow_booking),
    ('PANASONIC INVOICE & PACKING LIST', lambda t: re.search(r'INVOICE\s*&\s*PACKING\s*LIST', t) and re.search(r'PANASONIC\s+OPERATIONAL', t), panasonic_inv),
    ('SHELL INVOICE', lambda t: 'Shell Eastern Trading' in t and 'Tax Invoice' in t and 'Net Price' in t, shell_invoice),
    ('SHELL DELIVERY NOTE', lambda t: 'DELIVERY NOTE' in t and 'Shell Eastern Trading' in t and 'Tariff No' in t, shell_dn),
    ('DANMAR SEA WAYBILL', lambda t: 'Danmar Lines' in t and 'B.L. No.' in t and re.search(r'SEA WAYBILL|BILL OF LADING', t), danmar_bl),
    ('SANDVIK PACKING LIST', lambda t: 'Packing List' in t and 'Sandvik Mining and Construction Logistics' in t and 'Generic' in t, sandvik_pl),
)


def _bundled_plugins():
    """Reader plugins shipped next to the program (plugins\\*.py) - used by the desktop tool too. The server reloads
    the ones in PERMIT_SERVER_DATA\\plugins over these (toolrun.load_plugins)."""
    import importlib.util, sys
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    folder = os.path.join(base, 'plugins')
    out = []
    for fn in sorted(os.listdir(folder)) if os.path.isdir(folder) else []:
        if not fn.endswith('.py') or fn.startswith('_'):
            continue
        try:
            name = 'plugin_' + re.sub(r'\W', '_', fn[:-3])
            spec = importlib.util.spec_from_file_location(name, os.path.join(folder, fn))
            mod = importlib.util.module_from_spec(spec)
            sys.modules[name] = mod
            spec.loader.exec_module(mod)
            if hasattr(mod, 'match') and hasattr(mod, 'read'):
                if str(getattr(mod, 'NAME', '')).upper() in {r[0] for r in READERS}:
                    continue                    # built in now
                def wrap(path_, pages, C, _rd=mod.read):
                    return _rd(path_, pages, C)
                wrap._plugin = True
                out.append((str(getattr(mod, 'NAME', fn[:-3])).upper(), mod.match, wrap))
        except Exception:
            continue
    return tuple(out)


try:
    READERS = _bundled_plugins() + READERS
except Exception:
    pass


def detect(text):
    for name, test, fn in READERS:
        try:
            if test(text):
                return name, fn
        except Exception:
            continue
    return None, None


# readers whose documents can come many to one PDF (invoice + packing list, invoice + invoice ...): split by page
SPLIT_READERS = {'COMMSCOPE INVOICE', 'COMMSCOPE PACKAGING SUMMARY', 'MAQUET SUZHOU INVOICE', 'MAQUET SUZHOU PACKING LIST',
                 'ORGANON INVOICE', 'ORGANON PACKING LIST', 'GIVAUDAN COMMERCIAL INVOICE', 'ARNOLD SUHR INVOICE',
                 'ARNOLD SUHR PACKING LIST', 'ALLISON WUXI INVOICE', 'ALLISON WUXI PACKING LIST', 'HALLIBURTON INVOICE',
                 'HALLIBURTON MANIFEST', 'SHARKNINJA INVOICE', 'COMMSCOPE PACKAGING SUMMARY'}
FIRST_PAGE = re.compile(r'Page\s*:?\s*1\s*(?:/|of)\s*\d+', re.I)


def split_pages(pages):
    """[(reader name, fn, pages)] when the PDF holds several documents of the split readers (each starting on its own
    page: a new layout, or 'Page 1 of n' again), else None."""
    if len(pages) < 2:
        return None
    segs = []
    for p in pages:
        name, fn = detect(p['text'])
        start = name in SPLIT_READERS and (not segs or name != segs[-1][0] or FIRST_PAGE.search(p['text']))
        if start:
            segs.append([name, fn, [p]])
        elif segs and (name is None or name == segs[-1][0]):
            segs[-1][2].append(p)
        else:
            return None                    # a page of another kind: not a pack of these documents
    return segs if len(segs) > 1 else None


def read(path, pages, countries=None, log=None):
    """Known exporter layout -> [invoice document]; None when the layout is not one of these. A PDF holding several
    documents (e.g. 29 CommScope invoices + their packaging summaries, 4 Maquet invoices) is read document by document."""
    text = '\n'.join(p['text'] for p in pages)
    segs = split_pages(pages)
    if segs:
        C = countries or Countries()
        docs = []
        for name, fn, pg in segs:
            res = fn(path, pg, C)
            for d in (res if isinstance(res, list) else [res]):
                d['reader'] = d.get('reader') or name
                d['page_range'] = '%d-%d' % (pages.index(pg[0]) + 1, pages.index(pg[-1]) + 1)
                docs.append(d)
        if log:
            log('   %s: %d documents in one PDF (%s), %d item line(s)' % (
                os.path.basename(path), len(docs), ', '.join('%d %s' % (n, k) for k, n in collections.Counter(d.get('kind') for d in docs).items()),
                sum(len(d.get('lines', [])) for d in docs if d.get('kind') == 'INV')))
        return docs
    name, fn = detect(text)
    if not fn:
        return None
    C = countries or Countries()
    res = fn(path, pages, C)
    docs = res if isinstance(res, list) else [res]
    for d in docs:
        d['reader'] = d.get('reader') or name
    if log:
        log('   %s: %s layout, %d document(s), %d item line(s)' % (os.path.basename(path), name, len(docs),
                                                                   sum(len(d.get('lines', [])) for d in docs if d.get('kind') == 'INV')))
    return docs
