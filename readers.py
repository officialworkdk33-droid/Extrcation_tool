# -*- coding: utf-8 -*-
"""INVOICE PERMIT TOOL - document readers for ANY supplier.

Order used for every PDF / Excel file:
  1. built-in readers   : Agilent invoice, air waybill (HAWB / MAWB), MSD invoice / packing list, bill of lading
  2. supplier profile   : a supplier you taught once (tab "Suppliers") - its column names are used
  3. general reader     : finds the item table by its column headings (Part No, Description, HS, Qty, Amount ...)
                          and the header fields by their labels (Invoice No, Gross Weight, HAWB ...)
  4. AI reader (option) : Claude API - only when switched on (always / when the other readers fail)

Every reader returns documents in ONE common shape (see new_inv / new_pl), so the rest of the tool
(HS list, CASC, weights, consolidation, Excel) works the same for every supplier.
"""
import os, re, json, base64, collections, datetime

import permit_engine as E
from permit_engine import clean, squash, fnum, h8, NUM, Countries

# ======================================================================== common document shape
def new_inv(path, reader):
    return dict(kind='INV', file=os.path.basename(path), path=path, reader=reader, notes=[], title='INVOICE',
                invoice_date='', invoice_no='', invoice_type='', order_no='', order_date='', customer_po='',
                incoterm='', payment_terms='', hawb='', mawb='', ship_date='', port_dest='', gts_ref='',
                ship_from='', invoice_from='', invoice_to='', ship_to='', ship_from_block=[], invoice_from_block=[],
                invoice_to_block=[], ship_to_block=[], carrier='', tax_ids=[], uen='', ship_to_country='',
                ship_from_country='', direction='', boxes=None, package_unit='', total_gw=None, total_nw=None,
                currency='', total=None, coo_text='', hs_global='', lines=[], ocr=False)


def new_line(**k):
    d = dict(line='', item='', desc='', hs='', license='', qty=None, uom='', currency='', price=None, amount=None,
             order='', delivery='', eccn='', coo='', pl='', lots=[], boxes=[], nw=None, gw=None)
    d.update(k)
    return d


def new_pl(path, reader):
    return dict(kind='PL', file=os.path.basename(path), path=path, reader=reader, notes=[], invoice_refs=[],
                hawb='', mawb='', total_gw=None, total_nw=None, packages=None, package_unit='', containers=[],
                lines=[], ocr=False, shipper='', consignee='')


# ======================================================================== number / text helpers
CURRENCIES = ('USD', 'EUR', 'SGD', 'GBP', 'JPY', 'CNY', 'RMB', 'CHF', 'AUD', 'HKD', 'MYR', 'THB', 'KRW', 'TWD',
              'INR', 'CAD', 'NZD', 'SEK', 'DKK', 'NOK', 'IDR', 'PHP', 'VND', 'AED', 'SAR')
INCOTERMS_RE = r'\b(EXW|FCA|FAS|FOB|CFR|CNF|C&F|CIF|CPT|CIP|DAP|DPU|DAT|DDP|DDU)\b'
UEN_RE = re.compile(r'\b(?:UEN|CO\.?\s*REG(?:ISTRATION)?\.?\s*NO\.?|COMPANY\s+REG(?:ISTRATION)?\s+NO\.?)\s*[:.#]?\s*'
                    r'(\d{8,9}[A-Z]|[STR]\d{2}[A-Z]{2}\d{4}[A-Z])\b', re.I)


NUM_STYLE = {'style': 'auto'}       # 'us' (1,234.56) / 'eu' (1.234,56) / 'auto' - set per document


def detect_style(text):
    us = len(re.findall(r'\d,\d{3}\.\d{1,2}\b|\b\d+\.\d{2}\b', text))
    eu = len(re.findall(r'\d\.\d{3},\d{1,2}\b|\b\d+,\d{2}\b', text))
    NUM_STYLE['style'] = 'eu' if eu > us * 1.5 and eu >= 2 else 'us' if us else 'auto'
    return NUM_STYLE['style']


def num(s):
    """'1,234.56' / '1.234,56' / '1 234,5' / '(12.00)' -> float, else None."""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = str(s).strip().replace(' ', ' ')
    t = re.sub(r'^[A-Z]{3}\s*|\s*[A-Z]{3}$', '', t)            # currency code around the number
    t = t.replace('$', '').replace('€', '').strip()
    neg = t.startswith('(') and t.endswith(')')
    t = t.strip('()').replace(' ', '')
    st = NUM_STYLE['style']
    if st == 'eu' or (st == 'auto' and (re.fullmatch(r'-?\d{1,3}(\.\d{3})+(,\d+)?', t) or
                                         (re.fullmatch(r'-?\d+,\d{1,3}', t) and not re.fullmatch(r'-?\d{1,3},\d{3}', t)))):
        t = t.replace('.', '').replace(',', '.')                # European 1.234,56
    else:
        t = t.replace(',', '')
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


NUM_TOKEN = re.compile(r'^\(?-?[\d.,\s]*\d\)?$')


def is_num(s):
    return bool(s) and bool(NUM_TOKEN.match(str(s).strip())) and num(s) is not None


def label_value(text, labels, pattern=r'((?=[A-Z0-9\-/.]*\d)[A-Z0-9][A-Z0-9\-/.]{2,})', flags=re.I):
    """First 'LABEL : value' (same line) for any label regex."""
    for lab in labels:
        m = re.search(r'(?:%s)\s*[:.#]?\s*%s' % (lab, pattern), text, flags)
        if m:
            return m.group(1).strip()
    return ''


# ======================================================================== field synonyms (column headings)
FIELDS = collections.OrderedDict([
    ('line', ['line', 'line no', 'ln', 'no', 's/n', 'sn', 'sr', 'sr no', 'sl no', 'pos', 'position', 'item no.', '#']),
    ('item', ['part no', 'part number', 'part #', 'p/n', 'pn', 'material', 'material no', 'material number',
              'item code', 'item no', 'item number', 'product code', 'product no', 'product number', 'article',
              'article no', 'art no', 'sku', 'catalog no', 'cat no', 'catalogue no', 'model', 'model no', 'part',
              'reference', 'ref', 'customer part no', 'mfr part no', 'item']),
    ('desc', ['description', 'product description', 'item description', 'description of goods', 'goods description',
              'commodity description', 'particulars', 'designation', 'material description', 'goods', 'commodity']),
    ('hs', ['hs', 'hs code', 'hs no', 'hts', 'hts code', 'htsus', 'hs/hts', 'tariff', 'tariff code', 'tariff no',
            'commodity code', 'customs code', 'customs tariff', 'harmonized code', 'hs tariff']),
    ('coo', ['coo', 'country of origin', 'origin', 'c/o', 'made in', 'orig', 'country', 'ctry of origin', 'org']),
    ('qty', ['qty', 'quantity', 'qty shipped', 'quantity shipped', 'ship qty', 'shipped qty', 'pcs', 'units',
             'qty.', 'delivered qty', 'invoiced qty', 'quantity (pcs)', 'no. of units']),
    ('uom', ['uom', 'unit', 'um', 'u/m', 'unit of measure', 'qty unit', 'unit of measurement']),
    ('price', ['unit price', 'price', 'unit value', 'rate', 'price/unit', 'unit cost', 'price per unit', 'u/price',
               'net price', 'unit price usd', 'unit price eur']),
    ('amount', ['amount', 'total', 'total price', 'extended amount', 'ext price', 'extended price', 'line total',
                'value', 'total value', 'net value', 'extension', 'total amount', 'net amount', 'line amount',
                'invoice value', 'total usd', 'amount usd', 'total eur', 'amount eur']),
    ('nw', ['net weight', 'net wt', 'nw', 'n.w.', 'n/w', 'net kg', 'net wt (kg)', 'net weight (kg)', 'nett weight']),
    ('gw', ['gross weight', 'gross wt', 'gw', 'g.w.', 'g/w', 'gross kg', 'gross weight (kg)']),
    ('packages', ['cartons', 'ctns', 'ctn', 'carton', 'ctn.', 'packages', 'pkgs', 'no of pkgs', 'boxes', 'carton no',
                  'ctn no', 'box no', 'case no', 'case', 'pallets', 'plts', 'pallet no', 'no. of packages', 'package']),
    ('lot', ['lot', 'batch', 'lot no', 'batch no', 'lot number', 'batch number', 'serial', 'serial no']),
    ('eccn', ['eccn', 'us eccn', 'export control']),
    ('currency', ['currency', 'curr', 'cur', 'ccy']),
])
FIELD_LABELS = {'line': 'Line No', 'item': 'Item / Part No', 'desc': 'Description', 'hs': 'HS code', 'coo': 'COO',
                'qty': 'Quantity', 'uom': 'UOM', 'price': 'Unit price', 'amount': 'Amount', 'nw': 'Net weight',
                'gw': 'Gross weight', 'packages': 'Packages', 'lot': 'Lot / batch', 'eccn': 'ECCN',
                'currency': 'Currency', '-': '(ignore)'}


def _norm(s):
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9#/%. ]', ' ', str(s).lower())).strip(' .')


def match_field(head, profile_cols=None):
    """Column heading text -> field key ('' if unknown)."""
    h = _norm(head)
    if not h:
        return ''
    if profile_cols:
        for ph, f in profile_cols.items():
            if _norm(ph) == h:
                return f
    best, blen = '', 0
    for f, syns in FIELDS.items():
        for s in syns:
            if h == s or (len(s) > 3 and (h.startswith(s + ' ') or h.endswith(' ' + s) or (' ' + s + ' ') in (' ' + h + ' '))):
                if len(s) > blen:
                    best, blen = f, len(s)
    return best


# ======================================================================== supplier profiles
class Profiles:
    def __init__(self, path):
        self.path = path
        self.items = []
        try:
            with open(path, encoding='utf-8') as f:
                self.items = json.load(f)
        except Exception:
            self.items = []

    def save(self):
        with open(self.path, 'w', encoding='utf-8') as f:
            json.dump(self.items, f, indent=1)

    def find(self, text, kind=None):
        U = text.upper()
        best = None
        for p in self.items:
            if kind and p.get('kind') not in (None, '', kind):
                continue
            keys = [k for k in p.get('match', []) if k.strip()]
            if keys and all(k.upper() in U for k in keys):
                if not best or len(''.join(keys)) > len(''.join(best.get('match', []))):
                    best = p
        return best

    def upsert(self, prof):
        self.items = [p for p in self.items if p['name'].upper() != prof['name'].upper() or p.get('kind') != prof.get('kind')]
        self.items.append(prof)
        self.save()

    def delete(self, name, kind=None):
        self.items = [p for p in self.items if not (p['name'] == name and (kind is None or p.get('kind') == kind))]
        self.save()


# ======================================================================== page words -> lines
def word_lines(words, tol=3.0):
    """[[word,...], ...] grouped by line (top), each sorted by x."""
    rows = []
    for w in sorted(words, key=lambda w: (round(w['top'], 1), w['x0'])):
        if rows and abs(rows[-1][0] - w['top']) <= tol:
            rows[-1][1].append(w)
        else:
            rows.append([w['top'], [w]])
    return [sorted(r[1], key=lambda w: w['x0']) for r in rows]


def find_header(lines, profile_cols=None, min_fields=3):
    """Best heading line (or two lines joined) -> (index, [(field, x0, x1, text)])."""
    best = None
    for i, ln in enumerate(lines):
        for span in (1, 2):
            if i + span > len(lines):
                continue
            ws = [w for l in lines[i:i + span] for w in l]
            if span == 2 and abs(lines[i + 1][0]['top'] - lines[i][0]['top']) > 14:
                continue
            cols = []
            for w in ws:
                f = match_field(w['text'], profile_cols)
                if f and not is_num(w['text']):
                    cols.append([f, w['x0'], w['x1'], w['text']])
            # merge 2-line headings in the same column ("Unit" / "Price")
            if span == 2:
                cols.sort(key=lambda c: c[1])
            fields = {c[0] for c in cols} - {'-'}
            score = len(fields) + (2 if 'amount' in fields else 0) + (2 if 'qty' in fields else 0) + \
                (1 if 'desc' in fields else 0)
            if len(fields) >= min_fields and ('qty' in fields or 'amount' in fields or 'nw' in fields) and \
                    ('desc' in fields or 'item' in fields) and \
                    (not best or score > best[0]):
                # one column per field: keep the left-most occurrence
                seen, uniq = set(), []
                for c in sorted(cols, key=lambda c: c[1]):
                    if c[0] in seen and c[0] != '-':
                        continue
                    seen.add(c[0])
                    uniq.append(c)
                best = (score, i + span - 1, uniq)
    return (best[1], best[2]) if best else (None, None)


def column_bounds(cols, page_w=900):
    """Heading positions -> [(field, left, right)] using the midpoints between headings."""
    cs = sorted(cols, key=lambda c: c[1])
    out = []
    for j, c in enumerate(cs):
        left = (cs[j - 1][2] + c[1]) / 2.0 if j else -1
        right = (c[2] + cs[j + 1][1]) / 2.0 if j + 1 < len(cs) else page_w + 1
        out.append((c[0], left, right))
    return out


STOP_RE = re.compile(r'^\s*(sub\s*-?\s*total|total\b|grand total|invoice total|total amount|amount due|'
                     r'balance due|total value|total invoice|say\b|amount in words|special instructions|'
                     r'remarks\s*:|bank details|terms and conditions)', re.I)
COO_IN_TEXT = re.compile(r'(?:country\s+of\s+origin|c\.?o\.?o\.?|origin|made\s+in)\s*[:\-]?\s*([A-Za-z][A-Za-z .]{1,30})', re.I)
HS_IN_TEXT = re.compile(r'(?:HS|HTS|H\.S\.|tariff|commodity)\s*(?:code|no\.?|number|#)?\s*[:.]?\s*(\d{4}[.\s]?\d{2}[.\s]?\d{2}(?:[.\s]?\d{2})?)', re.I)


def table_lines(pages, kind='INV', profile=None, countries=None):
    """General line-item reader: heading columns by position, rows below them."""
    C = countries or Countries()
    pcols = (profile or {}).get('columns') or None
    out, cols, notes = [], None, []
    for pg in pages:
        lines = word_lines(pg['words'], tol=5.0 if pg.get('ocr') else 3.0)
        hi, hcols = find_header(lines, pcols)
        if hi is not None:
            cols = column_bounds(hcols)
            start = hi + 1
        elif cols is not None:
            start = 0
        else:
            continue
        cur = None
        for ln in lines[start:]:
            text = ' '.join(w['text'] for w in ln)
            if STOP_RE.search(text):
                cur = None
                if kind == 'INV':
                    break
                continue
            cells = collections.defaultdict(list)
            for w in ln:
                xc = (w['x0'] + w['x1']) / 2.0
                f = next((c[0] for c in cols if c[1] <= xc < c[2]), None)
                if f is None:                                    # outside -> nearest
                    f = min(cols, key=lambda c: min(abs(c[1] - xc), abs(c[2] - xc)))[0]
                cells[f].append(w['text'])
            cell = {f: clean(' '.join(v)) for f, v in cells.items()}
            if not any(c[0] == 'item' for c in cols):        # unnamed part-no column next to the line no.
                for f in ('line', 'desc'):
                    m = re.match(r'^(\d{1,4})\s+([A-Z0-9][A-Z0-9\-./]{2,})(?:\s+(.*))?$', cell.get(f, '')) if f == 'line' else \
                        re.match(r'^([A-Z0-9]*\d[A-Z0-9]*(?:[-./][A-Z0-9]+)+)\s+(.+)$', cell.get(f, ''))
                    if m and f == 'line':
                        cell['line'], cell['item'] = m.group(1), m.group(2)
                        if m.group(3):
                            cell['desc'] = clean(m.group(3) + ' ' + cell.get('desc', ''))
                        break
            amt = num(cell.get('amount')) if is_num(cell.get('amount', '').split(' ')[-1] if cell.get('amount') else '') else None
            if amt is None and cell.get('amount'):
                m = re.search(r'(-?[\d.,]*\d)\s*$', cell['amount'])
                amt = num(m.group(1)) if m else None
            qty_s = cell.get('qty', '')
            mq = re.match(r'^\s*(\(?-?[\d.,]*\d\)?)\s*([A-Za-z]{1,5})?\s*$', qty_s)
            qty = num(mq.group(1)) if mq else None
            nwv = num(cell.get('nw')) if is_num(cell.get('nw', '')) else None
            is_row = (amt is not None and (qty is not None or cell.get('price'))) if kind == 'INV' else \
                (qty is not None or nwv is not None) and bool(cell.get('item') or cell.get('desc'))
            if is_row:
                cur = new_line(line=cell.get('line', ''), item=cell.get('item', ''), desc=cell.get('desc', ''),
                               qty=qty, uom=(cell.get('uom') or (mq.group(2) if mq and mq.group(2) else '')).upper(),
                               price=num(re.sub(r'[^\d.,()-]', '', cell.get('price', '')) or None),
                               amount=amt, currency='')
                hs = re.sub(r'\D', '', cell.get('hs', ''))
                if len(hs) >= 6:
                    cur['hs'] = hs[:8] if len(hs) >= 8 else hs
                else:                                           # an 8-10 digit code printed in another column
                    for f, v in cell.items():
                        if f in ('qty', 'amount', 'price', 'item', 'line'):
                            continue
                        m = re.search(r'(?<![\d.,])(\d{4}\.?\d{2}\.?\d{2}(?:\.?\d{2})?)(?![\d.,])', v)
                        if m and len(re.sub(r'\D', '', m.group(1))) in (8, 10):
                            cur['hs'] = re.sub(r'\D', '', m.group(1))[:8]
                            cell[f] = clean(v.replace(m.group(1), ' '))
                            if f == 'desc':
                                cur['desc'] = cell[f]
                            break
                if cell.get('coo'):
                    c = C.code(cell['coo'])
                    cur['coo'] = c or (cell['coo'].upper() if len(cell['coo']) == 2 else '')
                for f in ('nw', 'gw'):
                    if cell.get(f) and is_num(cell[f].split(' ')[0]):
                        cur[f] = num(cell[f].split(' ')[0])
                if cell.get('packages') and is_num(cell['packages']):
                    cur['packages'] = num(cell['packages'])
                if cell.get('lot'):
                    cur['lots'] = [cell['lot']]
                if cell.get('eccn'):
                    cur['eccn'] = cell['eccn']
                if cell.get('currency') and cell['currency'].upper() in CURRENCIES:
                    cur['currency'] = cell['currency'].upper()
                for ccy in CURRENCIES:                           # "USD 1,234.00" in a cell
                    if re.search(r'\b%s\b' % ccy, cell.get('amount', '') + ' ' + cell.get('price', '')):
                        cur['currency'] = ccy
                        break
                if not cur['item'] and cur['desc']:
                    m = re.match(r'^([A-Z0-9][A-Z0-9\-./]{3,})\s+(.+)$', cur['desc'])
                    if m and re.search(r'\d', m.group(1)):
                        cur['item'], cur['desc'] = m.group(1), m.group(2)
                out.append(cur)
            elif cur is not None:
                # continuation line: more description, HS / COO / lot written under the item
                m = HS_IN_TEXT.search(text)
                if m and not cur['hs']:
                    cur['hs'] = re.sub(r'\D', '', m.group(1))[:8]
                    text = text.replace(m.group(0), ' ')
                m = COO_IN_TEXT.search(text)
                if m and not cur['coo']:
                    cur['coo'] = C.code(m.group(1)) or ''
                    if cur['coo']:
                        text = text.replace(m.group(0), ' ')
                m = re.search(r'\b(?:lot|batch)\s*(?:no\.?)?\s*[:#]?\s*([A-Z0-9][A-Z0-9\-/]{2,})', text, re.I)
                if m:
                    cur['lots'].append(m.group(1))
                    text = text.replace(m.group(0), ' ')
                extra = clean(cell.get('desc', '') if cell.get('desc') else '')
                for rx in (HS_IN_TEXT, COO_IN_TEXT, re.compile(r'\b(?:lot|batch)\s*(?:no\.?)?\s*[:#]?\s*[A-Z0-9][A-Z0-9\-/]{2,}', re.I)):
                    extra = clean(rx.sub(' ', extra))
                if extra and len(extra) < 120 and not re.match(r'^(page|continued|carried)', extra, re.I):
                    cur['desc'] = clean(cur['desc'] + ' ' + extra)
                elif not cell.get('desc') and clean(text) and len(ln) <= 6 and not is_num(text) \
                        and not re.search(r'page \d', text, re.I) and len(clean(text)) < 80:
                    pass
    if not out and cols is None:
        notes.append('no item table heading found (Qty / Amount / Description ...)')
    return out, notes


# ======================================================================== header fields (any layout)
PARTY_LABELS = {
    'seller': r'(shipper|exporter|seller|consignor|sold\s+by|vendor|supplier|from|manufacturer|invoice\s+from)',
    'buyer': r'(consignee|importer|buyer|sold\s+to|ship\s+to|bill\s+to|invoice\s+to|deliver\s+to|delivery\s+address|to)',
}


def party_block(words, label_re, countries):
    """Lines under a party label ('Consignee:' ...) -> (name, block lines, country)."""
    for w in words:
        if re.fullmatch(r'%s\s*(name\s+and\s+address|address|:|\(.*\))?\s*:?' % label_re, w['text'].strip(), re.I):
            sel = [x for x in words if w['x0'] - 8 <= x['x0'] < w['x0'] + 240 and w['top'] + 3 < x['top'] < w['top'] + 75]
            blk = [l for l in E._lines(sel)][:6]
            blk = [l for l in blk if not re.match(r'(tel|fax|phone|e-?mail|attn|contact)\b', l, re.I)]
            if blk:
                name = E._company(blk) or blk[0]
                return name, blk, countries.find_in(' '.join(blk), exclude=())
    return '', [], ''


PKG_UNITS = r'(packages|pkgs?|cartons?|ctns?|boxes|box|pallets?|plts?|cases?|crates?|skids?|drums?|bags?)'


def weight_value(text, label):
    """'Gross Weight: 1,234.5 KGS' -> 1234.5 (LB converted). Lines with 'total' first; never across lines."""
    pat = re.compile(r'(?:%s)[ \t]*(?:in[ \t]*)?(?:\(?[ \t]*(?:kgs?|kgm|lbs?)[ \t]*\)?)?[ \t]*[:.=]?[ \t]*(\d[\d.,]*)[ \t]*(KGS?|KGM|LBS?)?\b' % label, re.I)
    lines = text.splitlines()
    for group in ([l for l in lines if re.search(r'total', l, re.I)], lines):
        for ln in group:
            m = pat.search(ln)
            if m and num(m.group(1)):
                v = num(m.group(1))
                return round(v * 0.45359237, 3) if (m.group(2) or '').upper().startswith('LB') else v
    return None


def package_count(text):
    lines = text.splitlines()
    pats = [re.compile(r'\b(\d{1,5})[ \t]*%s\b' % PKG_UNITS, re.I),
            re.compile(r'(?:total|no\.?\s+of|number\s+of)[ \t]*%s[ \t]*[:.]?[ \t]*(\d{1,5})\b' % PKG_UNITS, re.I),
            re.compile(r'\b%s[ \t]*[:][ \t]*(\d{1,5})\b' % PKG_UNITS, re.I)]
    for group in ([l for l in lines if re.search(r'total|packages\s*:|no\.?\s+of', l, re.I)], lines):
        for p in pats:
            for ln in group:
                m = p.search(ln)
                if m:
                    g = m.groups()
                    n, u = (g[0], g[1]) if g[0].isdigit() else (g[1], g[0])
                    return int(n), u.upper()
    return None, ''


COMPANY_RE = re.compile(r'\b(INC|LTD|LIMITED|GMBH|CO\.?|CORP(?:ORATION)?|LLC|PTE|S\.?A\.?|B\.?V\.?|AG|S\.?P\.?A|SDN\s+BHD|'
                        r'PLC|KK|K\.K\.|SARL|SAS|AB|OY|NV|PTY)\b\.?', re.I)
LABEL_SPLIT = re.compile(r'\s{2,}|\b(?:ship|bill|sold|deliver|invoice)\s+to\b|\bconsignee\b|\bnotify\b|\btel\b|\bphone\b', re.I)


def header_fields(d, text, words, countries):
    C = countries
    first = text[:4000]
    if not d['invoice_no']:
        d['invoice_no'] = label_value(text, [r'invoice\s*(?:no|number|nr|#)\.?', r'inv\.?\s*(?:no|#)\.?',
                                             r'commercial\s+invoice\s*(?:no|#)\.?', r'proforma\s*(?:invoice)?\s*(?:no|#)\.?',
                                             r'document\s*(?:no|number)\.?', r'bill\s*(?:no|#)\.?'])
        if not d['invoice_no'] and words:
            for lab in ('Invoice No.', 'Invoice No', 'Invoice Number', 'Invoice #', 'Inv No'):
                v = E._below(words, lab, pattern=r'(?=[A-Z0-9\-/]*\d)[A-Z0-9][A-Z0-9\-/]{3,}')
                if v:
                    d['invoice_no'] = v
                    break
    if not d['invoice_date']:
        m = re.search(r'(?:invoice\s+)?date\s*(?:of\s+invoice)?\s*[:.]?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{4}-\d{2}-\d{2}|'
                      r'\d{1,2}[ -][A-Za-z]{3,9}[ -,]+\d{2,4}|[A-Za-z]{3,9}\.? \d{1,2},? \d{4})', first, re.I)
        d['invoice_date'] = m.group(1) if m else ''
    if not d['currency']:
        cnt = collections.Counter(re.findall(r'\b(%s)\b' % '|'.join(CURRENCIES), text))
        d['currency'] = cnt.most_common(1)[0][0] if cnt else ''
        if d['currency'] == 'RMB':
            d['currency'] = 'CNY'
    if not d['incoterm']:
        m = re.search(r'(?:incoterms?|terms\s+of\s+delivery|delivery\s+terms|trade\s+terms|price\s+terms|terms)'
                      r'\s*(?:\d{4})?\s*[:.]?\s*' + INCOTERMS_RE, text, re.I) or re.search(INCOTERMS_RE + r'\s+[A-Z][a-z]+', first)
        d['incoterm'] = m.group(1).upper().replace('C&F', 'CFR').replace('CNF', 'CFR') if m else ''
    if d['total'] is None:
        m = re.search(r'(?:grand\s+total|invoice\s+total|total\s+amount|total\s+value|total\s+invoice\s+value|amount\s+due|'
                      r'total\s+invoice)\s*(?:\((?:%s)\))?\s*[:.=]?\s*(?:%s)?\s*(\(?-?\d[\d.,]*\d\)?)' % ('|'.join(CURRENCIES), '|'.join(CURRENCIES)),
                      text, re.I)
        if m and num(m.group(1)) is not None:
            d['total'] = num(m.group(1))
    if d['total'] is None:
        for ln in reversed(text.splitlines()):
            if re.search(r'(grand\s+total|invoice\s+total|total\s+amount|total\s+value|total\s+invoice|amount\s+due|'
                         r'total\s+\(?(?:%s)\)?|^\s*total\b)' % '|'.join(CURRENCIES), ln, re.I) and \
                    not re.search(r'weight|qty|quantity|packages|cartons|boxes|pcs|kgs?\b', ln, re.I):
                ns = [x for x in re.findall(r'\(?-?\d[\d.,]*\d\)?|\d', ln) if num(x) is not None]
                if ns:
                    d['total'] = num(ns[-1])
                    break
    if d['total_gw'] is None:
        d['total_gw'] = weight_value(text, r'gross\s*w(?:eigh)?t\.?|\bG\.\s?W\.|\bGW\b|total\s+gross')
    if d['total_nw'] is None:
        d['total_nw'] = weight_value(text, r'net{1,2}\s*w(?:eigh)?t\.?|\bN\.\s?W\.|\bNW\b|total\s+net')
    if d['boxes'] is None:
        d['boxes'], d['package_unit'] = package_count(text)
    if not d['hawb']:
        d['hawb'] = label_value(text, [r'h\.?\s?a\.?\s?w\.?\s?b\.?', r'house\s*(?:awb|air\s*waybill|b/?l)', r'hbl',
                                       r'h\.?b/l'], r'(?:no\.?|number|#)?\s*[:.]?\s*([A-Z0-9][A-Z0-9\-]{5,})')
    if not d['mawb']:
        m = re.search(r'(?:m\.?a\.?w\.?b\.?|master\s*awb|air\s*waybill|awb)\s*(?:no\.?|number|#)?\s*[:.]?\s*(\d{3})\s*-?\s*(\d{4})\s*(\d{4})', text, re.I)
        d['mawb'] = ''.join(m.groups()) if m else ''
    if not d['coo_text']:
        m = re.search(r'country\s+of\s+origin\s*(?:of\s+goods)?\s*[:.]?\s*([A-Za-z][A-Za-z .]{2,30})(?:\n|$|\s{2,})', text, re.I) or \
            re.search(r'\bmade\s+in\s+([A-Za-z][A-Za-z .]{2,30})(?:\n|$|\s{2,})', text, re.I)
        d['coo_text'] = clean(m.group(1)) if m else ''
    if not d['hs_global']:
        m = HS_IN_TEXT.search(text)
        d['hs_global'] = re.sub(r'\D', '', m.group(1))[:8] if m else ''
    m = UEN_RE.search(text)
    d['uen'] = d['uen'] or (m.group(1).upper() if m else '')
    # parties
    if not d['invoice_from'] or not d['ship_to']:
        sname, sblk, scty = party_block(words, PARTY_LABELS['seller'], C) if words else ('', [], '')
        bname, bblk, bcty = party_block(words, PARTY_LABELS['buyer'], C) if words else ('', [], '')
        if not sname:
            m = re.search(r'(?:shipper|exporter|seller|consignor|sold\s+by|vendor)\s*[:]\s*(.+)', text, re.I)
            sname = clean(LABEL_SPLIT.split(m.group(1))[0]).strip(' ,') if m else ''
        if not sname:                                   # letterhead: first company-like line of the page
            for ln in text.splitlines()[:6]:
                if COMPANY_RE.search(ln) and not re.search(r'invoice|packing|to\s*:', ln, re.I):
                    sname = clean(ln)
                    break
        if not bname:
            m = re.search(r'(?:consignee|importer|buyer|sold\s+to|ship\s+to|bill\s+to|deliver\s+to)\s*[:]\s*(.+)', text, re.I)
            if m:
                full = m.group(1)
                bname = clean(re.split(r',|\s{2,}', LABEL_SPLIT.split(full)[0])[0]).strip(' ,')
                if not bcty:
                    bcty = C.find_in(full, exclude=())
        d['invoice_from'] = d['invoice_from'] or sname
        d['ship_from'] = d['ship_from'] or sname
        d['ship_to'] = d['ship_to'] or bname
        d['invoice_to'] = d['invoice_to'] or bname
        d['ship_from_block'] = d['ship_from_block'] or sblk
        d['ship_to_block'] = d['ship_to_block'] or bblk
        d['ship_from_country'] = d['ship_from_country'] or scty
        d['ship_to_country'] = d['ship_to_country'] or bcty
    if not d['direction']:
        if d['ship_to_country'] == 'SG' or (d['ship_to_country'] == '' and 'SINGAPORE' in ' '.join(d['ship_to_block']).upper()):
            d['direction'] = 'IMPORT'
        elif d['ship_from_country'] == 'SG' or 'SINGAPORE' in ' '.join(d['ship_from_block']).upper():
            d['direction'] = 'EXPORT'
        else:
            d['direction'] = 'IMPORT'
    if not d['carrier']:
        m = re.search(r'\b(KINTETSU WORLD\s*EXPRESS|DHL[A-Z ]{0,20}|FEDEX|UPS|EXPEDITORS|DB SCHENKER|SCHENKER|'
                      r'KUEHNE\s*\+?\s*NAGEL|NIPPON EXPRESS|YUSEN|HELLMANN|CEVA|DSV|AGILITY|BOLLORE|GEODIS|YAMATO)\b', text, re.I)
        d['carrier'] = clean(m.group(1)).upper() if m else ''
    return d


def finish_lines(d, countries):
    """Apply document-wide COO / HS / currency to lines that have none; sanity notes."""
    C = countries
    gcoo = C.code(d['coo_text']) if d.get('coo_text') else ''
    for i, l in enumerate(d['lines'], 1):
        if not l['line']:
            l['line'] = str(i)
        if not l['coo'] and gcoo:
            l['coo'] = gcoo
        if not l['currency']:
            l['currency'] = d['currency']
        if l['qty'] is None and l['price'] and l['amount']:
            l['qty'] = round(l['amount'] / l['price'], 4)
        if not l['uom']:
            l['uom'] = 'EA' if l['qty'] is not None and float(l['qty']).is_integer() else ''
    if not d['currency']:
        d['currency'] = next((l['currency'] for l in d['lines'] if l['currency']), '')
    if gcoo:
        d['notes_coo'] = 'COO %s taken from "Country of Origin" on the document for lines without their own COO' % gcoo
    return d


# ======================================================================== general reader (text PDF / OCR)
def generic_pdf(path, pages, kind, countries, profile=None):
    text = '\n'.join(p['text'] for p in pages)
    detect_style(text)
    words1 = (pages[0].get('hwords') or pages[0]['words']) if pages else []
    if kind == 'PL':
        d = new_pl(path, 'profile: ' + profile['name'] if profile else 'general')
        rows, notes = table_lines(pages, 'PL', profile, countries)
        d['lines'] = rows
        d['notes'] += notes
        h = header_fields(new_inv(path, ''), text, words1, countries)
        d.update(hawb=h['hawb'], mawb=h['mawb'], total_gw=h['total_gw'], total_nw=h['total_nw'], packages=h['boxes'],
                 package_unit=h['package_unit'], shipper=h['invoice_from'], consignee=h['ship_to'])
        refs = re.findall(r'(?:invoice|inv\.?)\s*(?:no|number|#)\.?\s*[:.]?\s*([A-Z0-9][A-Z0-9\-/]{3,})', text, re.I)
        d['invoice_refs'] = list(dict.fromkeys(refs))
        d['containers'] = list(dict.fromkeys(re.findall(r'\b([A-Z]{4}\d{7})\b', text)))
        d['ocr'] = any(p.get('ocr') for p in pages)
        return d
    d = new_inv(path, 'profile: ' + profile['name'] if profile else 'general')
    d['title'] = 'PROFORMA INVOICE' if re.search(r'PRO\s*-?FORMA', text[:1500], re.I) else \
        'TAX INVOICE' if re.search(r'TAX\s+INVOICE', text[:1500], re.I) else 'COMMERCIAL INVOICE'
    rows, notes = table_lines(pages, 'INV', profile, countries)
    d['lines'] = rows
    d['notes'] += notes
    header_fields(d, text, words1, countries)
    finish_lines(d, countries)
    d['ocr'] = any(p.get('ocr') for p in pages)
    if d['ocr']:
        d['notes_ocr'] = 'scanned invoice read by OCR - check item no / HS / qty / amounts'
    return d


def quality(d):
    """0..1 - how well a reader understood the invoice (used to decide on the AI fallback)."""
    if d.get('kind') == 'PL':
        return 1.0 if d['lines'] or d['total_gw'] or d['total_nw'] else 0.0
    if d.get('kind') != 'INV' or not d.get('lines'):
        return 0.0
    s = sum(l['amount'] or 0 for l in d['lines'])
    q = 0.6
    if d.get('total'):
        q = 1.0 if abs(s - d['total']) <= max(0.05, d['total'] * 0.001) else 0.3
    if not any(l['desc'] for l in d['lines']):
        q -= 0.2
    return q


# ======================================================================== Excel / CSV invoices
def generic_excel(path, countries, profile=None, kind=None):
    rows = E.read_rows(path) if not path.lower().endswith('.csv') else _csv_rows(path)
    pcols = (profile or {}).get('columns') or None
    hi, cmap = None, {}
    for i, r in enumerate(rows[:60]):
        m = {}
        for j, v in enumerate(r):
            if v is None:
                continue
            f = match_field(v, pcols)
            if f and f not in m.values() and not is_num(v):
                m[j] = f
        fs = set(m.values())
        if len(fs) >= 3 and ('qty' in fs or 'amount' in fs or 'nw' in fs):
            hi, cmap = i, m
            break
    text = '\n'.join(' '.join(str(v) for v in r if v not in (None, '')) for r in rows)
    detect_style(text)
    is_pl = kind == 'PL' or (kind is None and re.search(r'PACKING\s*LIST', text[:1500], re.I) and
                             'amount' not in cmap.values())
    d = new_pl(path, 'excel') if is_pl else new_inv(path, 'excel')
    if hi is None:
        d['notes'].append('no item table heading found in the Excel file')
        return d
    for r in rows[hi + 1:]:
        cell = {f: r[j] for j, f in cmap.items() if j < len(r) and r[j] not in (None, '')}
        if not cell:
            continue
        joined = ' '.join(str(v) for v in cell.values())
        if STOP_RE.search(' '.join(str(v) for v in r[:4] if v not in (None, ''))) or \
                any(re.match(r'^\s*(sub\s*)?total', str(v or ''), re.I) for v in r):
            if not is_pl and num(cell.get('amount')) is not None:
                d['total'] = num(cell.get('amount'))
            break
        qty, amt = num(cell.get('qty')), num(cell.get('amount'))
        if (amt is None and not is_pl) or (is_pl and qty is None and cell.get('nw') is None):
            continue
        l = new_line(line=str(cell.get('line', '') or ''), item=str(cell.get('item', '') or '').strip(),
                     desc=clean(cell.get('desc', '')), qty=qty, uom=str(cell.get('uom', '') or '').upper(),
                     price=num(cell.get('price')), amount=amt, currency=str(cell.get('currency', '') or '').upper())
        if l['item'].endswith('.0'):
            l['item'] = l['item'][:-2]
        hs = re.sub(r'\D', '', str(cell.get('hs', '') or ''))
        l['hs'] = hs[:8] if len(hs) >= 8 else (hs.zfill(8) if len(hs) == 7 else hs)
        if cell.get('coo'):
            l['coo'] = countries.code(str(cell['coo'])) or (str(cell['coo']).upper() if len(str(cell['coo'])) == 2 else '')
        l['nw'], l['gw'] = num(cell.get('nw')), num(cell.get('gw'))
        if cell.get('lot'):
            l['lots'] = [str(cell['lot'])]
        d['lines'].append(l)
    if is_pl:
        h = header_fields(new_inv(path, ''), text, [], countries)
        d.update(hawb=h['hawb'], total_gw=h['total_gw'] or (sum(l['gw'] or 0 for l in d['lines']) or None),
                 total_nw=h['total_nw'] or (sum(l['nw'] or 0 for l in d['lines']) or None), packages=h['boxes'])
        return d
    header_fields(d, text, [], countries)
    finish_lines(d, countries)
    return d


def _csv_rows(path):
    import csv
    with open(path, encoding='utf-8-sig', errors='replace') as f:
        return [[c if c != '' else None for c in r] for r in csv.reader(f)]


# ======================================================================== built-in: MSD + bill of lading (from MSD PERMIT TOOL)
def msd_docs(path, text, countries):
    import msd_engine as M
    k = M.classify(text)
    if k == 'INV':
        m = M.parse_invoice(text, path, countries)
        d = new_inv(path, 'MSD')
        d.update(invoice_no=m.get('invoice_no', ''), currency=m.get('currency', ''), total=m.get('total'),
                 incoterm=m.get('incoterm', ''), hawb=m.get('hawb', ''), uen=m.get('exporter_uen', ''),
                 invoice_from=m.get('exporter', ''), ship_from=m.get('exporter', ''), coo_text=m.get('coo_text', ''),
                 ship_to=m.get('consignee', ''), ship_to_country=m.get('dest', ''), title='COMMERCIAL INVOICE',
                 order_no=m.get('sap_order', ''))
        d['direction'] = 'EXPORT' if m.get('dest') and m.get('dest') != 'SG' else 'IMPORT'
        for i, l in enumerate(m.get('lines', []), 1):
            d['lines'].append(new_line(line=str(i), item=l['item'], desc=l['desc'], hs=l.get('hs') or m.get('hs', ''),
                                       qty=l['qty'], uom=l['uom'], price=l['price'], amount=l['amount'],
                                       currency=m.get('currency', ''), lots=[b for b in l.get('batches', []) if b]))
        d['delivery'] = m.get('delivery', '')
        finish_lines(d, countries)
        return [d]
    if k == 'PL':
        m = M.parse_packing_list(text, path, countries)
        d = new_pl(path, 'MSD')
        d.update(hawb=m.get('hawb', ''), total_gw=m.get('total_gw'), total_nw=m.get('total_nw'),
                 packages=m.get('pallets'), package_unit='PAT', containers=[m['container']] if m.get('container') else [],
                 delivery=m.get('delivery', ''))
        for key, mat in m.get('materials', {}).items():
            d['lines'].append(new_line(item=key, desc=mat.get('desc', ''), qty=mat.get('qty'), uom=mat.get('uom', ''),
                                       nw=mat.get('nw'), gw=mat.get('gw')))
        return [d]
    if k == 'BL':
        return [bl_doc(path, text, countries)]
    return []


def bl_doc(path, text, countries):
    import msd_engine as M
    b = M.parse_bl(text, path, countries)
    if b.get('bl_no'):                           # "MLCWPERH19044182" read as "PERH19044182" -> the whole number
        longer = [x for x in re.findall(r'\b[A-Z0-9]{%d,}\b' % (len(b['bl_no']) + 1), text) if x.endswith(b['bl_no'])]
        if longer:
            b['bl_no'] = max(longer, key=longer.count)
    d = dict(kind='AWB', file=os.path.basename(path), path=path, notes=[], reader='bill of lading',
             hawb=b.get('bl_no', ''), mawb='', pieces=b.get('packages'), gw=b.get('gw'), departure=b.get('pol', ''),
             legs=[], flights=[], executed='', destination=b.get('pod', ''), consignee='', shipper='', fwd_job='',
             awb_hs=b.get('hs', ''), awb_terms='', issued_by=b.get('forwarder', ''))
    cg = dict(mode='Sea')
    if b.get('bl_no'):
        cg['hawb'] = b['bl_no']
        cg['mawb'] = b['bl_no']
    if b.get('vessel'):
        cg['transport_id'] = b['vessel']
    if b.get('voyage'):
        cg['flight'] = b['voyage']
    if b.get('pol'):
        cg['loading_name'] = clean(b['pol'].split(',')[0]).upper()
        cg['loading_code'] = M.port_code(b['pol'], {}) or ''
    if b.get('pod'):
        cg['discharge_name'] = clean(b['pod'].split(',')[0]).upper()
        cg['discharge_code'] = M.port_code(b['pod'], {}) or ''
    if b.get('packages'):
        cg['packages_qty'], cg['packages_unit'] = b['packages'], b.get('package_unit', '')
    if b.get('gw'):
        cg['gw_kg'] = b['gw']
    if b.get('containers'):
        cg['containers'] = [c['no'] for c in b['containers']]
        cg['container_info'] = b['containers']
    d['cargo'] = cg
    return d


# ======================================================================== AI reader (Claude API)
AI = dict(mode='off', key='', model='claude-sonnet-5-5', status='')
AI_MODELS = ['claude-sonnet-5-5', 'claude-opus-5-5', 'claude-haiku-4-5-20251001']

AI_SCHEMA = {
    'type': 'object',
    'properties': {'documents': {'type': 'array', 'items': {'type': 'object', 'properties': {
        'doc_type': {'type': 'string', 'enum': ['invoice', 'packing_list', 'air_waybill', 'bill_of_lading', 'other']},
        'invoice_no': {'type': 'string'}, 'invoice_date': {'type': 'string', 'description': 'dd/mm/yyyy'},
        'seller_name': {'type': 'string'}, 'seller_country': {'type': 'string', 'description': '2-letter ISO'},
        'buyer_name': {'type': 'string'}, 'consignee_name': {'type': 'string'},
        'consignee_country': {'type': 'string', 'description': '2-letter ISO'}, 'uen': {'type': 'string'},
        'currency': {'type': 'string'}, 'incoterm': {'type': 'string'}, 'total_amount': {'type': 'number'},
        'hawb': {'type': 'string'}, 'mawb': {'type': 'string', 'description': 'master AWB / ocean B/L, digits only for AWB'},
        'gross_weight_kg': {'type': 'number'}, 'net_weight_kg': {'type': 'number'},
        'packages': {'type': 'number'}, 'package_unit': {'type': 'string'},
        'flight_no': {'type': 'string'}, 'vessel': {'type': 'string'}, 'voyage': {'type': 'string'},
        'port_of_loading': {'type': 'string'}, 'port_of_discharge': {'type': 'string'},
        'arrival_date': {'type': 'string', 'description': 'dd/mm/yyyy'},
        'containers': {'type': 'array', 'items': {'type': 'string'}},
        'carrier_or_forwarder': {'type': 'string'},
        'lines': {'type': 'array', 'items': {'type': 'object', 'properties': {
            'line_no': {'type': 'string'}, 'item_no': {'type': 'string'}, 'description': {'type': 'string'},
            'hs_code': {'type': 'string', 'description': 'digits only, as printed'},
            'country_of_origin': {'type': 'string', 'description': '2-letter ISO code'},
            'qty': {'type': 'number'}, 'uom': {'type': 'string'}, 'unit_price': {'type': 'number'},
            'amount': {'type': 'number'}, 'net_weight_kg': {'type': 'number'}, 'gross_weight_kg': {'type': 'number'},
            'lot': {'type': 'string'}, 'eccn': {'type': 'string'}}, 'required': ['description']}}},
        'required': ['doc_type', 'lines']}}},
    'required': ['documents']}

AI_PROMPT = ('You extract data from shipping documents (commercial / proforma / tax invoices, packing lists, '
             'air waybills, bills of lading) for a Singapore customs permit. Return EVERY item line exactly as printed '
             '- do not merge, skip, round or invent anything; leave a field out when it is not printed. Amounts and '
             'quantities as plain numbers. HS code digits only. Country of origin as 2-letter ISO code. If the file '
             'holds several documents (e.g. invoice + packing list) return each one separately. Call the tool once.')


def ai_available():
    return AI['mode'] != 'off' and bool(AI['key'])


def _ai_call(content, log=None):
    import urllib.request, ssl
    body = dict(model=AI['model'], max_tokens=32000, system=AI_PROMPT,
                tools=[dict(name='save_documents', description='Save the extracted documents', input_schema=AI_SCHEMA)],
                tool_choice=dict(type='tool', name='save_documents'),
                messages=[dict(role='user', content=content)])
    req = urllib.request.Request('https://api.anthropic.com/v1/messages', data=json.dumps(body).encode('utf-8'),
                                 headers={'x-api-key': AI['key'], 'anthropic-version': '2023-06-01',
                                          'content-type': 'application/json'})
    ctx = ssl.create_default_context()
    try:
        with urllib.request.urlopen(req, timeout=600, context=ctx) as r:
            res = json.loads(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        raise RuntimeError('AI reader: HTTP %s %s' % (e.code, e.read().decode('utf-8', 'replace')[:300]))
    for blk in res.get('content', []):
        if blk.get('type') == 'tool_use':
            return blk['input'].get('documents', [])
    raise RuntimeError('AI reader: no data returned (%s)' % res.get('stop_reason'))


def ai_read(path, pages=None, log=None, chunk_pages=20):
    """Claude reads the file. Text PDFs: the text is sent; scans / Excel: the PDF itself."""
    docs = []
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xls', '.xlsm', '.csv'):
        rows = E.read_rows(path) if ext != '.csv' else _csv_rows(path)
        txt = '\n'.join('\t'.join('' if v is None else str(v) for v in r) for r in rows)
        parts = [txt[i:i + 150000] for i in range(0, len(txt), 150000)] or ['']
        for n, p in enumerate(parts, 1):
            if log:
                log('   AI reading %s (part %d/%d) ...' % (os.path.basename(path), n, len(parts)))
            docs += _ai_call([dict(type='text', text='File %s (Excel, tab separated):\n%s' % (os.path.basename(path), p))])
        return docs
    use_text = pages and all(not p.get('ocr') for p in pages)
    if use_text:
        for s in range(0, len(pages), chunk_pages):
            part = pages[s:s + chunk_pages]
            if log:
                log('   AI reading %s pages %d-%d / %d ...' % (os.path.basename(path), s + 1, s + len(part), len(pages)))
            txt = '\n'.join('--- page %d ---\n%s' % (s + i + 1, p['text']) for i, p in enumerate(part))
            docs += _ai_call([dict(type='text', text='File %s, pages %d-%d of %d (text layer):\n%s' % (
                os.path.basename(path), s + 1, s + len(part), len(pages), txt))])
        return docs
    from pypdf import PdfReader, PdfWriter
    import io
    rd = PdfReader(path)
    n = len(rd.pages)
    for s in range(0, n, chunk_pages):
        w = PdfWriter()
        for i in range(s, min(n, s + chunk_pages)):
            w.add_page(rd.pages[i])
        buf = io.BytesIO()
        w.write(buf)
        if log:
            log('   AI reading %s pages %d-%d / %d (scan) ...' % (os.path.basename(path), s + 1, min(n, s + chunk_pages), n))
        docs += _ai_call([dict(type='document', source=dict(type='base64', media_type='application/pdf',
                                                          data=base64.b64encode(buf.getvalue()).decode('ascii'))),
                          dict(type='text', text='File %s, pages %d-%d of %d.' % (os.path.basename(path), s + 1,
                                                                                  min(n, s + chunk_pages), n))])
    return docs


def from_ai(path, raw_docs, countries):
    """AI JSON -> common document shape. Pieces of the same document (page chunks) are joined."""
    C = countries
    merged = []
    for r in raw_docs:
        t = r.get('doc_type', 'other')
        key = (t, (r.get('invoice_no') or '').strip(), (r.get('hawb') or '').strip())
        prev = next((m for m in merged if m['doc_type'] == t and (
            (key[1] and (m.get('invoice_no') or '').strip() == key[1]) or (not key[1] and not m.get('invoice_no')))), None)
        if prev and t in ('invoice', 'packing_list'):
            prev['lines'] = prev.get('lines', []) + r.get('lines', [])
            for k, v in r.items():
                if k != 'lines' and v not in (None, '', []) and prev.get(k) in (None, '', []):
                    prev[k] = v
        else:
            merged.append(dict(r))
    out = []
    for r in merged:
        t = r.get('doc_type')
        g = lambda k: r.get(k) if r.get(k) not in (None, '') else None
        if t == 'invoice':
            d = new_inv(path, 'AI')
            d.update(invoice_no=g('invoice_no') or '', invoice_date=g('invoice_date') or '', currency=(g('currency') or '').upper(),
                     incoterm=(g('incoterm') or '').upper()[:3], total=g('total_amount'), hawb=g('hawb') or '',
                     mawb=re.sub(r'\D', '', g('mawb') or '') if re.search(r'\d{3}\D?\d{8}', g('mawb') or '') else (g('mawb') or ''),
                     total_gw=g('gross_weight_kg'), total_nw=g('net_weight_kg'),
                     boxes=int(g('packages')) if g('packages') else None, package_unit=g('package_unit') or '',
                     invoice_from=g('seller_name') or '', ship_from=g('seller_name') or '',
                     ship_to=g('consignee_name') or g('buyer_name') or '', invoice_to=g('buyer_name') or '',
                     ship_to_country=(g('consignee_country') or '').upper()[:2],
                     ship_from_country=(g('seller_country') or '').upper()[:2], uen=g('uen') or '',
                     carrier=g('carrier_or_forwarder') or '', title='INVOICE (AI)')
            d['direction'] = 'EXPORT' if d['ship_from_country'] == 'SG' and d['ship_to_country'] not in ('', 'SG') else 'IMPORT'
            for i, l in enumerate(r.get('lines', []), 1):
                coo = (l.get('country_of_origin') or '').upper()
                coo = coo if len(coo) == 2 else (C.code(coo) if coo else '')
                d['lines'].append(new_line(line=str(l.get('line_no') or i), item=l.get('item_no') or '',
                                           desc=clean(l.get('description', '')), hs=re.sub(r'\D', '', l.get('hs_code') or '')[:8],
                                           coo=coo, qty=l.get('qty'), uom=(l.get('uom') or '').upper(),
                                           price=l.get('unit_price'), amount=l.get('amount'), currency=d['currency'],
                                           nw=l.get('net_weight_kg'), gw=l.get('gross_weight_kg'),
                                           lots=[l['lot']] if l.get('lot') else [], eccn=l.get('eccn') or ''))
            finish_lines(d, C)
            d['notes_ai'] = 'read by AI - check item no / HS / qty / amounts against the invoice'
            out.append(d)
        elif t == 'packing_list':
            d = new_pl(path, 'AI')
            d.update(hawb=g('hawb') or '', total_gw=g('gross_weight_kg'), total_nw=g('net_weight_kg'),
                     packages=int(g('packages')) if g('packages') else None, package_unit=g('package_unit') or '',
                     containers=r.get('containers') or [], invoice_refs=[g('invoice_no')] if g('invoice_no') else [])
            for l in r.get('lines', []):
                d['lines'].append(new_line(item=l.get('item_no') or '', desc=clean(l.get('description', '')),
                                           qty=l.get('qty'), uom=(l.get('uom') or '').upper(),
                                           nw=l.get('net_weight_kg'), gw=l.get('gross_weight_kg')))
            out.append(d)
        elif t in ('air_waybill', 'bill_of_lading'):
            sea = t == 'bill_of_lading'
            cg = dict(mode='Sea' if sea else 'Air')
            for k_ai, k_cg in (('mawb', 'mawb'), ('hawb', 'hawb'), ('vessel', 'transport_id'), ('arrival_date', 'date')):
                if g(k_ai):
                    cg[k_cg] = g(k_ai)
            if g('flight_no') or g('voyage'):
                cg['flight'] = (g('flight_no') or g('voyage')).replace(' ', '')
            if cg.get('mawb') and not sea:
                cg['mawb'] = re.sub(r'\D', '', cg['mawb'])
            if g('packages'):
                cg['packages_qty'], cg['packages_unit'] = int(g('packages')), g('package_unit') or ''
            if g('gross_weight_kg'):
                cg['gw_kg'] = g('gross_weight_kg')
            for side, k in (('loading', 'port_of_loading'), ('discharge', 'port_of_discharge')):
                v = (g(k) or '').upper()
                if v:
                    iata = re.search(r'\b([A-Z]{3})\b', v)
                    if not sea and iata and iata.group(1) in E.AIRPORTS:
                        cg[side + '_code'], cg[side + '_name'] = E.AIRPORTS[iata.group(1)]
                    else:
                        code = re.search(r'\b([A-Z]{5})\b', v)
                        cg[side + '_code'] = code.group(1) if code else ''
                        cg[side + '_name'] = clean(re.sub(r'\b[A-Z]{5}\b|[(),]', ' ', v))
            if r.get('containers'):
                cg['containers'] = r['containers']
            out.append(dict(kind='AWB', file=os.path.basename(path), path=path, reader='AI',
                            notes=['read by AI - check MAWB / flight / pieces / weight'], hawb=g('hawb') or '',
                            mawb=cg.get('mawb', ''), pieces=cg.get('packages_qty'), gw=cg.get('gw_kg'),
                            departure=g('port_of_loading') or '', legs=[], flights=[], executed='',
                            destination=g('port_of_discharge') or '', consignee=g('consignee_name') or '',
                            shipper=g('seller_name') or '', fwd_job='', awb_hs='', awb_terms='',
                            issued_by=g('carrier_or_forwarder') or '', cargo=cg))
    return out


# ======================================================================== dispatcher
def classify_generic(text):
    head = text[:2500].upper()
    if re.search(r'PACKING\s*LIST|PACKING\s*SLIP|SHIPPING\s*SPEC|DELIVERY\s*NOTE', head) and \
            not re.search(r'COMMERCIAL\s+INVOICE|TAX\s+INVOICE|PRO\s*-?\s*FORMA', head):
        return 'PL'
    if re.search(r'INVOICE', head):
        return 'INV'
    if re.search(r'BILL\s+OF\s+LADING|SEA\s*WAYBILL|B/L\s*NO', head):
        return 'BL'
    return ''


def read_file(path, countries=None, profiles=None, log=None):
    """Any supplier document -> list of documents in the common shape."""
    C = countries or Countries()
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xls', '.xlsm', '.csv'):
        rows = E.read_rows(path) if ext != '.csv' else _csv_rows(path)
        try:                                   # known Excel layouts (e.g. VF / Infor Nexus packing list)
            import export_docs
            known = export_docs.read_excel(path, rows, C, log)
            if known:
                return known
        except Exception as e:
            if log:
                log('   exporter Excel reader: %s - general reader used' % e)
        text = '\n'.join(' '.join(str(v) for v in r if v not in (None, '')) for r in rows[:80])
        prof = profiles.find(text) if profiles else None
        if AI['mode'] == 'always' and ai_available():
            return from_ai(path, ai_read(path, None, log), C)
        d = generic_excel(path, C, prof, prof.get('kind') if prof else None)
        if prof:
            d['reader'] = 'profile: ' + prof['name']
        if quality(d) < 0.5 and ai_available():
            if log:
                log('   general reader not sure (%s) - AI reading' % os.path.basename(path))
            return from_ai(path, ai_read(path, None, log), C)
        return [d]
    pages = E.read_pages(path, log)
    if not pages:
        return [dict(kind='', file=os.path.basename(path), path=path, notes=['empty file'])]
    try:                                       # known exporter layouts (UD Trucks, Kalmar, Getinge, Cummins ...)
        import export_docs
        known = export_docs.read(path, pages, C, log)
        if known:
            return known
    except Exception as e:
        if log:
            log('   exporter reader: %s - general reader used' % e)
    if AI['mode'] == 'always' and ai_available():
        return from_ai(path, ai_read(path, pages, log), C)
    segs = E.segment_pages(pages)
    if len(segs) <= 1:
        return _read_pdf_pages(path, pages, C, profiles, log)
    # one PDF holding several documents (e.g. DHL pack: air waybill + e-mail + invoice + packing lists)
    out = []
    for kind, pgs in segs:
        if kind == 'EMAIL':
            continue
        if kind == 'AWB':
            out.append(E.parse_awb(path, pgs))
        elif kind == 'PL' and any(E.AG_PL_RE.match(ln.strip()) for p in pgs for ln in p['text'].splitlines()):
            out.append(E.parse_agilent_pl(path, pgs))
        else:
            out += [d for d in _read_pdf_pages(path, pgs, C, profiles, log) if d.get('kind')]
    if log:
        log('   %s: %d documents in one PDF (%s)' % (os.path.basename(path), len(out),
                                                    ', '.join(d.get('kind', '?') for d in out)))
    return out or _read_pdf_pages(path, pages, C, profiles, log)


def _read_pdf_pages(path, pages, C, profiles=None, log=None):
    text = '\n'.join(p['text'] for p in pages)
    first = pages[0]['text']
    out = []
    # 1. built-in readers
    if E.classify(first) == 'INV' or E.classify(text[:6000]) == 'INV':
        out = E.parse_pdf_pages(path, pages, C)
    elif E.classify_awb(first):
        out = [E.parse_awb(path, pages)]
    elif re.search(r'\bMSD\b|MERCK\s+SHARP', text[:6000], re.I):
        try:
            out = msd_docs(path, text, C)
        except Exception as e:
            out = []
            if log:
                log('   MSD reader: %s' % e)
    if not out and classify_generic(first) == 'BL':
        try:
            out = [bl_doc(path, text, C)]
        except Exception:
            out = []
    # 2./3. supplier profile or general reader
    if not out:
        kind = classify_generic(first) or ('INV' if re.search(r'INVOICE', text, re.I) else '')
        prof = profiles.find(text, kind) if profiles else None
        if prof and prof.get('kind'):
            kind = prof['kind']
        if kind in ('INV', 'PL'):
            out = [generic_pdf(path, pages, kind, C, prof)]
    # 4. AI fallback
    bad = not out or all(quality(d) < 0.5 for d in out if d.get('kind') in ('INV', 'PL'))
    if bad and ai_available() and AI['mode'] in ('fallback', 'always'):
        if log:
            log('   built-in / general reader could not read %s well - AI reading' % os.path.basename(path))
        try:
            ai = from_ai(path, ai_read(path, pages, log), C)
            if ai:
                return ai
        except Exception as e:
            if log:
                log('ERROR %s' % e)
    if not out:
        out = [dict(kind='', file=os.path.basename(path), path=path,
                    notes=['not recognised (no invoice / packing list / waybill layout found)' +
                           (' - switch on AI reading or teach this supplier on the Suppliers tab' if not ai_available() else '')])]
    for d in out:
        if d.get('kind') in ('INV', 'PL') and quality(d) < 0.5:
            d.setdefault('notes', []).append('reader not sure about this document (%s) - check the items, or teach the '
                                             'supplier on the Suppliers tab' % d.get('reader', ''))
    return out


# ======================================================================== teaching helpers (Suppliers tab)
def guess_supplier(text):
    for ln in text.splitlines()[:8]:
        if COMPANY_RE.search(ln) and not re.search(r'invoice|packing|to\s*:', ln, re.I):
            return clean(ln)[:60]
    m = re.search(r'(?:shipper|exporter|seller|vendor)\s*:\s*(.+)', text, re.I)
    return clean(LABEL_SPLIT.split(m.group(1))[0])[:60] if m else ''


def heading_candidates(path):
    """Possible item-table heading lines of a file -> ([{label, headings[]}], text)."""
    ext = os.path.splitext(path)[1].lower()
    cand = []
    if ext in ('.xlsx', '.xls', '.xlsm', '.csv'):
        rows = E.read_rows(path) if ext != '.csv' else _csv_rows(path)
        text = '\n'.join(' '.join(str(v) for v in r if v not in (None, '')) for r in rows[:80])
        for r in rows[:60]:
            hs = [str(v).strip() for v in r if v not in (None, '') and not is_num(v)]
            if len(hs) >= 3:
                sc = len({match_field(h) for h in hs} - {''})
                cand.append(dict(label=' | '.join(hs)[:150], headings=hs, score=sc))
    else:
        pages = E.read_pages(path)
        text = '\n'.join(p['text'] for p in pages)
        for pg in pages[:2]:
            lines = word_lines(pg['words'])
            for i, ln in enumerate(lines):
                for span in (1, 2):
                    if i + span > len(lines):
                        continue
                    if span == 2 and abs(lines[i + 1][0]['top'] - ln[0]['top']) > 14:
                        continue
                    ws = [w['text'] for l in lines[i:i + span] for w in l if not is_num(w['text'])]
                    if len(ws) < 3:
                        continue
                    sc = len({match_field(h) for h in ws} - {''})
                    cand.append(dict(label=(' | '.join(ws))[:150] + ('   (2 lines)' if span == 2 else ''), headings=ws, score=sc))
    seen, out = set(), []
    for c in sorted(cand, key=lambda c: -c['score']):
        if c['label'] in seen:
            continue
        seen.add(c['label'])
        out.append(c)
    return out[:20], text


def read_with_profile(path, prof, countries):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xls', '.xlsm', '.csv'):
        return generic_excel(path, countries, prof, prof.get('kind'))
    pages = E.read_pages(path)
    return generic_pdf(path, pages, prof.get('kind') or 'INV', countries, prof)
