# -*- coding: utf-8 -*-
"""FESTO PERMIT TOOL - engine (shared rules from the Agilent / MSD tools; Festo readers in festo_docs.py)
Reads Agilent Technologies invoices (PDF: IMPORT / EXPORT INVOICE), extracts the item lines and cargo
details, and fills the dec-upload template (Header / Item / Invoice / Casc / Container) - one Excel per
HAWB - plus a CHECK REPORT and EXCEL_ALL.zip.
"""
import math
import logging as _logging
_logging.getLogger('pdfminer').setLevel(_logging.ERROR)   # harmless font warnings (FontBBox ...) not shown
import os, re, zipfile, datetime, collections, json
from xml.sax.saxutils import escape


def tne_round(v):
    """Container weight in TNE as a round figure: x.5 and above -> up, below x.5 -> down (5.5 -> 6, 5.49 -> 5), at least 1."""
    return max(1, int(math.floor(round(float(v), 6) + 0.5)))


# ---------------------------------------------------------------- helpers (shared with RS COMPO TOOL)
def h8(v):
    """Normalise an HS code to an 8-digit string."""
    if v is None:
        return ''
    if isinstance(v, float):
        v = int(v)
    s = str(v).strip().replace('.', '').replace(' ', '')
    if s.isdigit():
        return s.zfill(8)
    return s


def num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').strip())
    except ValueError:
        return 0.0


def read_rows(path, sheet=None):
    """Return all rows (list of lists) of a sheet in .xlsx/.xlsm/.xls."""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.xls':
        try:
            import xlrd
        except ImportError:
            raise RuntimeError('Reading .xls needs the xlrd package. Save the file as .xlsx or install xlrd.')
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_name(sheet) if sheet else book.sheet_by_index(0)
        out = []
        for i in range(sh.nrows):
            row = []
            for j, c in enumerate(sh.row(i)):
                v = c.value
                if c.ctype == xlrd.XL_CELL_EMPTY:
                    v = None
                elif c.ctype == xlrd.XL_CELL_NUMBER and float(v).is_integer():
                    v = int(v)
                row.append(v)
            out.append(row)
        return out
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    out = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return out


def find_header(rows, must, limit=40):
    """Find the header row containing all names in `must`; return (index, {name_lower: col})."""
    for i, r in enumerate(rows[:limit]):
        names = {str(v).strip().lower(): j for j, v in enumerate(r) if v is not None and str(v).strip()}
        if all(m.lower() in names for m in must):
            return i, names
    raise RuntimeError('Header row with %s not found' % ', '.join(must))




# ---------------------------------------------------------------- xlsx template writer (keeps every sheet/format)
def col2num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def num2col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def xml_unescape(s):
    return s.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&')


class Sheet:
    ROW_RE = re.compile(r'<row\b[^>]*?(?:/>|>.*?</row>)', re.S)
    CELL_RE = re.compile(r'<c\b[^>]*?(?:/>|>.*?</c>)', re.S)

    def __init__(self, xml):
        self.xml = xml
        m = re.search(r'<sheetData\s*/>|<sheetData>(.*?)</sheetData>', xml, re.S)
        self.span = m.span()
        inner = m.group(1) or ''
        self.rows = collections.OrderedDict()
        for rm in self.ROW_RE.finditer(inner):
            rx = rm.group(0)
            rn = int(re.search(r'\br="(\d+)"', rx).group(1))
            head = re.match(r'<row\b[^>]*?(?=/?>)', rx).group(0)
            cells = collections.OrderedDict()
            for cm in self.CELL_RE.finditer(rx):
                cx = cm.group(0)
                ref = re.search(r'\br="([A-Z]+)\d+"', cx).group(1)
                cells[ref] = cx
            self.rows[rn] = [head, cells, False]

    def _row(self, rn):
        if rn not in self.rows:
            self.rows[rn] = ['<row r="%d"' % rn, collections.OrderedDict(), True]
        return self.rows[rn]

    def get(self, ref):
        col, rn = re.match(r'([A-Z]+)(\d+)', ref).groups()
        r = self.rows.get(int(rn))
        return r[1].get(col) if r else None

    def set(self, ref, value, style=None):
        if value is None or value == '':
            return
        col, rn = re.match(r'([A-Z]+)(\d+)', ref).groups()
        row = self._row(int(rn))
        old = row[1].get(col)
        s_attr = ''
        if old:
            if '<f>' in old or '<f ' in old:
                raise RuntimeError('Cell %s has a formula' % ref)
            m = re.search(r'\bs="(\d+)"', old)
            if m:
                s_attr = ' s="%s"' % m.group(1)
        if style is not None:
            s_attr = ' s="%d"' % style
        if isinstance(value, bool):
            value = str(value)
        if isinstance(value, (int, float)):
            cx = '<c r="%s"%s><v>%s</v></c>' % (ref, s_attr, repr(float(value)) if isinstance(value, float) else value)
        else:
            cx = '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s_attr, escape(str(value)))
        row[1][col] = cx
        row[2] = True

    def set_cached(self, ref, val):
        cx = self.get(ref)
        if not cx or '<f' not in cx:
            return
        col, rn = re.match(r'([A-Z]+)(\d+)', ref).groups()
        if re.search(r'<v>.*?</v>|<v/>', cx, re.S):
            cx = re.sub(r'<v>.*?</v>|<v/>', '<v>%s</v>' % escape(val), cx, count=1, flags=re.S)
        else:
            cx = cx.replace('</c>', '<v>%s</v></c>' % escape(val))
        self.rows[int(rn)][1][col] = cx

    def tostring(self):
        out = []
        for rn in sorted(self.rows):
            head, cells, changed = self.rows[rn]
            if changed:
                head = re.sub(r'\sspans="[^"]*"', '', head)
            body = ''.join(cells[c] for c in sorted(cells, key=col2num))
            out.append(head + ('>' + body + '</row>' if body else '/>'))
        sd = '<sheetData>' + ''.join(out) + '</sheetData>'
        xml = self.xml[:self.span[0]] + sd + self.xml[self.span[1]:]
        # dimension
        maxr = max(self.rows) if self.rows else 1
        maxc = max((col2num(c) for r in self.rows.values() for c in r[1]), default=1)
        xml = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:%s%d"/>' % (num2col(maxc), maxr), xml, count=1)
        return xml


class Template:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.infos = z.infolist()
        self.data = {i.filename: z.read(i.filename) for i in self.infos}
        z.close()
        self.sst = []
        if 'xl/sharedStrings.xml' in self.data:
            x = self.data['xl/sharedStrings.xml'].decode('utf-8')
            for si in re.findall(r'<si>(.*?)</si>', x, re.S):
                self.sst.append(xml_unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))))
        wb = self.data['xl/workbook.xml'].decode('utf-8')
        rels = self.data['xl/_rels/workbook.xml.rels'].decode('utf-8')
        relmap = {m.group(1): m.group(2) for m in re.finditer(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels)}
        relmap.update({m.group(2): m.group(1) for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
        self.sheets = collections.OrderedDict()
        for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb):
            tgt = relmap[m.group(2)].lstrip('/')
            self.sheets[xml_unescape(m.group(1))] = tgt if tgt.startswith('xl/') else 'xl/' + tgt
        self._add_styles()

    def _add_styles(self):
        st = self.data['xl/styles.xml'].decode('utf-8')
        ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', st)] + [163]
        a, b = max(ids) + 1, max(ids) + 2
        nf = '<numFmt numFmtId="%d" formatCode="0.0000"/><numFmt numFmtId="%d" formatCode="dd/mm/yyyy"/>' % (a, b)
        if '<numFmts' in st:
            st = re.sub(r'<numFmts count="(\d+)">', lambda m: '<numFmts count="%d">' % (int(m.group(1)) + 2), st, count=1)
            st = st.replace('</numFmts>', nf + '</numFmts>', 1)
        else:
            st = st.replace('<fonts', '<numFmts count="2">' + nf + '</numFmts><fonts', 1)
        m = re.search(r'<cellXfs count="(\d+)">', st)
        base = int(m.group(1))
        xfs = ('<xf numFmtId="%d" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
               '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
               '<xf numFmtId="%d" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
               '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
               '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>') % (a, b)
        st = st.replace(m.group(0), '<cellXfs count="%d">' % (base + 5), 1)
        i = st.find('</cellXfs>')
        st = st[:i] + xfs + st[i:]
        self.data['xl/styles.xml'] = st.encode('utf-8')
        self.S4, self.S2, self.SDATE, self.STXT, self.SWRAP = base, base + 1, base + 2, base + 3, base + 4
        wb = self.data['xl/workbook.xml'].decode('utf-8')
        if 'fullCalcOnLoad' not in wb:
            wb = re.sub(r'<calcPr\b', '<calcPr fullCalcOnLoad="1"', wb, count=1)
        self.data['xl/workbook.xml'] = wb.encode('utf-8')

    def sheet(self, name):
        return Sheet(self.data[self.sheets[name]].decode('utf-8'))

    def headers(self, sh):
        """{header text lower: column letter} from row 1."""
        out = {}
        r = sh.rows.get(1)
        if not r:
            return out
        for col, cx in r[1].items():
            v = None
            m = re.search(r'<v>(.*?)</v>', cx, re.S)
            if 't="s"' in cx and m:
                v = self.sst[int(m.group(1))]
            elif 't="inlineStr"' in cx:
                v = xml_unescape(''.join(re.findall(r'<t[^>]*>(.*?)</t>', cx, re.S)))
            elif m:
                v = xml_unescape(m.group(1))
            if v:
                out.setdefault(v.strip().lower(), col)
        return out

    def save(self, path, changed):
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
            for info in self.infos:
                n = info.filename
                data = changed.get(n, self.data[n])
                if isinstance(data, str):
                    data = data.encode('utf-8')
                zo.writestr(info, data)


def template_lists(path):
    """Dropdown lists from the template's List-* sheets: {sheet: [(text, code), ...]}."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for ws in wb.worksheets:
        if ws.title.startswith('List-'):
            out[ws.title] = [(str(r[0]), '' if len(r) < 2 or r[1] is None else str(r[1]))
                             for r in ws.iter_rows(values_only=True) if r and r[0] is not None]
    wb.close()
    return out


def lookup(lst, text):
    for t, c in lst:
        if t == text:
            return c
    return ''



def sheet_clear_from(sh, first_row=2):
    """Remove every value below the header (keeps cell styles, formulas untouched)."""
    for rn, row in sh.rows.items():
        if rn < first_row:
            continue
        for col, cx in list(row[1].items()):
            if '<f>' in cx or '<f ' in cx:
                continue
            if '<v' in cx or '<is>' in cx:
                m = re.search(r'\bs="(\d+)"', cx)
                row[1][col] = '<c r="%s%d"%s/>' % (col, rn, ' s="%s"' % m.group(1) if m else '')
                row[2] = True


Sheet.clear_from = sheet_clear_from


# ================================================================ constants
DEFAULT_UOM_RULES = [
    ('KGM', 'NET WEIGHT', 'x', 1), ('TNE', 'NET WEIGHT', '/', 1000),
    ('NMB', 'QTY', 'x', 1), ('TEN', 'QTY', '/', 10), ('TPR', 'QTY', '/', 10),
    ('CEN', 'QTY', '/', 100), ('MIL', 'QTY', '/', 1000),
]
# invoice UOMs that are a plain piece count (QTY based HS UOM rules only use these)
COUNT_UOMS = {'EA', 'EACH', 'PC', 'PCS', 'PCE', 'NMB', 'UN', 'UNIT', 'UNITS', 'NOS', 'NO', 'ST', 'SET'}
KG_UOMS = {'KG', 'KGS', 'KGM'}

DEFAULT_DESC_TEMPLATE = '{ITEM} , {DESC} , QTY : {QTY} {UOM} , NW : {NW} KGM , GW : {GW} KGM'
DESC_VARS = ('ITEM', 'DESC', 'QTY', 'UOM', 'NW', 'GW', 'HS', 'COO', 'BATCH', 'INVOICE')
REMARK_VARS = ('INV', 'HAWB', 'DELIVERY', 'JOB', 'PO')

# Permit type -> Declaration type code (master prompt)
PERMIT_SCHEME = collections.OrderedDict([('IG', 'GST'), ('ME', 'APS'), ('IT', 'SFZ'), ('IM', 'SFZ'),
                                         ('OD', 'DRT'), ('OT', 'DRT')])
INSURANCE_PERMITS = ('ME', 'IM', 'IT', 'IG')
TERM_CHARGES = {'CIF': (), 'CFR': ('ins',), 'CNI': ('frt',), 'EXW': ('frt', 'ins'), 'FOB': ('frt', 'ins'), 'FAS': ('frt', 'ins')}
TERM_MAP = {'DAP': 'CIF', 'FCA': 'FOB'}

COUNTRY_ALIASES = {
    'SOUTH KOREA': 'KR', 'KOREA': 'KR', 'REPUBLIC OF KOREA': 'KR', 'KOREA, REPUBLIC OF': 'KR',
    'UNITED KINGDOM': 'GB', 'UK': 'GB', 'GREAT BRITAIN': 'GB', 'ENGLAND': 'GB',
    'USA': 'US', 'UNITED STATES': 'US', 'UNITED STATES OF AMERICA': 'US', 'U.S.A.': 'US',
    'NETHERLANDS': 'NL', 'THE NETHERLANDS': 'NL', 'HOLLAND': 'NL', 'VIETNAM': 'VN', 'VIET NAM': 'VN',
    'TAIWAN': 'TW', 'HONG KONG': 'HK', 'CHINA': 'CN', "PEOPLE'S REPUBLIC OF CHINA": 'CN',
    'RUSSIA': 'RU', 'CZECH REPUBLIC': 'CZ', 'UAE': 'AE', 'UNITED ARAB EMIRATES': 'AE',
    'IRELAND': 'IE', 'SINGAPORE': 'SG', 'AUSTRALIA': 'AU', 'NEW ZEALAND': 'NZ', 'JAPAN': 'JP',
    'MALAYSIA': 'MY', 'THAILAND': 'TH', 'INDONESIA': 'ID', 'PHILIPPINES': 'PH', 'INDIA': 'IN',
    'GERMANY': 'DE', 'FRANCE': 'FR', 'ITALY': 'IT', 'SPAIN': 'ES', 'SWITZERLAND': 'CH', 'BELGIUM': 'BE',
    'CANADA': 'CA', 'BRAZIL': 'BR', 'MEXICO': 'MX', 'SWEDEN': 'SE', 'DENMARK': 'DK', 'AUSTRIA': 'AT',
    'PUERTO RICO': 'PR', 'ISRAEL': 'IL', 'TURKEY': 'TR', 'SAUDI ARABIA': 'SA', 'SOUTH AFRICA': 'ZA',
}

UOM_RULES = collections.OrderedDict()


def set_uom_rules(rules=None):
    global UOM_RULES
    out = collections.OrderedDict()
    for r in (rules if rules else DEFAULT_UOM_RULES):
        try:
            uom, basis, calc, factor = r
            uom = str(uom).strip().upper()
            basis = 'NET WEIGHT' if str(basis).strip().upper().startswith('N') else 'QTY'
            calc = '/' if str(calc).strip() in ('/', 'divide', 'DIVIDE') else 'x'
            factor = float(str(factor).replace(',', '').strip())
            if uom and factor > 0:
                out[uom] = (basis, calc, factor)
        except (ValueError, TypeError):
            continue
    UOM_RULES = out
    return out


def rule_text(uom):
    r = UOM_RULES.get(uom)
    if not r:
        return '(no rule)'
    basis, calc, f = r
    b = 'net weight' if basis == 'NET WEIGHT' else 'qty'
    return b if f == 1 else '%s %s %g' % (b, '*' if calc == 'x' else '/', f)


set_uom_rules()


# ================================================================ small helpers
NUM = r'\d[\d,]*(?:\.\d+)?'


def fnum(s):
    try:
        return float(str(s).replace(',', '').strip())
    except (TypeError, ValueError):
        return None


def fmt_plain(v):
    """72000.0 -> '72000', 7567.172 -> '7567.172' (no thousands separator)."""
    if v is None:
        return ''
    s = ('%.6f' % v).rstrip('0').rstrip('.')
    return s


def fmt_thou(v):
    """2088.0 -> '2,088', 10105.602 -> '10,105.602'."""
    if v is None:
        return ''
    s = '{:,.3f}'.format(round(v, 3)).rstrip('0').rstrip('.')
    return s


def clean(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def squash(s):
    return re.sub(r'[^A-Z0-9]', '', str(s or '').upper())


def pdf_text(path):
    """Text of all pages, layout preserved (pypdf layout mode)."""
    from pypdf import PdfReader
    r = PdfReader(path)
    out = []
    for p in r.pages:
        try:
            t = p.extract_text(extraction_mode='layout')
        except Exception:                       # some PDFs (e.g. B/L drafts) break layout mode
            try:
                t = p.extract_text()
            except Exception:
                t = ''
        out.append(t or '')
    return '\n'.join(out)


# ================================================================ country lookup
class Countries:
    def __init__(self, pairs=None):
        self.code2name = {}
        self.name2code = {}
        for code, name in (pairs or []):
            code, name = str(code).strip().upper(), str(name).strip().upper()
            if len(code) == 2:
                self.code2name[code] = name
                self.name2code[name] = code
        for n, c in COUNTRY_ALIASES.items():
            self.name2code.setdefault(n, c)
            self.code2name.setdefault(c, n)
        self.names = sorted(self.name2code, key=len, reverse=True)

    def code(self, text):
        t = clean(text).upper().strip(' .,:')
        if not t:
            return ''
        if len(t) == 2 and t in self.code2name:
            return t
        if t in self.name2code:
            return self.name2code[t]
        for n in self.names:
            if re.search(r'(?<![A-Z])%s(?![A-Z])' % re.escape(n), t):
                return self.name2code[n]
        return ''

    def find_in(self, text, exclude=('SG',)):
        """Country code of the last country name found in text (excluding codes in exclude)."""
        best = None
        U = text.upper()
        for n in self.names:
            for m in re.finditer(r'(?<![A-Z])%s(?![A-Z])' % re.escape(n), U):
                c = self.name2code[n]
                if c in exclude:
                    continue
                if best is None or m.start() > best[0]:
                    best = (m.start(), c)
        return best[1] if best else ''


LOCATION_NAMES = {'PPZ': 'PASIR PANJANG FTZ', 'ME': 'MAJOR EXPORTER SCHEME', 'KZ': 'KEPPEL FTZ', 'JZ': 'JURONG FTZ', 'TZ': 'TUAS TERMINAL GATEWAY', 'CZ': 'CHANGI FTZ', 'O': 'OTHERS'}

def read_email_file(path):
    """Text of an instruction e-mail saved as .txt / .eml (body) / .msg (needs extract_msg)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.eml':
        import email
        from email import policy
        with open(path, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=policy.default)
        body = msg.get_body(preferencelist=('plain', 'html'))
        t = body.get_content() if body else ''
        return re.sub(r'<[^>]+>', '\n', t)
    if ext == '.msg':
        try:
            import extract_msg
        except ImportError:
            raise RuntimeError('Reading .msg needs the extract_msg package - copy the e-mail text and paste it instead')
        return extract_msg.Message(path).body or ''
    if ext == '.pdf':
        return pdf_text(path)
    with open(path, encoding='utf-8', errors='replace') as f:
        return f.read()


def load_master(path):
    """Product master: Item No -> {hs, uom, coo, casc, brand, desc}."""
    out = collections.OrderedDict()
    if not path or not os.path.exists(path):
        return out
    rows = read_rows(path)
    hi, cols = find_header(rows, ['Item No'])
    def g(r, name):
        c = cols.get(name.lower())
        if c is None or c >= len(r) or r[c] is None:
            return ''
        return str(r[c]).strip()
    for r in rows[hi + 1:]:
        it = g(r, 'Item No')
        if it.endswith('.0'):
            it = it[:-2]
        if not it:
            continue
        hs = g(r, 'HS Code')
        out[it] = dict(desc=g(r, 'Description'), hs=h8(hs)[:8] if hs else '', uom=g(r, 'HS UOM').upper(),
                       coo=g(r, 'COO').upper(), casc=g(r, 'CASC Product Code 1'), brand=g(r, 'Brand'),
                       remarks=g(r, 'Remarks'))
    return out


def save_master(path, master):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = 'MASTER'
    ws.append(MASTER_COLS)
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='1F4E78')
    for it, d in master.items():
        ws.append([it, d.get('desc', ''), d.get('hs', ''), d.get('uom', ''), d.get('coo', ''), d.get('casc', ''),
                   d.get('brand', ''), d.get('remarks', '')])
    for col, w in zip('ABCDEFGH', (12, 45, 12, 9, 7, 20, 14, 40)):
        ws.column_dimensions[col].width = w
    for r in ws.iter_rows(min_row=2):
        r[0].number_format = '@'
        r[2].number_format = '@'
    ws.freeze_panes = 'A2'
    wb.save(path)


HS_LIST_INFO = {'control_column': '', 'rpn': set()}
PRODUCT_CODES = {}            # hs8 -> CA product codes (PRODUCT CODE sheet of the UPDATE HS CODE Excel)
NOT_CONTROL_WORDS = {'NON CONTROL ITEM', '', 'N', 'NO', 'NON', 'NIL', 'NONE', 'NA', 'N/A', '-', '0', 'FALSE', 'NOT CONTROLLED',
                     'NON CONTROLLED', 'NON-CONTROLLED', 'NON CONTROL', 'NON-CONTROL', 'NOT CONTROL', 'UNCONTROLLED'}
YES_CONTROL_WORDS = {'Y', 'YES', '1', 'TRUE', 'C', 'CONTROL', 'CONTROLLED', 'CONTROLLED ITEM', 'CONTROL ITEM'}


def control_from_list(value, default_code):
    """HS list control value -> (controlled?, CASC code).  Y / YES / CONTROLLED -> default code (MISC);
    a code such as IDAMISC -> that code; N / NO / NON-CONTROLLED / blank -> not controlled."""
    v = clean(value).upper()
    if v in NOT_CONTROL_WORDS:
        return False, ''
    if v in YES_CONTROL_WORDS:
        return True, default_code
    return True, v if re.fullmatch(r'[A-Z0-9]{3,15}', v) else default_code


def _hs_rows(path):
    """Values only (formula columns ignored) of the first sheet."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    ws = wb.worksheets[0]
    rows = []
    for r in ws.iter_rows(values_only=True):
        rows.append([None if (isinstance(v, str) and v.startswith('=')) or type(v).__name__ == 'ArrayFormula' else v
                     for v in r])
    wb.close()
    return rows


def hs_list_from_excel(path):
    """UPDATE HS CODE Excel -> data dict {'hs': {code: [matched, uom, desc, ctl_imp, ctl_exp, ca_imp, ca_exp, duty]},
    'rpn': [6-digit codes], 'columns': {...}}.
    Understands the APX layout (HS Code | HS Description | HS UOM | ... | Import | Export | Transhipment (CA agency) |
    Import | Export | Transhipment (Control / Non Control) | Duty items | ... RPN check table) and simple lists
    ('HS Code' + 'HS UOM' (+ 'Control'))."""
    rows = _hs_rows(path) if path.lower().endswith(('.xlsx', '.xlsm')) else read_rows(path)
    hi = next((i for i, r in enumerate(rows[:40]) if any(str(v).strip().lower() in ('hs code', 'input hs code')
                                                           for v in r if v is not None)), None)
    if hi is None:
        raise RuntimeError('HS list: header row with "HS Code" not found')
    head = [str(v).strip().lower() if v is not None else '' for v in rows[hi]]
    data = rows[hi + 1:]
    def first(pred):
        return next((j for j, n in enumerate(head) if pred(n)), None)
    c_in = first(lambda n: n == 'input hs code') if 'input hs code' in head else first(lambda n: n == 'hs code')
    c_mt = first(lambda n: 'matched' in n)
    c_uom = first(lambda n: n == 'hs uom')
    if c_uom is None:
        c_uom = first(lambda n: ('uom' in n or n in ('unit', 'hs unit')) and 'casc' not in n)
    c_desc = first(lambda n: n in ('hs description', 'description') or 'hs desc' in n)
    c_duty = first(lambda n: 'duty item' in n)
    # Import / Export columns: those holding Control / Non Control = control, others = CA agency
    def vals(j):
        return {str(r[j]).strip().upper() for r in data[:3000] if j < len(r) and r[j] not in (None, '')}
    ctl, ca = {}, {}
    for j, n in enumerate(head):
        for side in ('import', 'export'):
            if n == side or n.startswith(side + ' ') and ('control' in n or 'ca' in n):
                v = vals(j)
                if v and v <= {'CONTROL', 'NON CONTROL', 'NON-CONTROL', 'CONTROLLED', 'Y', 'N', 'YES', 'NO', '-'}:
                    ctl.setdefault(side, j)
                elif 'ca agency' not in n or side not in ca:
                    ca.setdefault(side, j)
    c_gen = first(lambda n: 'control' in n and 'rpn' not in n)       # simple list: one "Control" column
    if c_gen is not None and not ctl:
        ctl['import'] = ctl['export'] = c_gen
    rpn = set()
    for r in data:                                                     # RPN 6-digit check table
        for j, v in enumerate(r):
            if isinstance(v, str) and v.strip().upper() == 'CHECK RPN CONTROL' and j > 0 and r[j - 1]:
                rpn.add(str(r[j - 1]).strip()[:6])
    if not rpn:                                                        # RPN table without the label: a column of
        ncol = max(len(r) for r in data) if data else 0                # only 6-digit codes, right of the main table
        start = max([c for c in (c_duty, c_desc, c_uom) if c is not None] + [0]) + 1
        for j in range(start, ncol):
            v = [str(r[j]).strip() for r in data if j < len(r) and r[j] not in (None, '')]
            if 10 <= len(v) <= 2000 and all(re.fullmatch(r'\d{6}', x) for x in v):
                rpn.update(v)
                break
    out = {}
    g = lambda r, j: (str(r[j]).strip() if j is not None and j < len(r) and r[j] not in (None, '') else '')
    for r in data:
        if c_in is None or c_in >= len(r) or r[c_in] in (None, ''):
            continue
        k = h8(r[c_in])
        if not k.isdigit():
            continue
        mt = h8(r[c_mt]) if c_mt is not None and g(r, c_mt) else k
        rec = [mt[:8], g(r, c_uom).upper(), g(r, c_desc), g(r, ctl.get('import')), g(r, ctl.get('export')),
               g(r, ca.get('import')), g(r, ca.get('export')), g(r, c_duty)]
        for key in {k, k[:8]}:
            out.setdefault(key, rec)
    cols = dict(hs=head[c_in] if c_in is not None else '', uom=head[c_uom] if c_uom is not None else '',
                control_import=('col %s' % num2col(ctl['import'] + 1)) if 'import' in ctl else '',
                control_export=('col %s' % num2col(ctl['export'] + 1)) if 'export' in ctl else '',
                ca_import=('col %s' % num2col(ca['import'] + 1)) if 'import' in ca else '')
    pc = product_codes_from_excel(path) if path.lower().endswith(('.xlsx', '.xlsm')) else {}
    return dict(hs=out, rpn=sorted(rpn), columns=cols, source=os.path.basename(path), count=len(out), pcodes=pc,
                pcode_count=sum(len(v) for v in pc.values()))


def product_codes_from_excel(path):
    """Sheet 'PRODUCT CODE' (Singapore Customs HSCA product code master): {hs8: [[code, description, uom, import CA,
    export CA], ...]}.  Uses the master block 'Key (HS|Code) | Product Code | Product Description | Product UOM |
    HS Code | ... | Import CA | Export CA'."""
    import openpyxl
    out = collections.OrderedDict()
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    except Exception:
        return out
    try:
        for ws in wb.worksheets[1:]:
            if 'PRODUCT' not in ws.title.upper():
                continue
            cols = None
            for r in ws.iter_rows(values_only=True):
                if cols is None:
                    head = [str(v).strip().lower() if v is not None else '' for v in r]
                    if 'product code' in head and 'hs code' in head:
                        k0 = next((j for j, n in enumerate(head) if n.startswith('key')), None)
                        base = k0 if k0 is not None else 0
                        f = lambda name: next((j for j in range(base, len(head)) if head[j] == name), None)
                        cols = dict(code=f('product code'), desc=f('product description'), uom=f('product uom'),
                                    hs=f('hs code'), imp=f('import ca'), exp=f('export ca'))
                    continue
                g = lambda k: (str(r[cols[k]]).strip() if cols.get(k) is not None and cols[k] < len(r)
                               and r[cols[k]] not in (None, '') and not str(r[cols[k]]).startswith('=') else '')
                hs, code = h8(g('hs'))[:8], g('code').upper()
                if hs.isdigit() and len(hs) == 8 and code:
                    out.setdefault(hs, []).append([code, g('desc'), g('uom').upper(), g('imp'), g('exp')])
    finally:
        wb.close()
    return out


def save_hs_data(xlsx, json_path):
    d = hs_list_from_excel(xlsx)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(d, f, separators=(',', ':'))
    return d


def load_hs_list(path):
    """Built-in HS data (.json made from the UPDATE HS CODE Excel) or the Excel itself.
    Returns {hs: (matched, uom, desc, info)}; info = dict(ctl_imp, ctl_exp, ca_imp, ca_exp, duty)."""
    if not path or not os.path.exists(path):
        return {}
    if path.lower().endswith('.json'):
        with open(path, encoding='utf-8') as f:
            d = json.load(f)
    else:
        d = hs_list_from_excel(path)
    HS_LIST_INFO.clear()
    HS_LIST_INFO.update(columns=d.get('columns', {}), rpn=set(d.get('rpn', [])), source=d.get('source', ''),
                        count=d.get('count', 0),
                        control_column=d.get('columns', {}).get('control_import') or d.get('columns', {}).get('control_export') or '',
                        uom_column=d.get('columns', {}).get('uom', ''))
    PRODUCT_CODES.clear()
    PRODUCT_CODES.update(d.get('pcodes') or {})
    HS_LIST_INFO['pcode_count'] = d.get('pcode_count', 0)
    out = {}
    for k, r in d['hs'].items():
        out[k] = (r[0], r[1], r[2], dict(ctl_imp=r[3], ctl_exp=r[4], ca_imp=r[5], ca_exp=r[6], duty=r[7]))
    return out


def product_codes(hs):
    """CA product codes of an HS code: [[code, description, uom, import CA, export CA], ...]"""
    return PRODUCT_CODES.get((hs or '')[:8], [])


def hs_candidates(hs, hsl, n=40):
    """HS-list codes that share the first 6 (else 4) digits: [(code, 'UOM - description')]."""
    if not hs or not hsl:
        return []
    for k in (6, 4):
        c = sorted(x for x in hsl if x[:k] == hs[:k])
        if c:
            return [(x, '%s - %s' % (hsl[x][1], hsl[x][2])) for x in c[:n]]
    return []


HS_STOP = {'OTHER', 'OTHERS', 'PARTS', 'THEREOF', 'WITH', 'WITHOUT', 'WHETHER', 'NOT', 'AND', 'FOR', 'THE', 'OF', 'OR',
           'KGM', 'NMB', 'TNE', 'LTR', 'MTR', 'MTK', 'TEN', 'TPR', 'CEN', 'MIL', 'USED', 'TYPE', 'KIND', 'SUCH', 'THAN',
           'MADE', 'BEING', 'HAVING', 'INCLUDING', 'EXCEEDING', 'NOT', 'SUITABLE'}


def _hs_words(t):
    return {w.rstrip('S') for w in re.findall(r'[A-Z]{4,}', (t or '').upper()) if w not in HS_STOP}


def hs_auto_map(hs, hsl, desc=''):
    """Invoice HS not in the UPDATE HS CODE list -> Singapore 8-digit code by the HS code itself.
    Only the first 6 digits are international (digits 7-8 differ per country, e.g. EU vs ASEAN), so the
    Singapore code is chosen inside the same 6-digit subheading:
    1. only one code in the subheading                                   -> that code
    2. several: the item description clearly matches one HS description  -> that code
    3. else the subheading's "OTHER" code (description starting OTHER, highest code; else highest code ending 9)
    Returns (code, how) or ('', '')."""
    hs = (hs or '')[:8]
    if not hs or not hsl or len(hs) < 6:
        return '', ''
    idx = _HS6.get(id(hsl))
    if idx is None:
        idx = {}
        for c in hsl:
            if len(c) == 8 and c.isdigit():
                idx.setdefault(c[:6], []).append(c)
        for v in idx.values():
            v.sort()
        _HS6.clear()
        _HS6[id(hsl)] = idx
    grp = idx.get(hs[:6], [])
    if not grp:
        return '', ''
    if len(grp) == 1:
        return grp[0], 'only code under %s' % hs[:6]
    dw = _hs_words(desc)
    if dw:
        sc = sorted(((len(dw & _hs_words(hsl[c][2])), c) for c in grp), reverse=True)
        if sc[0][0] >= 1 and sc[0][0] > sc[1][0]:
            return sc[0][1], 'description match under %s' % hs[:6]
    oth = [c for c in grp if re.match(r'\s*OTHER', (hsl[c][2] or '').upper())]
    nines = [c for c in grp if c.endswith('9')]
    pick = (oth or nines or [grp[-1]])[-1]
    return pick, '"other" code under %s' % hs[:6]


_HS6 = {}
AIRLINES = {  # IATA flight prefix -> airline (In Transport Identifier for air permits)
    'SQ': 'SINGAPORE AIRLINES', 'CX': 'CATHAY PACIFIC AIRWAYS', 'LH': 'LUFTHANSA', 'LX': 'SWISS', 'EK': 'EMIRATES',
    'QR': 'QATAR AIRWAYS', 'TK': 'TURKISH AIRLINES', 'CI': 'CHINA AIRLINES', 'BR': 'EVA AIR', 'KE': 'KOREAN AIR',
    'OZ': 'ASIANA AIRLINES', 'JL': 'JAPAN AIRLINES', 'NH': 'ALL NIPPON AIRWAYS', 'MU': 'CHINA EASTERN AIRLINES',
    'CA': 'AIR CHINA', 'CZ': 'CHINA SOUTHERN AIRLINES', 'CK': 'CHINA CARGO AIRLINES', 'FX': 'FEDEX EXPRESS',
    '5X': 'UPS AIRLINES', 'QY': 'EUROPEAN AIR TRANSPORT LEIPZIG', '3S': 'AEROLOGIC', 'ET': 'ETHIOPIAN AIRLINES',
    'EY': 'ETIHAD AIRWAYS', 'AF': 'AIR FRANCE', 'KL': 'KLM ROYAL DUTCH AIRLINES', 'BA': 'BRITISH AIRWAYS',
    'AA': 'AMERICAN AIRLINES', 'UA': 'UNITED AIRLINES', 'DL': 'DELTA AIR LINES', 'TG': 'THAI AIRWAYS',
    'MH': 'MALAYSIA AIRLINES', 'GA': 'GARUDA INDONESIA', 'PR': 'PHILIPPINE AIRLINES', 'VN': 'VIETNAM AIRLINES',
    'AI': 'AIR INDIA', 'TR': 'SCOOT', 'QF': 'QANTAS', 'NZ': 'AIR NEW ZEALAND', 'SV': 'SAUDIA', 'LO': 'LOT POLISH AIRLINES',
    'AY': 'FINNAIR', 'OS': 'AUSTRIAN AIRLINES', 'SK': 'SAS SCANDINAVIAN AIRLINES', 'IB': 'IBERIA', 'AZ': 'ITA AIRWAYS',
    'HX': 'HONG KONG AIRLINES', 'MF': 'XIAMEN AIRLINES', '3U': 'SICHUAN AIRLINES', 'HU': 'HAINAN AIRLINES',
    'O3': 'SF AIRLINES', 'CV': 'CARGOLUX', 'PO': 'POLAR AIR CARGO', 'K4': 'KALITTA AIR', 'NCA': 'NIPPON CARGO AIRLINES',
    'KZ': 'NIPPON CARGO AIRLINES', 'LD': 'AIR HONG KONG', 'RU': 'AIRBRIDGECARGO', 'D7': 'AIRASIA X', 'AK': 'AIRASIA',
}


def render_desc(tmpl, vals):
    """Fill the description template. A ' , ' segment whose fields are all empty is dropped."""
    out = []
    for seg in re.split(r'\s,\s', tmpl):
        keys = re.findall(r'\{(\w+)\}', seg)
        if keys and all(not str(vals.get(k.upper(), '')).strip() for k in keys):
            continue
        try:
            out.append(seg.format(**{k: vals.get(k.upper(), '') for k in keys}))
        except (KeyError, IndexError, ValueError):
            out.append(seg)
    return clean(' , '.join(out))


DEFAULT_REMARKS = 'job no : {JOB}\nINV : ( {INV} )'


def render_remarks(tmpl, vals):
    """Fill the Trader's Remarks template. A line whose fields are all empty is left out
    (e.g. no Job No -> no "job no :" line)."""
    if not tmpl:
        return ''
    try:
        out = []
        for line in tmpl.split('\n'):
            keys = re.findall(r'\{(\w+)\}', line)
            if keys and all(not str(vals.get(k, '')).strip() for k in keys):
                continue
            out.append(line.format(**vals))
        return '\n'.join(out).strip()
    except (KeyError, IndexError, ValueError) as e:
        raise RuntimeError("Trader's Remarks template error: %s (use only %s)" % (
            e, ' '.join('{%s}' % v for v in REMARK_VARS)))


def _put(sh, H, header, row, value, style=None, log=None):
    col = H.get(header.lower())
    if not col:
        if log:
            log('CHECK template has no "%s" column - not filled' % header)
        return
    sh.set('%s%d' % (col, row), value, style)


EXPORT_HEADER_CHECK = ['Exporter UEN', 'Exporter', 'Declaration Type', 'Cargo Pack Type', 'Outward Transport',
                       'Departure Date', 'Out Transport Identifier', 'Out Flight Number', 'Outward MAWB',
                       'Outward Carrier Agent Name', 'Discharge Port Code',
                       'Final Destination Country Code', 'Release Location Code', 'Receipt Location Code',
                       'Total Gross Weight', 'Total Outer Pack Qty']


def _date(text, tpl):
    """dd/mm/yyyy (or dd.mm.yyyy / yyyy-mm-dd) -> Excel date serial + date style; else the text as typed."""
    t = str(text).strip()
    for fmt in ('%d/%m/%Y', '%d.%m.%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y'):
        try:
            d = datetime.datetime.strptime(t, fmt).date()
            return (d - datetime.date(1899, 12, 30)).days, tpl.SDATE
        except ValueError:
            continue
    return t, None


def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|]+', '_', str(s)).strip(' ._')
    return s or 'NO_NAME'



# ================================================================ AGILENT
MASTER_COLS = ['Item No', 'Description', 'HS Code', 'HS UOM', 'COO', 'CASC Product Code 1', 'Brand', 'Remarks']

# invoice UOMs counted as pieces for NMB / TEN / CEN ... HS Qty rules
COUNT_UOMS = COUNT_UOMS | {'PK', 'PKT', 'PAC', 'PACK', 'BX', 'BOX', 'KIT'}
PACK_UOMS = {'PK', 'PKT', 'PAC', 'PACK', 'BX', 'BOX', 'KIT'}

DEFAULT_DESC_TEMPLATE = '{DESC}'          # Item Description = the invoice product description only
DESC_VARS = ('ITEM', 'DESC', 'QTY', 'UOM', 'NW', 'GW', 'HS', 'COO', 'BATCH', 'INVOICE', 'ORDER', 'ECCN')
REMARK_VARS = ('INV', 'HAWB', 'MAWB', 'JOB', 'PO', 'ORDER')
DEFAULT_REMARKS = 'job no : {JOB}\nINV : ( {INV} )'

IMPORT_PERMITS = ('IG', 'ME', 'IT', 'IM')
EXPORT_PERMITS = ('OD', 'OT')
TERM_MAP = {'DAP': 'CIF', 'DDP': 'CIF', 'CPT': 'CFR', 'CIP': 'CIF', 'DPU': 'CIF', 'DAT': 'CIF',
            'FCA': 'FOB', 'EXW': 'EXW', 'FOB': 'FOB', 'CIF': 'CIF', 'CFR': 'CFR', 'FAS': 'FAS'}

# generic header key -> (IMPORT column, EXPORT column) in the dec-upload Header sheet
DIR_COLS = {
    'party': ('Importer', 'Exporter'), 'party_uen': ('Importer Uen', 'Exporter UEN'),
    'mawb': ('Inward MAWB', 'Outward MAWB'),
    'carrier_agent': ('Inward Carrier Agent Name', 'Outward Carrier Agent Name'),
    'carrier_agent_uen': ('Inward Carrier Agent Uen', 'Outward Carrier Agent Uen'),
    'transport': ('Inward Transport', 'Outward Transport'),
    'transport_mode': ('Inward Transport Mode', 'Outward Transport Mode'),
    'date': ('Arrival Date', 'Departure Date'),
    'flight': ('In Flight Number', 'Out Flight Number'),
    'transport_id': ('In Transport Identifier', 'Out Transport Identifier'),
    'item_hawb': ('In HAWB', 'Out HAWB'), 'item_mawb': ('In MAWB/OBL', 'Out MAWB/OBL'),
}
IMPORT_HEADER_CHECK = ['Importer', 'Importer Uen', 'Declaration Type', 'Cargo Pack Type', 'Inward Transport',
                       'Arrival Date', 'In Transport Identifier', 'In Flight Number', 'Inward MAWB',
                       'Inward Carrier Agent Name', 'Loading Port Code', 'Release Location Code',
                       'Receipt Location Code', 'Total Gross Weight', 'Total Outer Pack Qty']
EXPORT_HEADER_CHECK = ['Exporter', 'Exporter UEN', 'Declaration Type', 'Cargo Pack Type', 'Outward Transport',
                       'Departure Date', 'Out Transport Identifier', 'Out Flight Number', 'Outward MAWB',
                       'Outward Carrier Agent Name', 'Discharge Port Code', 'Final Destination Country Code',
                       'Release Location Code', 'Receipt Location Code', 'Total Gross Weight', 'Total Outer Pack Qty']

NOT_CONTROLLED_ECCN = {'EAR99', 'NLR', 'N/A', 'NA', 'N', ''}
ECCN_RE = r'(?:EAR99|NLR|N/?A|\d[A-E]\d{3}(?:\.[A-Za-z0-9.]+)?)'
MAIN_RE = re.compile(r'^(\d{1,4})\s+(\S+)\s+(.*?)\s+(?:(\d{8,10})\s+(?:(\S+)\s+)?)?(\d[\d,]*\.\d+)\s*([A-Z]{1,4})\s+'
                     r'([A-Z]{3})\s+(\d[\d,]*\.\d+)\s+([A-Z]{3})\s+(\d[\d,]*\.\d{2})\s*$')
NUMLINE_RE = re.compile(r'^(\d{8,12})\b\s*(.*)$')
INCOTERMS = r'(EXW|FCA|FAS|FOB|CFR|CIF|CPT|CIP|DAP|DPU|DAT|DDP|DDU)'


# ---------------------------------------------------------------- pdf words helpers
def _page_words(page):
    out = []
    for w in page.extract_words(keep_blank_chars=True, x_tolerance=1.5, y_tolerance=2):
        t = w['text'].strip()
        if t:
            out.append(dict(text=t, x0=w['x0'], x1=w['x1'], top=w['top']))
    return out


def _lab(words, text, start=False):
    T = text.upper()
    for w in words:
        if w['text'].upper() == T:
            return w
    Q = squash(text)                       # OCR: "Invoice No" / "_Carrier Reference No.|" ...
    for w in words:
        if squash(w['text']) == Q:
            return w
    if start:
        for w in words:
            if w['text'].upper().startswith(T) or squash(w['text']).startswith(Q):
                return w
    return None


def _lines(sel, tol=3.5):
    rows = []
    for w in sorted(sel, key=lambda w: (w['top'], w['x0'])):
        if rows and abs(rows[-1][0] - w['top']) <= tol:
            rows[-1][1].append(w)
        else:
            rows.append([w['top'], [w]])
    return [clean(' '.join(x['text'] for x in sorted(r[1], key=lambda w: w['x0']))) for r in rows]


def _region(words, x0, x1, y0, y1):
    return _lines([w for w in words if x0 <= w['x0'] < x1 and y0 < w['top'] < y1])


def _below(words, label, dx=95, dy=26, start=False, pattern=None):
    lw = _lab(words, label, start)
    if not lw:
        return ''
    c = [w for w in words if lw['x0'] - 8 <= w['x0'] <= lw['x0'] + dx and lw['top'] + 4 < w['top'] < lw['top'] + dy]
    if pattern:
        c = [w for w in c if re.fullmatch(pattern, w['text'])]
    c.sort(key=lambda w: (w['top'], w['x0']))
    return c[0]['text'] if c else ''


def _company(lines):
    """Company name from an address block: first line(s) up to the street line."""
    name = []
    for ln in lines:
        if re.match(r'(TAX ID|CONTACT|TEL)\b', ln, re.I):
            break
        if name and (re.match(r'^(\d|NO\.?\s*\d|C/O\b|P\.?O\.? BOX)', ln, re.I) or
                     re.search(r'\b(STREET|AVENUE|ROAD|DRIVE|STR\.|STRASSE|BLVD|LANE|WAY|BUILDING)\b', ln, re.I)):
            break
        name.append(ln)
        if len(name) == 2 or not (ln.endswith(',') or re.search(r'\b(SINGAPORE|TECHNOLOGIES|AGILENT)$', ln, re.I)):
            if len(name) == 1 and re.search(r'(,|SINGAPORE|TECHNOLOGIES)$', ln, re.I):
                continue
            break
    n = ' '.join(name)
    n = re.sub(r',\s*\(', ' (', n)
    n = re.sub(r'\s*\((?:[A-Z0-9_]{2,6})\)\s*$', '', n)      # (SG10), (ATSI) org codes
    return clean(n).strip(' ,')


def _mdy(s):
    """Agilent prints MM-DD-YYYY -> dd/mm/yyyy."""
    m = re.match(r'(\d{2})-(\d{2})-(\d{4})$', s or '')
    if not m:
        return s or ''
    return '%s/%s/%s' % (m.group(2), m.group(1), m.group(3))


def classify(text):
    u = text.upper()
    if 'AGILENT' in u and re.search(r'(IMPORT|EXPORT|COMMERCIAL|PROFORMA|TAX)?\s*INVOICE', u) and \
            re.search(r'PRODUCT\s+NUMBER', u):
        return 'INV'
    if re.search(r'PRODUCT\s+NUMBER.*PRODUCT\s+DESCRIPTION', u) and 'INVOICE' in u:
        return 'INV'
    return ''


# ---------------------------------------------------------------- invoice
def parse_invoice_pages(pages, path, countries=None):
    """pages: list of (text, words) of ONE invoice."""
    C = countries or Countries()
    text = '\n'.join(t for t, _ in pages)
    w1 = pages[0][1]
    d = dict(kind='INV', file=os.path.basename(path), path=path, notes=[], lines=[])
    U = text.upper()
    d['title'] = next((t for t in ('IMPORT INVOICE', 'EXPORT INVOICE', 'COMMERCIAL INVOICE', 'PROFORMA INVOICE',
                                   'TAX INVOICE') if t in U), 'INVOICE')
    m = re.search(r'Invoice Date:\s*(\d{2}-\d{2}-\d{4})', text)
    d['invoice_date'] = _mdy(m.group(1)) if m else ''
    d['invoice_no'] = _below(w1, 'Invoice No.')
    d['invoice_type'] = _below(w1, 'Invoice Type')
    if not re.match(r'^[A-Z0-9-]{4,}$', d['invoice_no'] or ''):
        m = re.search(r'Invoice No\.?\s*Invoice Type.*?\n(?:.*\n){0,3}?.*?\b(\d{6,10})\s+([A-Z]{3,5})\s*$', text, re.M)
        d['invoice_no'] = m.group(1) if m else ''
    d['order_no'] = _below(w1, 'Order No.')
    if d['order_no'] and not re.search(r'\d', d['order_no']):
        d['order_no'] = ''
    d['order_date'] = _mdy(_below(w1, 'Order Date', pattern=r'\d{2}-\d{2}-\d{4}'))
    po = _below(w1, 'Customer P.O. No.')
    d['customer_po'] = po if po and not re.match(r'\d{2}-\d{2}-\d{4}$', po) else ''
    inc = _below(w1, 'Incoterms')
    m = re.match(INCOTERMS, inc or '')
    d['incoterm'] = m.group(1) if m else ''
    if not d['incoterm']:
        m = re.search(r'Incoterms.*\n.*?\b' + INCOTERMS + r'\b', text)
        d['incoterm'] = m.group(1) if m else ''
    d['payment_terms'] = _below(w1, 'Payment Terms')
    d['hawb'] = _below(w1, 'Carrier Reference No.', pattern=r'[A-Z0-9][A-Z0-9-]{5,}')
    d['ship_date'] = _mdy(_below(w1, 'Ship Date', pattern=r'\d{2}-\d{2}-\d{4}'))
    d['port_dest'] = _below(w1, 'Port of Destination', dx=110)
    if re.match(r'\d{2}-\d{2}-\d{4}$', d['port_dest']):
        d['port_dest'] = ''
    d['gts_ref'] = _below(w1, 'GTS Ref No', pattern=r'[A-Z0-9-]{4,}')
    # address blocks
    L = {k: _lab(w1, k) for k in ('Ship From', 'Invoice From', 'Invoice No.', 'Invoice To', 'Intermediate Ship To',
                                  'Ship To', 'Notify Party', 'Carrier', 'Carrier Reference No.', 'Line No.', 'Contact')}
    # OCR can split "Intermediate Ship To": the left-most "Ship To" below "Invoice To" is the Ship To block
    st = [w for w in w1 if squash(w['text']) == 'SHIPTO']
    if st:
        it_ = L.get('Invoice To')
        below = [w for w in st if not it_ or w['top'] > it_['top'] + 30] or st
        L['Ship To'] = min(below, key=lambda w: w['x0'])
        if not L.get('Intermediate Ship To'):
            L['Intermediate Ship To'] = next((w for w in w1 if squash(w['text']).startswith('INTERMEDIATE')), None) or \
                next((w for w in st if w is not L['Ship To'] and w['x0'] > 150), None)
    if not L.get('Carrier') and L.get('Carrier Reference No.'):
        cr = L['Carrier Reference No.']
        L['Carrier'] = dict(text='Carrier', x0=max(cr['x0'] - 75, 20), x1=cr['x0'] - 5, top=cr['top'])
    def blk(lab, right, bottom, extra_left=6):
        a, b, c = L.get(lab), L.get(right), L.get(bottom)
        if not a:
            return []
        x1 = b['x0'] - 3 if b else a['x0'] + 230
        y1 = c['top'] - 2 if c else a['top'] + 80
        return [ln for ln in _region(w1, a['x0'] - extra_left, x1, a['top'] + 4, y1)]
    d['ship_from_block'] = blk('Ship From', 'Invoice From', 'Invoice To')
    d['invoice_from_block'] = blk('Invoice From', 'Invoice No.', 'Intermediate Ship To')
    d['invoice_to_block'] = blk('Invoice To', 'Intermediate Ship To', 'Ship To')
    d['ship_to_block'] = blk('Ship To', 'Notify Party', 'Carrier' if L.get('Carrier') else 'Line No.')
    d['carrier_block'] = blk('Carrier', 'Carrier Reference No.', 'Line No.', extra_left=20)
    for k in ('ship_from_block', 'invoice_from_block', 'invoice_to_block', 'ship_to_block'):
        d[k] = [ln for ln in d[k] if not re.match(r'(Tax ID|Contact|Tel)\b', ln, re.I)]
    d['carrier'] = clean(' '.join(ln for ln in d['carrier_block'] if not re.fullmatch(r'(YES|NO|YES NO)', ln, re.I)))
    if not re.search(r'[A-Z]{3}', d['carrier']):
        m = re.search(r'\b(KINTETSU WORLD\s*EXPRESS|DHL[A-Z ]*|FEDEX|UPS|EXPEDITORS|DB SCHENKER|SCHENKER|KUEHNE\s*\+?\s*NAGEL|'
                      r'NIPPON EXPRESS|YUSEN|HELLMANN|CEVA|DSV|AGILITY|BOLLORE|GEODIS)\b', pages[0][0], re.I)
        d['carrier'] = clean(m.group(1)).upper() if m else ''
    d['ship_from'] = _company(d['ship_from_block'])
    d['invoice_from'] = _company(d['invoice_from_block'])
    d['invoice_to'] = _company(d['invoice_to_block'])
    d['ship_to'] = _company(d['ship_to_block'])
    m = re.search(r'\bRN\s*:\s*(\d{8,9}[A-Z]|[STR]\d{2}[A-Z]{2}\d{4}[A-Z])\b', text)
    d['uen'] = m.group(1) if m else ''
    taxes = re.findall(r'Tax ID(?: No\.)?\s+([A-Z0-9][A-Z0-9-]{5,})', pages[0][0])
    d['tax_ids'] = list(dict.fromkeys(taxes))
    d['ship_to_country'] = C.find_in(' '.join(d['ship_to_block']), exclude=()) or ''
    d['ship_from_country'] = C.find_in(' '.join(d['ship_from_block']), exclude=()) or ''
    if not d['ship_from_country'] and any(re.match(r'D-\d{5}\b', ln) for ln in d['ship_from_block']):
        d['ship_from_country'] = 'DE'
    if 'IMPORT' in d['title']:
        d['direction'] = 'IMPORT'
    elif 'EXPORT' in d['title']:
        d['direction'] = 'EXPORT'
    else:
        d['direction'] = 'EXPORT' if d['ship_to_country'] and d['ship_to_country'] != 'SG' else 'IMPORT'
    # totals
    m = re.search(r'Total Number of Boxes\s*(\d+)', text, re.I)
    d['boxes'] = int(m.group(1)) if m else None
    m = re.search(r'Total Gross Weight in K(?:gs|G)\.?\s*(' + NUM + ')', text, re.I) or \
        re.search(r'(?:Total\s+)?Gross\s*W(?:eigh)?t\.?[^0-9\n]{0,25}?(' + NUM + r')\s*(?:KGS?|KGM)?\b', text, re.I)
    d['total_gw'] = fnum(m.group(1)) if m else None
    if d['total_gw'] == 0:
        d['total_gw'] = None
    m = re.search(r'Total Net Weight in K(?:gs|G)\.?\s*(' + NUM + ')', text, re.I) or \
        re.search(r'(?:Total\s+)?Net\s*W(?:eigh)?t\.?[^0-9\n]{0,25}?(' + NUM + r')\s*(?:KGS?|KGM)?\b', text, re.I)
    d['total_nw'] = fnum(m.group(1)) if m else None
    if d['total_nw'] == 0:
        d['total_nw'] = None
    m = re.search(r'Total Amount:?\s*([A-Z]{3})\s*(' + NUM + ')', text)
    d['currency'] = m.group(1) if m else ''
    d['total'] = fnum(m.group(2)) if m else None
    # ---- item lines
    cur, stage = None, 0
    for t, _ in pages:
        for raw in t.splitlines():
            ln = raw.strip()
            m = MAIN_RE.match(ln)
            if m:
                cur = dict(line=m.group(1), item=m.group(2), desc=clean(m.group(3)), hs=m.group(4) or '',
                           license=m.group(5) or '', qty=fnum(m.group(6)), uom=m.group(7),
                           currency=m.group(8), price=fnum(m.group(9)), amount=fnum(m.group(11)),
                           order='', delivery='', eccn='', coo='', pl='', lots=[], boxes=[])
                d['lines'].append(cur)
                if not d['currency']:
                    d['currency'] = cur['currency']
                stage = 1
                continue
            if cur is None:
                continue
            if re.match(r'(Total Number of Boxes|Special Instructions|This is a Computer|IMPORT INVOICE|EXPORT INVOICE'
                        r'|Invoice From|Line No\.)', ln, re.I):
                if re.match(r'(Total Number|Special Instructions)', ln, re.I):
                    stage = 9
                continue
            mc = re.search(r'C[Oo0][Oo0]\s*:\s*([A-Z]{2})\b', ln)
            if mc:
                cur['coo'] = mc.group(1)
            mn = NUMLINE_RE.match(ln)
            if mn and stage in (1, 2) and len(mn.group(1)) >= 9:
                rest = re.sub(r'C[Oo0][Oo0]\s*:\s*[A-Z]{2}\b', '', mn.group(2)).strip()
                me = re.search(r'\s*\b(' + ECCN_RE + r')(?:\s+([YN]|YES|NO))?\s*$', rest)
                if me:
                    cur['eccn'] = cur['eccn'] or me.group(1)
                    rest = rest[:me.start()].strip()
                if rest:
                    cur['desc'] = clean(cur['desc'] + ' ' + rest)
                if stage == 1:
                    cur['order'] = mn.group(1)
                    stage = 2
                else:
                    cur['delivery'] = mn.group(1)
                    stage = 3
                continue
            mp = re.match(r'PL\s*:\s*(\S+)', ln)
            if mp:
                cur['pl'] = mp.group(1)
                stage = 4
                continue
            ml = re.search(r'Lot No\s*:\s*(\S+)', ln, re.I)
            if ml:
                cur['lots'].append(ml.group(1))
            mb = re.match(r'^(\d{8})(?:\s|$)', ln)
            if mb and stage >= 3:
                cur['boxes'].append(mb.group(1))
    if not d['lines']:
        d['notes'].append('no item lines could be read')
    return d


# ---------------------------------------------------------------- air waybill (HAWB / MAWB)
AIRPORTS = {  # IATA -> (UN/LOCODE, name)
    'LEJ': ('DELEJ', 'LEIPZIG'), 'STR': ('DESTR', 'STUTTGART'), 'FRA': ('DEFRA', 'FRANKFURT'),
    'MUC': ('DEMUC', 'MUNICH'), 'CGN': ('DECGN', 'COLOGNE'), 'HAM': ('DEHAM', 'HAMBURG'), 'DUS': ('DEDUS', 'DUSSELDORF'),
    'HHN': ('DEHHN', 'HAHN'), 'AMS': ('NLAMS', 'AMSTERDAM'), 'CDG': ('FRCDG', 'PARIS'), 'LHR': ('GBLHR', 'LONDON HEATHROW'),
    'BRU': ('BEBRU', 'BRUSSELS'), 'LGG': ('BELGG', 'LIEGE'), 'ZRH': ('CHZRH', 'ZURICH'), 'BSL': ('CHBSL', 'BASEL'),
    'VIE': ('ATVIE', 'VIENNA'), 'CPH': ('DKCPH', 'COPENHAGEN'), 'LUX': ('LULUX', 'LUXEMBOURG'), 'MXP': ('ITMXP', 'MILAN'),
    'MAD': ('ESMAD', 'MADRID'), 'BCN': ('ESBCN', 'BARCELONA'), 'DUB': ('IEDUB', 'DUBLIN'), 'IST': ('TRIST', 'ISTANBUL'),
    'DXB': ('AEDXB', 'DUBAI'), 'DOH': ('QADOH', 'DOHA'), 'HKG': ('HKHKG', 'HONG KONG'), 'PVG': ('CNPVG', 'SHANGHAI PUDONG'),
    'PEK': ('CNPEK', 'BEIJING'), 'HGH': ('CNHGH', 'HANGZHOU'), 'SZX': ('CNSZX', 'SHENZHEN'), 'CTU': ('CNCTU', 'CHENGDU'),
    'XMN': ('CNXMN', 'XIAMEN'), 'TSN': ('CNTSN', 'TIANJIN'), 'NKG': ('CNNKG', 'NANJING'), 'MNL': ('PHMNL', 'MANILA'),
    'CGK': ('IDCGK', 'JAKARTA'), 'SGN': ('VNSGN', 'HO CHI MINH CITY'), 'HAN': ('VNHAN', 'HANOI'),
    'MEL': ('AUMEL', 'MELBOURNE'), 'AKL': ('NZAKL', 'AUCKLAND'), 'MAA': ('INMAA', 'CHENNAI'), 'CAN': ('CNCAN', 'GUANGZHOU'), 'TPE': ('TWTPE', 'TAIPEI'), 'ICN': ('KRICN', 'INCHEON'),
    'NRT': ('JPNRT', 'TOKYO NARITA'), 'KIX': ('JPKIX', 'OSAKA KANSAI'), 'BKK': ('THBKK', 'BANGKOK'),
    'KUL': ('MYKUL', 'KUALA LUMPUR'), 'PEN': ('MYPEN', 'PENANG'), 'SIN': ('SGSIN', 'SINGAPORE'),
    'LAX': ('USLAX', 'LOS ANGELES'), 'SFO': ('USSFO', 'SAN FRANCISCO'), 'ORD': ('USORD', 'CHICAGO'),
    'JFK': ('USJFK', 'NEW YORK'), 'BOS': ('USBOS', 'BOSTON'), 'CVG': ('USCVG', 'CINCINNATI'), 'MEM': ('USMEM', 'MEMPHIS'),
    'ANC': ('USANC', 'ANCHORAGE'), 'IAH': ('USIAH', 'HOUSTON'), 'DFW': ('USDFW', 'DALLAS'), 'ATL': ('USATL', 'ATLANTA'),
    'BOM': ('INBOM', 'MUMBAI'), 'DEL': ('INDEL', 'DELHI'), 'BLR': ('INBLR', 'BANGALORE'), 'SYD': ('AUSYD', 'SYDNEY'),
}
CITY_IATA = {v[1]: k for k, v in AIRPORTS.items()}
CITY_IATA.update({'SHANGHAI': 'PVG', 'PUDONG': 'PVG', 'FRANKFURT AM MAIN': 'FRA', 'TOKYO': 'NRT', 'OSAKA': 'KIX',
                  'SEOUL': 'ICN', 'LONDON': 'LHR', 'PARIS': 'CDG', 'NEW YORK': 'JFK', 'MILANO': 'MXP', 'HONGKONG': 'HKG'})


def classify_awb(text):
    u = text.upper()
    if re.search(r'AIR\s*WAY\s*BILL|\bHAWB\b|\bMAWB\b', u) and re.search(r"SHIPPER|CONSIGNEE|AIRPORT OF DEPARTURE", u):
        return True
    return False


def parse_awb(path, pages):
    """House / master air waybill -> cargo details.  pages: [dict(text, words, chars, ocr)]"""
    pg = pages[0]
    text = '\n'.join(p['text'] for p in pages)
    t1 = pg['text']
    words = pg['words']
    d = dict(kind='AWB', file=os.path.basename(path), path=path, notes=[])
    m = re.search(r'MAWB\s*(?:No\.?|Number|#)?\s*:?\s*(\d{3})\s*-?\s*(\d{4})\s*(\d{4})', text, re.I)
    d['mawb'] = ''.join(m.groups()) if m else ''
    m = re.search(r'Customer Air Waybill No\.?\s*(\S+)', text) or re.match(r'\s*([A-Z0-9]{8,})\s+MAWB', t1) or \
        re.search(r'H\.?A\.?W\.?B\.?\s*(?:No\.?|Number|#)?\s*:?\s*([A-Z0-9]{5,})', text, re.I)
    d['hawb'] = m.group(1) if m else ''
    if not d['hawb'] and pg.get('ocr'):
        m = re.search(r'(\d{10,12})', os.path.basename(path))
        if m:
            d['hawb'] = m.group(1)
            d['notes'].append('HAWB %s taken from the file name (not readable on the scan)' % d['hawb'])
    m = re.search(r'Issued By\s+(.+?)\s*$', t1, re.M)
    d['issued_by'] = clean(m.group(1)) if m else ''
    if not d['issued_by'] or not re.search(r'[A-Z]{3}', d['issued_by']):
        m = re.search(r'\b(DHL GLOBAL FORWARDING[^\n]*?(?:LTD\.?|LIMITED|GMBH)|KINTETSU WORLD EXPRESS[^\n]*?(?:GMBH|LTD\.?|INC\.?)|'
                      r'(?:EXPEDITORS|SCHENKER|KUEHNE\s*\+\s*NAGEL|NIPPON EXPRESS|CEVA|DSV|YUSEN|HELLMANN)[^\n]{0,40}?(?:LTD\.?|GMBH|INC\.?|A/S|AG))',
                      text, re.I)
        d['issued_by'] = clean(m.group(1)) if m else d['issued_by']
    m = re.search(r'^\s*(\d+)\s+(\d[\d,]*\.?\d*)\s*(K|L)\b\s*(?:[A-Z]\s+)?(\d[\d,]*\.?\d*)?', t1, re.M)
    if not m and pg.get('ocr'):
        m = re.search(r'Pieces.*\n(?:.*\n)?\s*(\d{1,4})\s+(\d[\d,]*\.\d+)\s*(K|L|KG)?\b\s*(\d[\d,]*\.\d+)?', t1)
    d['pieces'] = int(m.group(1)) if m else None
    gw = fnum(m.group(2)) if m else None
    if gw is not None and (m.group(3) or 'K') == 'L':
        gw = round(gw * 0.45359237, 3)
    d['gw'] = gw
    d['chargeable'] = fnum(m.group(4)) if m and m.group(4) else None
    lab = _lab(words, 'Airport of Departure', start=True)
    d['departure'] = ''
    if lab:
        c = [w for w in words if w['x0'] < lab['x0'] + 150 and lab['top'] + 4 < w['top'] < lab['top'] + 16]
        d['departure'] = clean(' '.join(w['text'] for w in sorted(c, key=lambda w: w['x0']))).upper()
    # routing  "LEJ QY SIN 3S EUR ..."
    legs = []
    lab = _lab(words, 'To')
    if lab:
        row = sorted([w for w in words if lab['top'] + 4 < w['top'] < lab['top'] + 16 and w['x0'] < 280],
                     key=lambda w: w['x0'])
        toks = [w['text'] for w in row]
        i = 0
        while i + 1 < len(toks) and re.fullmatch(r'[A-Z]{3}', toks[i]) and re.fullmatch(r'[A-Z0-9]{2}', toks[i + 1]):
            legs.append((toks[i], toks[i + 1]))
            i += 2
    d['legs'] = legs
    # requested flight / date (characters can overlap the destination text - read them by line height)
    flights = []
    lab = _lab(words, 'Requested Flight/Date')
    if lab:
        band = [c for c in (pg['chars'] or []) if lab['top'] + 6 < c['top'] < lab['top'] + 20 and c['x0'] < lab['x1'] + 60]
        grp = collections.OrderedDict()
        for c in band:
            grp.setdefault(round(c['top'], 1), []).append(c)
        for top, cs in grp.items():
            t = ''.join(c['text'] for c in sorted(cs, key=lambda c: c['x0'])).strip()
            m = re.search(r'([A-Z0-9]{1,8})\s*/\s*(\d{1,2})$', t)
            if m and '[' not in t:
                flights.append((min(c['x0'] for c in cs), m.group(1), int(m.group(2))))
    flights.sort()
    d['flights'] = [(f, dd) for _, f, dd in flights]
    if not d['flights'] and pg.get('ocr'):         # scanned: "... RFS/5   0530/6" on the line after the label
        m = re.search(r'Requested Flight.*\n(.*)', t1)
        if m:
            d['flights'] = [(f, int(dd)) for f, dd in re.findall(r'\b([A-Z0-9]{2,8})\s*/\s*(\d{1,2})\b', m.group(1))]
    if not legs and pg.get('ocr'):                 # scanned: read the routing box again, cropped and sharpened
        row = ''
        la, lb = _lab(words, 'First Carrier', start=True) or _lab(words, 'To By First Carrier', start=True), \
            _lab(words, 'Airport Destination', start=True)
        if pg.get('img') is not None and la:
            try:
                import ocr_runner
                from PIL import ImageFilter, ImageOps
                kk = 1.0 / pg['k']
                y0 = la['top'] - 6
                y1 = (lb['top'] + 22) if lb and lb['top'] > y0 else y0 + 55
                crop = pg['img'].crop((int(15 * kk), int(y0 * kk), int(320 * kk), int(y1 * kk)))
                crop = ImageOps.autocontrast(crop.filter(ImageFilter.SHARPEN))
                row = ocr_runner.image_to_string(crop, 6)
            except Exception:
                row = ''
        if not row:
            m = re.search(r'First Carrier.*\n(?:\s*\n)*(.*)\n', t1)
            row = m.group(1) if m else ''
        lines_ = [l for l in row.splitlines() if re.search(r'\b[A-Z]{3}\b', l.upper()) and 'CARRIER' not in l.upper()]
        if lines_:
            u = re.sub(r'[|\\/_]', ' ', lines_[0].upper())
            legs = [(a, b) for a, b in re.findall(r'\b([A-Z]{3})\s+([A-Z0-9]{2})\b', u) if a not in ('EUR', 'USD', 'SGD')]
            if not legs:
                legs = [(a, '') for a in re.findall(r'\b([A-Z]{3})\b', u) if a in AIRPORTS]
            d['legs'] = legs
        fl = re.search(r'Requested Flight.*\n(?:\s*\n)*(.*)', row)
        if fl and not d['flights']:
            d['flights'] = [(f, int(dd)) for f, dd in re.findall(r'\b([A-Z0-9]{2,8})\s*/\s*(\d{1,2})\b', fl.group(1))]
    if not d['flights']:                           # "CK0287 / 23 / SIN" (DHL and other house AWBs)
        fl = re.findall(r'\b([A-Z0-9]{2}\s?\d{2,4}[A-Z]?)\s*/\s*(\d{1,2})\s*/\s*([A-Z]{3})\b', text)
        if fl:
            d['flights'] = [(f.replace(' ', ''), int(dd)) for f, dd, via in fl]
            d['flight_dest'] = fl[-1][2]
    if not legs and not d.get('departure'):        # "Origin Station ... SHANGHAI PUDONG INTERNATIONAL APT"
        m = re.search(r'Origin\s+Station.*\n((?:.*\n){1,3})', text, re.I)
        blk = (m.group(1) if m else '').upper()
        for city in sorted(CITY_IATA, key=len, reverse=True):
            if city in blk and city != 'SINGAPORE':
                d['departure'] = city
                break
    if any(p.get('ocr') for p in pages):
        d['notes'].append('scanned document read by OCR - check MAWB / flight / pieces / weight')
        if not d['flights']:
            d['notes'].append('flight number / date not readable on the scan - type them on tab 3')
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})\s+\d{2}:\d{2}', text) or re.search(r'Executed on.*?(\d{2})[./](\d{2})[./](\d{4})', text)
    d['executed'] = '%s/%s/%s' % m.groups() if m else ''
    if not d['executed']:                         # "18-Sep-26" / "18 SEP 2026"
        mo = dict(JAN=1, FEB=2, MAR=3, APR=4, MAY=5, JUN=6, JUL=7, AUG=8, SEP=9, OCT=10, NOV=11, DEC=12)
        m = re.search(r'\b(\d{1,2})[-\s]([A-Za-z]{3})[A-Za-z]*[-\s,]+(\d{2}|\d{4})\b', text)
        if m and m.group(2).upper() in mo:
            yy = int(m.group(3)) + (2000 if len(m.group(3)) == 2 else 0)
            d['executed'] = '%02d/%02d/%d' % (int(m.group(1)), mo[m.group(2).upper()], yy)
    lab = _lab(words, 'Airport Destination')
    d['destination'] = ''
    if lab:
        c = [w for w in words if w['x0'] < lab['x0'] + 120 and lab['top'] + 4 < w['top'] < lab['top'] + 16]
        dest = clean(' '.join(w['text'] for w in sorted(c, key=lambda w: w['x0']))).upper()
        d['destination'] = 'SINGAPORE CHANGI' if 'SINGAPORE' in dest else dest
    m = re.search(r'Consignee:\s*(.+)', text)
    d['consignee'] = clean(m.group(1)) if m else ''
    m = re.search(r'Shipper:\s*(.+)', text)
    d['shipper'] = clean(m.group(1)) if m else ''
    m = re.search(r'JOB NO\.?\s*:?\s*(\S+)', text)
    d['fwd_job'] = m.group(1) if m else ''
    m = re.search(r'\bHS:\s*(\d{6,10})', text)
    d['awb_hs'] = m.group(1) if m else ''
    m = re.search(r'Prepaid\s*/\s*([A-Z]{3})\b', text)
    d['awb_terms'] = m.group(1) if m else ''
    # ---- cargo dict in the same keys as the cargo details box
    cg = dict(mode='Air')
    if d['mawb']:
        cg['mawb'] = d['mawb']
    if d['hawb']:
        cg['hawb'] = d['hawb']
    if d['pieces']:
        cg['packages_qty'], cg['packages_unit'] = d['pieces'], 'PCS'
    if d['gw']:
        cg['gw_kg'] = d['gw']
    if legs:
        last_carrier = legs[-1][1]
        if legs[-1][0] not in ('SIN',):            # last "to" is not Singapore -> it is the transit airport
            origin = legs[-1][0]
        else:
            origin = legs[-2][0] if len(legs) > 1 else CITY_IATA.get(d['departure'].split('[')[0].strip(), '')
        if origin:
            code, name = AIRPORTS.get(origin, ('', origin))
            cg['loading_code'], cg['loading_name'] = code or origin, name
            if not code:
                d['notes'].append('loading airport %s not in the airport table - IATA code used' % origin)
        if d['flights']:
            f, day = d['flights'][-1]
            cg['flight'] = f if re.match(r'[A-Z]', f) and not f.startswith('RFS') else last_carrier + f
            if d['executed']:
                dd, mm, yy = (int(x) for x in d['executed'].split('/'))
                mo, yr = (mm, yy) if day >= dd else ((mm % 12) + 1, yy + (mm == 12))
                cg['date'] = '%02d/%02d/%d' % (day, mo, yr)
                cg['date_note'] = 'from AWB requested flight date - confirm the actual arrival date'
    else:
        iata = CITY_IATA.get(d['departure'].split('[')[0].strip(), '') if d['departure'] else ''
        if iata:
            cg['loading_code'], cg['loading_name'] = AIRPORTS[iata]
        if d['flights']:
            f, day = d['flights'][-1]
            cg['flight'] = f
            if d['executed']:
                dd, mm, yy = (int(x) for x in d['executed'].split('/'))
                mo, yr = (mm, yy) if day >= dd else ((mm % 12) + 1, yy + (mm == 12))
                cg['date'] = '%02d/%02d/%d' % (day, mo, yr)
                cg['date_note'] = 'from the AWB flight date - confirm the actual arrival date'
    d['cargo'] = cg
    return d


OCR = dict(enabled=True, cmd='', dpi=300, status='')


def set_tesseract(cmd):
    """Find tesseract (typed path / next to the exe / inside the exe / installed). True when OCR works."""
    import ocr_runner
    ok = ocr_runner.setup(cmd or '')
    st = ocr_runner.STATE
    OCR['cmd'] = st['cmd']
    OCR['status'] = ('%s  (%s)' % (st['version'], st['cmd'])) if ok else ('not available - ' + st['error'])
    OCR['tried'] = st['tried']
    return ok


def ocr_fix(t):
    t = t.replace('\u2014', '-').replace('\u2013', '-').replace('\u2018', "'").replace('\u2019', "'")
    t = re.sub(r'(?<=\d),(?=\d{3}\.\d{2}\b)', ',', t)
    t = re.sub(r'\bCO0\b|\bC00\b', 'COO', t)
    t = re.sub(r'(?m)^[|_\-\u2014 ]+', '', t)
    t = re.sub(r'\bEAR[QSOo09]{2,3}\b', 'EAR99', t)                 # EARQ9 / EARS9 -> EAR99
    t = re.sub(r'\bIll\b', 'III', t)
    t = re.sub(r'\bIl\b', 'II', t)
    t = re.sub(r'\|(?=-?[A-Za-z])', 'I', t)                         # |-AS -> I-AS
    t = t.replace('\u00a7', '5')                                     # §500-1193 -> 5500-1193
    t = re.sub(r'(?<=\d)m!', 'ml', t)
    return t


def _ocr_words(img, k):
    """Tesseract sparse-text word boxes -> phrases (words close together on a line), in PDF points."""
    import ocr_runner
    d = ocr_runner.image_to_data(img, 11)
    toks = []
    for i, t in enumerate(d['text']):
        t = (t or '').strip().strip('|_')
        if not t or float(d['conf'][i]) < 20:
            continue
        toks.append(dict(text=t, x0=d['left'][i] * k, x1=(d['left'][i] + d['width'][i]) * k, top=d['top'][i] * k,
                         key=(d['block_num'][i], d['par_num'][i], d['line_num'][i])))
    out = []
    for w in toks:
        if out and out[-1]['key'] == w['key'] and w['x0'] - out[-1]['x1'] < 6.5:
            out[-1]['text'] += ' ' + w['text']
            out[-1]['x1'] = w['x1']
        else:
            out.append(dict(w))
    return [dict(text=w['text'], x0=w['x0'], x1=w['x1'], top=w['top']) for w in out]


def read_pages(path, log=None):
    """[dict(text, words, chars, ocr)] - text layer when present, else OCR (scanned page)."""
    import pdfplumber
    pages, need = [], []
    with pdfplumber.open(path) as pdf:
        for i, p in enumerate(pdf.pages):
            t = p.extract_text() or ''
            if len(re.sub(r'\s', '', t)) >= 25:
                pages.append(dict(text=t, words=_page_words(p), chars=p.chars, ocr=False))
            else:
                pages.append(None)
                need.append(i)
    if need:
        if not OCR['enabled']:
            raise RuntimeError('scanned PDF (no text) - OCR is switched off')
        import pypdfium2 as pdfium
        import ocr_runner
        if not ocr_runner.STATE['cmd'] and not ocr_runner.setup():
            raise RuntimeError('scanned PDF (no text) - Tesseract OCR not available: %s' % ocr_runner.STATE['error'])
        doc = pdfium.PdfDocument(path)
        k = 72.0 / OCR['dpi']
        imgs = {}

        def one(i):
            img = imgs.pop(i)
            text = ocr_fix(ocr_runner.image_to_string(img, 6))
            first = i == 0 or re.search(r'Page\s*1\s*of\s*\d+', text)
            return i, dict(text=text, words=_ocr_words(img, k) if first else [], chars=None, ocr=True,
                           img=img if first else None, k=k)
        todo = list(need)
        if todo and todo[0] == 0:                 # page 1 first: an air waybill needs only page 1
            imgs[0] = doc[0].render(scale=OCR['dpi'] / 72.0).to_pil().convert('L')
            if log:
                log('   OCR %s page 1 / %d ...' % (os.path.basename(path), len(pages)))
            _, pages[0] = one(0)
            todo = todo[1:]
            if classify_awb(pages[0]['text']) and not classify(pages[0]['text']) and todo:
                imgs[todo[0]] = doc[todo[0]].render(scale=OCR['dpi'] / 72.0).to_pil().convert('L')
                _, pages[todo[0]] = one(todo[0])             # page 2: attachment of the waybill or another document?
                todo = todo[1:]
                if not re.search(r'INVOICE|PACKING|Subject:|Product\s*Number', pages[1]['text'], re.I):
                    for i in todo:
                        pages[i] = dict(text='', words=[], chars=None, ocr=True)
                    todo = []
        from concurrent.futures import ThreadPoolExecutor
        workers = max(1, min(4, (os.cpu_count() or 2)))
        for b in range(0, len(todo), workers):
            batch = todo[b:b + workers]
            for i in batch:
                imgs[i] = doc[i].render(scale=OCR['dpi'] / 72.0).to_pil().convert('L')
            if log:
                log('   OCR %s pages %d-%d / %d ...' % (os.path.basename(path), batch[0] + 1, batch[-1] + 1, len(pages)))
            with ThreadPoolExecutor(len(batch)) as ex:
                for i, pg in ex.map(one, batch):
                    pages[i] = pg
        doc.close()
    return pages


# ---------------------------------------------------------------- combined PDFs: split pages by document
def page_kind(text):
    head = text[:600]
    if classify(text):
        return 'INV'
    if re.search(r'^\s*Packing\s*List', text, re.I | re.M) and re.search(r'Packing\s*List', head, re.I):
        return 'PL'
    if classify_awb(text):
        return 'AWB'
    if re.search(r'^From:.*\n(?:.*\n){0,5}Subject:', text, re.M):
        return 'EMAIL'
    return ''


def _awb_key(text):
    m = re.search(r'MAWB[^0-9\n]{0,15}(\d{3})\s*-?\s*(\d{4})\s*(\d{4})', text, re.I)
    return ''.join(m.groups()) if m else ''


def _same_awb(first, text):
    """Next air waybill page = attachment of the same waybill (same / no MAWB shown) or a new waybill?"""
    a, b = _awb_key(first), _awb_key(text)
    return not a or not b or a == b


def segment_pages(pages):
    """[(kind, [pages])] - one PDF can hold air waybill + e-mail + invoice + packing lists (e.g. DHL document packs)."""
    segs = []
    for pg in pages:
        k = page_kind(pg['text'])
        if not k:                                   # continuation page / attachment -> previous document
            if segs:
                segs[-1][1].append(pg)
            else:
                segs.append(['', [pg]])
            continue
        if segs and segs[-1][0] == k and (k != 'AWB' or _same_awb(segs[-1][1][0]['text'], pg['text'])):
            segs[-1][1].append(pg)            # AWB attachment pages (same MAWB / HAWB) stay with their waybill
        else:
            segs.append([k, [pg]])
    return [(k, ps) for k, ps in segs]


# ---------------------------------------------------------------- Agilent packing list
AG_PL_RE = re.compile(r'^\s*(\d+)\s+(\d+)\s+(\d[\d,.]*)\s+([A-Z]{1,4})\s+(\S+)\s+(.+?)\s+(\d[\d.,]*)\s*(KG|LB)S?\s+(\d{6,})\s*$')


def parse_agilent_pl(path, pages):
    """Agilent packing list: one row per box (qty, product, description, box weight, box no)."""
    text = '\n'.join(p['text'] for p in pages)
    d = dict(kind='PL', file=os.path.basename(path), path=path, reader='Agilent packing list', notes=[], invoice_refs=[],
             hawb='', mawb='', total_gw=None, total_nw=None, packages=None, package_unit='BOX', containers=[],
             lines=[], ocr=any(p.get('ocr') for p in pages), shipper='', consignee='', deliveries=[])
    m = re.search(r'Carrier\s*Ref\.?\s*No\.?\s*:?\s*([A-Z0-9]{5,})', text, re.I)
    d['hawb'] = m.group(1) if m else ''
    d['deliveries'] = list(dict.fromkeys(x.lstrip('0') for x in re.findall(r'Delivery\s*No\.?\s*:?\s*(\d{8,})', text, re.I)))
    cur = None
    for raw in text.splitlines():
        ln = raw.strip()
        m = AG_PL_RE.match(ln)
        if m:
            w = fnum(m.group(7))
            if m.group(8).upper().startswith('LB') and w:
                w = round(w * 0.45359237, 3)
            cur = dict(line=m.group(1), item=m.group(5), desc=m.group(6), qty=fnum(m.group(3)), uom=m.group(4),
                       gw=w, nw=None, coo='', boxes=[m.group(9)], lots=[], hs='', amount=None, price=None,
                       currency='', eccn='', order='', delivery='', license='', pl='')
            d['lines'].append(cur)
            continue
        mc = re.match(r'^C[Oo0][Oo0]\s*:\s*([A-Z]{2})\b', ln)
        if mc and cur is not None:
            cur['coo'] = mc.group(1)
    boxes = [int(x) for x in re.findall(r'Total\s*No\.?\s*of\s*Boxes\s*:?\s*(\d+)', text, re.I)]
    wts = [fnum(x) for x in re.findall(r'Total\s*Weight\s*:?\s*(\d[\d.,]*)\s*KG', text, re.I)]
    d['packages'] = sum(boxes) or len({b for l in d['lines'] for b in l['boxes']}) or None
    d['total_gw'] = round(sum(w for w in wts if w), 3) or (round(sum(l['gw'] or 0 for l in d['lines']), 3) or None)
    if not d['lines']:
        d['notes'].append('packing list rows could not be read')
    return d


def pl_weights_to_invoices(invs, pls, issues=None):
    """Packing list box weights (gross) -> invoice lines by product number, split by qty.  Sets l['pl_gw']."""
    pm = collections.OrderedDict()
    for p in pls:
        for l in p['lines']:
            k = squash(l['item'])
            if not k or l.get('gw') is None:
                continue
            m = pm.setdefault(k, dict(qty=0.0, gw=0.0))
            m['qty'] += l['qty'] or 0
            m['gw'] += l['gw']
    if not pm:
        return
    lines = collections.defaultdict(list)
    for inv in invs:
        for l in inv['lines']:
            lines[squash(l['item'])].append(l)
    for k, ls in lines.items():
        m = pm.get(k)
        if not m:
            if issues is not None:
                issues.append(('CHECK', 'Item %s is on the invoice but not on the packing list' % ls[0]['item']))
            continue
        tq = sum(l['qty'] or 0 for l in ls)
        for l in ls:
            l['pl_gw'] = m['gw'] * ((l['qty'] or 0) / tq if tq else 1.0 / len(ls))
        if issues is not None and tq and m['qty'] and abs(tq - m['qty']) > 0.0005:
            issues.append(('MISMATCH', 'Item %s quantity: invoice %s, packing list %s' % (ls[0]['item'], fmt_plain(tq), fmt_plain(m['qty']))))
    if issues is not None:
        for k in pm:
            if k not in lines:
                issues.append(('CHECK', 'Packing list item %s is not on the invoice' % k))


def parse_pdf(path, countries=None, log=None):
    """Returns a LIST of documents. One PDF can hold several documents (air waybill, e-mail, invoices, packing
    lists - e.g. a DHL document pack); every part is read with its own reader. Scanned PDFs are read by OCR."""
    allp = read_pages(path, log)
    if not allp:
        return [dict(kind='', file=os.path.basename(path), path=path, notes=['empty PDF'])]
    import danfoss_festo_docs as festo_docs, danfoss_docs
    fd = danfoss_docs.read_pdf(path, allp, countries)     # Danfoss commercial invoice / packing lists
    if fd:
        return fd
    fd = festo_docs.read_pdf(path, allp, countries)       # B/L, customs instructions (and Festo documents)
    if fd:
        txt = '\n'.join(p['text'] for p in allp)
        for d in fd:
            if d.get('doc_type') == 'Customs instructions':
                danfoss_docs.improve_ci(d, txt)
            elif d.get('doc_type') == 'B/L':
                danfoss_docs.improve_bl(d, txt)
        return fd
    out = []
    for kind, pgs in segment_pages(allp):
        if kind == 'AWB':
            out.append(parse_awb(path, pgs))
        elif kind == 'PL':
            out.append(parse_agilent_pl(path, pgs))
        elif kind == 'EMAIL':
            continue
        else:
            out += _inv_docs(path, pgs, countries)
    return out or [dict(kind='', file=os.path.basename(path), path=path,
                        notes=['not recognised (no Festo invoice / packing list / B/L / customs instructions / air waybill)'])]


def _inv_docs(path, allp, countries=None):
    docs, cur = [], []
    for pg in allp:
        if cur and re.search(r'Page\s*1\s*of\s*\d+', pg['text']):
            docs.append(cur)
            cur = []
        cur.append(pg)
    if cur:
        docs.append(cur)
    out = []
    for pgs in docs:
        text = '\n'.join(p['text'] for p in pgs)
        if classify(text) == 'INV':
            d = parse_invoice_pages([(p['text'], p['words']) for p in pgs], path, countries)
            if any(p['ocr'] for p in pgs):
                d['ocr'] = True
                ub = len({b for l in d['lines'] for b in l['boxes']})
                if ub and d['boxes'] != ub:
                    d['notes_box'] = 'Total Number of Boxes read as %s on the scan - %d Box No. listed, %d used' % (
                        d['boxes'], ub, ub)
                    d['boxes'] = ub
                ls = round(sum(l['amount'] or 0 for l in d['lines']), 2)
                if d['total'] is None or abs((d['total'] or 0) - ls) > 0.01:
                    d['notes_total'] = 'Total Amount %s on the scan - item lines total %s used (check the invoice)' % (
                        'not readable' if d['total'] is None else '{:,.2f} read'.format(d['total']), '{:,.2f}'.format(ls))
                    d['total'] = ls
                    d['currency'] = d['currency'] or next((l['currency'] for l in d['lines'] if l['currency']), '')
                d['notes_ocr'] = 'scanned invoice read by OCR - check item no / HS / qty / amounts'
                if not d['hawb']:
                    m = re.search(r'(\d{10,12})', os.path.basename(path))
                    if m:
                        d['hawb'], d['hawb_from_name'] = m.group(1), True
        else:
            d = dict(kind='', file=os.path.basename(path), path=path,
                     notes=['Not recognised as an Agilent invoice' + (' (scanned - OCR text unclear)'
                                                                       if any(p['ocr'] for p in pgs) else '')])
        out.append(d)
    return out


# ---------------------------------------------------------------- cargo details (pre-alert / AWB / e-mail text)
CARGO_FIELDS = [  # (key, label regex) - "LABEL: value" or "LABEL value"
    ('mawb', r'(?:MAWB|MASTER\s*AWB|MASTER\s*AIR\s*WAYBILL|OBL|OCEAN\s*B/?L|MBL|M\.?B/?L)\s*(?:NO\.?|NUMBER|#)?'),
    ('hawb', r'(?:HAWB|HOUSE\s*AWB|HOUSE\s*AIR\s*WAYBILL|HBL|H\.?B/?L)\s*(?:NO\.?|NUMBER|#)?'),
    ('flight', r'(?:FLIGHT|FLT|FLIGHT\s*NO\.?|VOYAGE|VOY)\s*(?:NO\.?|#)?'),
    ('transport_id', r'(?:VESSEL|VSL|VESSEL\s*NAME|AIRLINE|CARRIER\s*NAME)'),
    ('date', r'(?:ETA|ARRIVAL\s*DATE|DATE\s*OF\s*ARRIVAL|ETD|DEPARTURE\s*DATE|FLIGHT\s*DATE)'),
    ('loading', r'(?:PORT\s*OF\s*LOADING|POL|LOADING\s*PORT|AIRPORT\s*OF\s*DEPARTURE|ORIGIN|FROM)'),
    ('discharge', r'(?:PORT\s*OF\s*DISCHARGE|POD|DISCHARGE\s*PORT|AIRPORT\s*OF\s*DESTINATION|DESTINATION|TO)'),
    ('packages', r'(?:PCS|PIECES|NO\.?\s*OF\s*(?:PCS|PKGS|PACKAGES|PIECES)|PACKAGES|PKGS|CARTONS|CTNS)'),
    ('gw', r'(?:GROSS\s*WEIGHT|GROSS\s*WT|G\.?\s*W\.?|GW|ACTUAL\s*WEIGHT)'),
    ('carrier_agent', r'(?:CARRIER\s*AGENT|INWARD\s*CARRIER\s*AGENT|OUTWARD\s*CARRIER\s*AGENT|CARRIER)'),
    ('permit_type', r'(?:TYPE\s*OF\s*PERMIT|PERMIT\s*TYPE)'),
    ('job_no', r'(?:JOB\s*(?:NO|NUMBER|#)\.?)'),
    ('uen', r'(?:UEN|IMPORTER\s*UEN|EXPORTER\s*UEN)\s*(?:NO\.?|#)?'),
    ('release', r'(?:PLACE\s*OF\s*RELEASE|RELEASE\s*LOCATION)'),
    ('receipt', r'(?:PLACE\s*OF\s*RECEIPT|RECEIPT\s*LOCATION)'),
    ('permit_remarks', r'(?:PERMIT\s*REMARKS|TRADER\'?S\s*REMARKS)'),
]


def parse_cargo(text):
    """Cargo details from a pasted pre-alert / AWB / B/L / e-mail ("LABEL: value" lines)."""
    out = {}
    if not text or not text.strip():
        return out
    for raw in text.splitlines():
        ln = raw.strip()
        if not ln:
            continue
        for key, lab in CARGO_FIELDS:
            m = re.match(r'(%s)\s*[:=\-]\s*(.+)$' % lab, ln, re.I)
            if m and clean(m.group(2)):
                out.setdefault(key, clean(m.group(2)))
                break
    U = text.upper()
    if 'mawb' not in out:
        m = re.search(r'\b(\d{3})[- ]?(\d{4})\s?(\d{4})\b', text)
        if m:
            out['mawb'] = '%s%s%s' % m.groups()
    else:
        m = re.match(r'(\d{3})[\s-]?(\d{4})\s?(\d{4})\b', out['mawb'])
        out['mawb'] = '%s%s%s' % m.groups() if m else out['mawb'].split()[0]
    if 'flight' in out:
        m = re.match(r'([A-Z0-9]{2})\s*-?\s*(\d{1,4}[A-Z]?)\b', out['flight'].upper())
        out['flight'] = m.group(1) + m.group(2) if m else out['flight'].split('/')[0].strip()
    if 'date' in out:
        dt = out['date'].upper().replace('.', '/').strip()
        for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d%b%Y', '%d-%b-%Y', '%d %b %Y', '%d%b%y', '%d-%b-%y',
                    '%d %b %y', '%d/%m/%y'):
            try:
                out['date'] = datetime.datetime.strptime(dt.split()[0] if fmt.count(' ') == 0 else dt[:11], fmt).strftime('%d/%m/%Y')
                break
            except ValueError:
                continue
    if 'hawb' in out:
        out['hawb'] = out['hawb'].split()[0]
    conts = list(dict.fromkeys(re.findall(r'\b([A-Z]{4}\d{7})\b', U)))
    if conts:
        out['containers'] = conts
    if 'packages' in out:
        m = re.search(r'(\d[\d,]*)\s*([A-Z]{2,5})?', out['packages'].upper())
        out['packages_qty'] = int(fnum(m.group(1))) if m else None
        out['packages_unit'] = (m.group(2) or '') if m else ''
    if 'gw' in out:
        out['gw_kg'] = fnum(re.search(NUM, out['gw']).group(0)) if re.search(NUM, out['gw']) else None
    if out.get('permit_type'):
        out['permit_type'] = out['permit_type'].upper()[:2]
    if 'uen' in out:
        m = re.search(r'\b(\d{8,9}[A-Z]|[STR]\d{2}[A-Z]{2}\d{4}[A-Z]|M\d-\d{7}-[A-Z])\b', out['uen'].upper())
        out['uen'] = m.group(1) if m else ''
    for k in ('loading', 'discharge'):
        if k in out:
            v = out[k].upper()
            m = re.match(r'^([A-Z]{3}|[A-Z]{5})\s*[-/(]\s*(.+?)\)?$', v)
            if m:
                out[k + '_code'], out[k + '_name'] = m.group(1), m.group(2)
            elif re.fullmatch(r'[A-Z]{5}', v):
                out[k + '_code'] = v
            elif re.fullmatch(r'[A-Z]{3}', v):
                out[k + '_code'] = v              # IATA airport code - check
            else:
                out[k + '_name'] = v
    if out.get('vessel') or re.search(r'\b(VESSEL|VSL|B/L|OBL|CONTAINER)\b', U):
        out['mode'] = 'Sea'
    elif re.search(r'\b(FLIGHT|FLT|AWB|AIRPORT)\b', U):
        out['mode'] = 'Air'
    return {k: v for k, v in out.items() if v not in (None, '', [])}


# ---------------------------------------------------------------- shipments (one per HAWB)
def build_shipments(docs):
    ships = collections.OrderedDict()
    gi = []
    import danfoss_festo_docs as festo_docs, danfoss_docs
    danfoss_docs.dedupe(docs, gi)
    festo_docs.link_docs(docs, gi)
    awbs = [d for d in docs if d.get('kind') == 'AWB']
    pls = [d for d in docs if d.get('kind') == 'PL']
    for d in docs:
        if d.get('kind') in ('AWB', 'PL'):
            continue
        if d.get('kind') != 'INV':
            gi.append((d['file'], 'CHECK', '; '.join(d.get('notes') or ['not an invoice'])))
            continue
        key = d['hawb'] or ('INV:' + (d['invoice_no'] or d['file']))
        sh = ships.get(key)
        if not sh:
            sh = ships[key] = dict(key=key, hawb=d['hawb'], hawb_src='Invoice' if d['hawb'] else '', invs=[],
                                   issues=[], overrides={}, include=True, direction=d['direction'],
                                   party='', party_src='', party_uen='', party_uen_src='', dest='', dest_src='', awbs=[],
                                   pls=[])
        sh['invs'].append(d)
        if d['direction'] != sh['direction']:
            sh['issues'].append(('MISMATCH', '%s is %s but other invoices of this HAWB are %s'
                                 % (d['invoice_no'], d['direction'], sh['direction'])))
        for n in d['notes']:
            sh['issues'].append(('ERROR', '%s: %s' % (d['file'], n)))
        if d.get('notes_shared'):
            sh['issues'].append(('CHECK', d['notes_shared']))
        if d.get('hawb_from'):
            sh['hawb_src'] = 'from ' + d['hawb_from']
    for a in awbs:
        hit = [sh for sh in ships.values() if a['hawb'] and squash(sh['hawb']) == squash(a['hawb'])]
        if not hit and len(ships) == 1:
            hit = list(ships.values())
            if a['hawb']:
                hit[0]['issues'].append(('MISMATCH', '%s %s: HAWB/HBL %s, invoice %s'
                                         % (a.get('doc_type', 'Air waybill'), a['file'], a['hawb'], hit[0]['hawb'] or '-')))
        if hit:
            hit[0]['awbs'].append(a)
        else:
            gi.append((a['file'], 'CHECK', 'Air waybill HAWB %s does not match any invoice' % (a['hawb'] or '?')))
    for pl in pls:
        dl = set(pl.get('deliveries', []))
        hit = [sh for sh in ships.values() if pl['hawb'] and squash(sh['hawb']) == squash(pl['hawb'])] or \
              [sh for sh in ships.values() if dl & {l['delivery'].lstrip('0') for i in sh['invs'] for l in i['lines']}]
        if not hit and len(ships) == 1:
            hit = list(ships.values())
        if hit:
            hit[0]['pls'].append(pl)
        else:
            gi.append((pl['file'], 'CHECK', 'Packing list (Carrier Ref %s) does not match any invoice' % (pl['hawb'] or '?')))
    for sh in ships.values():
        if any(p.get('supplier') == 'DANFOSS' for p in sh['pls']):
            danfoss_docs.check_pl(sh)
        elif any(p.get('supplier') == 'FESTO' for p in sh['pls']):
            festo_docs.check_pl(sh)
        elif sh['pls']:
            pl_weights_to_invoices(sh['invs'], sh['pls'], sh['issues'])
            pk = sum(p['packages'] or 0 for p in sh['pls'])
            ib = sum(i['boxes'] or 0 for i in sh['invs'])
            if pk and ib and pk != ib:
                sh['issues'].append(('MISMATCH', 'Boxes: invoice %s, packing lists %s' % (ib, pk)))
            pg = sum(p['total_gw'] or 0 for p in sh['pls'])
            ig = sum(i['total_gw'] or 0 for i in sh['invs'])
            if pg and ig and abs(pg - ig) > 0.05:
                sh['issues'].append(('MISMATCH', 'Gross weight: invoice %s KG, packing lists %s KG' % (fmt_thou(ig), fmt_thou(pg))))
        nos = [i['invoice_no'] for i in sh['invs']]
        dup = [n for n, c in collections.Counter(nos).items() if c > 1]
        for n in dup:
            sh['issues'].append(('MISMATCH', 'Invoice %s added more than once (same PDF twice?)' % n))
    return ships, gi


def ship_label(sh):
    return sh['hawb'] or ', '.join(i['invoice_no'] for i in sh['invs'])


def _party_default(sh, ctx):
    i0 = sh['invs'][0]
    if i0.get('party_name') and sh['direction'] == 'IMPORT':
        return i0['party_name'].upper()
    if sh['direction'] == 'IMPORT':
        name = i0['ship_to'] or i0['invoice_to']
    else:
        name = i0['invoice_from'] or i0['ship_from']
    return name.upper()


def resolve(sh, ctx):
    """(Re)build the items of one shipment.
    ctx: master, hs_list, countries, default_uom, blank_hs_not_in_list, desc_template, desc_max, brand,
         cargo{}, party_name, party_uen, uen_book{}, consolidate ('none'|'item'|'hs_coo'), direction_mode"""
    master, hsl = ctx.get('master', {}), ctx.get('hs_list', {})
    issues = list(sh['issues'])
    box = ctx.get('cargo') or {}
    cg = {}
    for a in sh.get('awbs', []):
        for k, v in a['cargo'].items():
            cg.setdefault(k, v)
        for n in a['notes']:
            issues.append(('CHECK', '%s: %s' % (a['file'], n)))
        if a['cargo'].get('date_note'):
            issues.append(('CHECK', 'Arrival Date %s %s' % (a['cargo'].get('date', ''), a['cargo']['date_note'])))
        if a.get('awb_hs'):
            issues.append(('CHECK', 'Air waybill shows one HS %s for the whole shipment - item HS codes from the invoice used'
                           % a['awb_hs']))
        imp = _party_default(sh, ctx)
        if a.get('consignee') and sh['direction'] == 'IMPORT' and squash(a['consignee'])[:20] != squash(imp)[:20]:
            issues.append(('CHECK', 'Air waybill consignee "%s" vs importer "%s"' % (a['consignee'], imp)))
    if (cg.get('mode') == 'Air' or cg.get('flight')) and cg.get('flight') and not cg.get('transport_id'):
        al = AIRLINES.get(re.sub(r'\d.*$', '', cg['flight'])[:3]) or AIRLINES.get(cg['flight'][:2])
        if al:
            cg['transport_id'] = al
            issues.append(('CHECK', 'In Transport Identifier %s from flight %s - check' % (al, cg['flight'])))
    awb_cg = dict(cg)
    cg.update(box)                        # cargo details box (typed / pasted) wins over the air waybill
    sh['cargo'] = cg
    dm = (ctx.get('direction_mode') or 'Auto').upper()
    if dm in ('IMPORT', 'EXPORT'):
        sh['direction'] = dm
    else:
        sh['direction'] = sh['invs'][0]['direction']
    exp = sh['direction'] == 'EXPORT'
    # ---- party (importer for IMPORT, exporter for EXPORT)
    if sh.get('party_src') != 'Entered':
        sh['party'] = (ctx.get('party_name') or '').upper() or _party_default(sh, ctx)
        sh['party_src'] = 'Tab 3' if ctx.get('party_name') else 'Invoice'
    if sh.get('party_uen_src') != 'Entered':
        book = {k.upper(): v for k, v in (ctx.get('uen_book') or {}).items()}
        doc_uen = next((i.get('uen') for i in sh['invs'] if i.get('uen')), '') if sh['direction'] == 'IMPORT' else ''
        uen = ctx.get('party_uen') or cg.get('uen') or doc_uen or book.get(sh['party'].upper(), '')
        sh['party_uen'], sh['party_uen_src'] = uen, ('Tab 3' if ctx.get('party_uen') else 'Cargo details'
                                                      if cg.get('uen') else 'Invoice (RN)' if doc_uen else 'Saved' if uen else '')
    if not sh['party_uen']:
        issues.append(('MISSING', '%s UEN not in the documents - type it on tab 3 (remembered) or tab 2'
                       % ('Exporter' if exp else 'Importer')))
    if exp and sh.get('dest_src') != 'Entered':
        sh['dest'] = sh['invs'][0]['ship_to_country']
        if not sh['dest']:
            issues.append(('MISSING', 'Final destination country not found in Ship To - type it on tab 2'))
    # ---- HAWB
    if sh.get('hawb_src') != 'Entered':
        if not sh['hawb'] and cg.get('hawb'):
            sh['hawb'], sh['hawb_src'] = cg['hawb'], 'Cargo details'
        elif sh['hawb'] and cg.get('hawb') and squash(cg['hawb']) != squash(sh['hawb']):
            issues.append(('MISMATCH', 'HAWB: invoice Carrier Reference No. %s, cargo details %s (invoice used)'
                           % (sh['hawb'], cg['hawb'])))
    # ---- weights / packages
    gw_inv = sum(i['total_gw'] or 0 for i in sh['invs']) or None
    nw_inv = sum(i['total_nw'] or 0 for i in sh['invs']) or None
    boxes = sum(i['boxes'] or 0 for i in sh['invs']) or None
    sh['gw_doc'], sh['boxes_doc'] = gw_inv, boxes
    # gross weight: typed on tab 2 > invoice > air waybill > cargo details box
    if sh.get('gw_override'):
        sh['total_gw'], sh['gw_src'] = sh['gw_override'], 'typed on tab 2'
    elif gw_inv:
        sh['total_gw'], sh['gw_src'] = gw_inv, 'invoice'
    elif awb_cg.get('gw_kg'):
        sh['total_gw'], sh['gw_src'] = awb_cg['gw_kg'], 'air waybill'
    elif cg.get('gw_kg'):
        sh['total_gw'], sh['gw_src'] = cg['gw_kg'], 'cargo details box'
    else:
        sh['total_gw'], sh['gw_src'] = None, ''
    sh['packages'] = sh.get('pkg_override') or boxes or cg.get('packages_qty')
    for src, c in (('air waybill', awb_cg), ('cargo details', box if len(ctx.get('ships_n', [1])) <= 1 else {})):
        if gw_inv and c.get('gw_kg') and abs(gw_inv - c['gw_kg']) > 0.05:
            issues.append(('MISMATCH', 'Gross weight: invoice %s KG, %s %s KG' % (fmt_thou(gw_inv), src, fmt_thou(c['gw_kg']))))
        if boxes and c.get('packages_qty') and boxes != c['packages_qty']:
            issues.append(('MISMATCH', 'Packages: invoice %s boxes, %s %s %s'
                           % (boxes, src, c['packages_qty'], c.get('packages_unit', ''))))
    if not sh['total_gw']:
        issues.append(('MISSING', 'Total Gross Weight not found (invoice / air waybill) - type it on tab 2 "Gross Wt (KG)"'))
    if not sh['packages']:
        issues.append(('MISSING', 'Number of boxes / packages not on the invoice - type it on tab 2'))
    info = {c['no']: c for c in cg.get('containers_info', [])}
    sh['containers'] = [dict(no=c, size=info.get(c, {}).get('size', ''), tne=info.get(c, {}).get('tne'),
                             seal=info.get(c, {}).get('seal', '')) for c in cg.get('containers', [])]
    dec = 0                   # container weight (TNE) always a round figure: .5 and above up (5.5 -> 6), below down (5.49 -> 5)
    gw_c = sh.get('total_gw')
    for c in sh['containers']:
        raw_t, how = c.get('tne'), 'B/L / customs instructions'
        if raw_t in (None, '') and gw_c and len(sh['containers']) == 1:
            raw_t, how = gw_c / 1000.0, 'gross weight %s KG / 1000' % fmt_thou(gw_c)
        if raw_t not in (None, ''):
            up = tne_round(raw_t)
            c['tne'] = up
            issues.append(('CHECK', 'Container %s weight %s TNE (%s) rounded to %s TNE'
                           % (c['no'], fmt_plain(round(float(raw_t), 6)), how, ('%.' + str(dec) + 'f') % up)))
        else:
            issues.append(('MISSING', 'Container %s weight (TNE) not known - type it in the Containers box on tab 3' % c['no']))
    if sh.get('containers_edit') is not None:
        sh['containers'] = [dict(c) for c in sh['containers_edit'] if c.get('no')]
        for c in sh['containers']:            # typed weights too: a round figure (rounded up)
            try:
                if c.get('tne') not in (None, ''):
                    c['tne'] = tne_round(c['tne'])
            except (TypeError, ValueError):
                pass
    sh['container'] = sh['containers'][0]['no'] if sh['containers'] else ''
    sh['cont_size'] = ''
    sh['currency'] = next((i['currency'] for i in sh['invs'] if i['currency']), '')
    sh['charges'] = round(sum(i.get('charges') or 0 for i in sh['invs']), 2)
    sh['freight_amount'] = cg.get('freight_amount')
    sh['freight_currency'] = cg.get('freight_currency', '')
    if cg.get('freight_amount'):
        issues.append(('CHECK', 'Freight %s %s from the customs instructions ("%s") - written as Freight Charge'
                       % (cg.get('freight_currency', ''), cg['freight_amount'], cg.get('permit_remarks', ''))))
    if cg.get('special'):
        issues.append(('CHECK', 'Forwarder special instructions: %s' % cg['special']))
    sh['incoterm'] = next((i['incoterm'] for i in sh['invs'] if i['incoterm']), '')
    curs = {i['currency'] for i in sh['invs'] if i['currency']}
    if len(curs) > 1:
        issues.append(('MISMATCH', 'More than one invoice currency: %s' % ', '.join(sorted(curs))))
    incs = {i['incoterm'] for i in sh['invs'] if i['incoterm']}
    if len(incs) > 1:
        issues.append(('CHECK', 'More than one Incoterm: %s (first used)' % ', '.join(sorted(incs))))
    # ---- raw lines + document checks
    raw = []
    for inv in sh['invs']:
        if inv.get('ocr'):
            issues.append(('CHECK', 'Invoice %s: %s' % (inv['invoice_no'] or inv['file'], inv.get('notes_ocr') or 'scanned invoice read by OCR - check item no / HS / qty / amounts')))
        for kk in ('notes_box', 'notes_total'):
            if inv.get(kk):
                issues.append(('CHECK', 'Invoice %s: %s' % (inv['invoice_no'], inv[kk])))
        if inv.get('hawb_from_name'):
            issues.append(('CHECK', 'Invoice %s: HAWB %s taken from the file name (not readable on the scan)'
                           % (inv['invoice_no'], inv['hawb'])))
        line_sum = round(sum(l['amount'] or 0 for l in inv['lines']) + (inv.get('charges') or 0) + (inv.get('freight') or 0), 2)
        cmode = sh.get('charges_mode') or ctx.get('charges_mode') or 'other'
        sh['charges_mode_used'] = cmode
        if inv.get('charges'):
            issues.append(('CHECK', 'Invoice %s: %s %s %s (in the invoice total, not in the item lines) - %s'
                           % (inv['invoice_no'], inv.get('charges_label') or 'Other charges', inv['currency'],
                              '{:,.2f}'.format(inv['charges']),
                              'ADDED to the last item of the invoice' if cmode == 'last' else
                              'declared as Other Taxable Charge (Invoice sheet: goods value + this charge)' if cmode == 'otc' else
                              'included in the Invoice sheet total (overall shipment value), not in any item')))
        if inv['total'] is not None and abs(line_sum - inv['total']) > 0.01:
            issues.append(('MISMATCH', 'Invoice %s: Total Amount %s but item lines add up to %s (%d lines read) - '
                           'check for a line the tool could not read' % (inv['invoice_no'], '{:,.2f}'.format(inv['total']),
                                                                          '{:,.2f}'.format(line_sum), len(inv['lines']))))
        elif inv['total'] is None:
            issues.append(('CHECK', 'Invoice %s: "Total Amount" not found - line total %s used'
                           % (inv['invoice_no'], '{:,.2f}'.format(line_sum))))
        ub = {b for l in inv['lines'] for b in l['boxes']}
        if inv['boxes'] and ub and len(ub) != inv['boxes']:
            issues.append(('MISMATCH', 'Invoice %s: Total Number of Boxes %s but %d different Box No. are listed'
                           % (inv['invoice_no'], inv['boxes'], len(ub))))
        nums = [int(l['line']) for l in inv['lines'] if l['line'].isdigit()]
        if nums and len(set(nums)) != len(nums):
            issues.append(('CHECK', 'Invoice %s: a Line No. is repeated' % inv['invoice_no']))
        for l in inv['lines']:
            if l['price'] and l['qty'] and l['amount'] and abs(l['price'] * l['qty'] - l['amount']) > max(0.05, l['amount'] * 0.0005):
                issues.append(('CHECK', 'Invoice %s line %s (%s): qty %s x unit price %s = %s, invoice shows %s'
                               % (inv['invoice_no'], l['line'], l['item'], fmt_plain(l['qty']), fmt_plain(l['price']),
                                  '{:,.2f}'.format(l['price'] * l['qty']), '{:,.2f}'.format(l['amount']))))
            if l['currency'] != inv['currency'] and inv['currency']:
                issues.append(('MISMATCH', 'Invoice %s line %s currency %s, invoice total in %s'
                               % (inv['invoice_no'], l['line'], l['currency'], inv['currency'])))
            dsc = l['desc']
            if l.get('type_code') and ctx.get('type_in_desc', True):
                dsc = clean(dsc + ' ' + l['type_code'])
            raw.append(dict(key='%s|%s' % (inv['invoice_no'], l['line']), line=l['line'], item=l['item'],
                            inv_desc=dsc, doc_nw=l.get('nw'), delivery=l.get('delivery', ''), qty=l['qty'], uom=l['uom'], amount=l['amount'], price=l['price'],
                            invoice=inv['invoice_no'], currency=inv['currency'] or l['currency'], doc_hs=l['hs'],
                            doc_coo=l['coo'], eccn=l['eccn'], order=l['order'], batches=list(l['lots']),
                            boxes=list(l['boxes']), src=inv['file'], pl_gw=l.get('pl_gw')))
        if inv.get('charges') and cmode == 'last' and inv['lines'] and raw and raw[-1]['invoice'] == inv['invoice_no']:
            raw[-1]['amount'] = round((raw[-1]['amount'] or 0) + inv['charges'], 2)     # user said yes: charges in last item
            raw[-1]['charge_add'] = inv['charges']
            raw[-1]['key_note'] = 'incl. %s %s' % (inv.get('charges_label') or 'charges', '{:,.2f}'.format(inv['charges']))
    sh['charges_in_items'] = round(sum(r.get('charge_add') or 0 for r in raw), 2)
    # ---- weights: packing list box weight per item (GW x 85 %) > invoice total GW x 85 %, split by item value
    gw = sh['total_gw']
    pl_rows = [r for r in raw if r.get('pl_gw') is not None]
    tv = sum(r['amount'] or 0 for r in raw)
    nw_total = None
    if nw_inv:
        nw_total, nw_how = nw_inv, 'Invoice total NW, split by value'
    elif gw:
        nw_total, nw_how = gw * 0.85, '85% of GW, split by value'
    doc_rows = [r for r in raw if r.get('doc_nw') is not None]
    if doc_rows and not ctx.get('ignore_line_nw'):
        # net weight per invoice line (Festo "Weight: 6,480 G") - gross weight split in proportion to net weight
        snw = sum(r['doc_nw'] for r in doc_rows)
        rest = [r for r in raw if r.get('doc_nw') is None]
        k = (gw / nw_inv) if (gw and nw_inv) else (gw / snw if gw and snw and not rest else None)
        for r in doc_rows:
            r['nw'], r['nw_src'] = r['doc_nw'], 'Invoice line weight'
            r['gw'] = r['doc_nw'] * k if k else None
        if rest:
            left = (nw_inv - snw) if nw_inv else ((gw * 0.85 - snw) if gw else 0)
            rv = sum(r['amount'] or 0 for r in rest)
            for r in rest:
                f = (r['amount'] or 0) / rv if rv else 0
                r['nw'] = left * f if left > 0 else None
                r['gw'] = r['nw'] * k if (k and r['nw'] is not None) else None
                r['nw_src'] = 'rest of total NW, split by value'
        if nw_inv and abs(snw - nw_inv) > max(0.05, nw_inv * 0.001) and not rest:
            issues.append(('MISMATCH', 'Item net weights add up to %s KG, invoice Net weight %s KG'
                           % (fmt_thou(round(snw, 3)), fmt_thou(nw_inv))))
        issues.append(('NET WEIGHT', 'Net weight per line from the invoice (%d line(s)); gross weight split in proportion '
                       'to net weight%s' % (len(doc_rows), ('; %d line(s) without weight: rest split by value' % len(rest))
                                            if rest else '')))
    elif pl_rows and not nw_inv:
        for r in pl_rows:
            r['nw'], r['gw'], r['nw_src'] = r['pl_gw'] * 0.85, r['pl_gw'], 'Packing list box GW x 85%'
        rest = [r for r in raw if r.get('pl_gw') is None]
        rv = sum(r['amount'] or 0 for r in rest)
        left = (gw - sum(r['pl_gw'] for r in pl_rows)) if gw else 0
        for r in rest:
            f = (r['amount'] or 0) / rv if rv else 0
            r['nw'], r['gw'], r['nw_src'] = (left * 0.85 * f if left > 0 else None), (left * f if left > 0 else None), \
                '85% of remaining GW, split by value'
        issues.append(('NET WEIGHT', 'Net weight not on the invoice - packing list box weights x 85%% for %d line(s)%s'
                       % (len(pl_rows), ('; %d line(s) without packing list weight: remaining GW x 85%%, split by value'
                                         % len(rest)) if rest else '')))
    elif nw_total and tv > 0:
        for r in raw:
            s = (r['amount'] or 0) / tv
            r['nw'], r['gw'], r['nw_src'] = nw_total * s, (gw * s if gw else None), nw_how
        if not nw_inv and not pl_rows:
            issues.append(('NET WEIGHT', 'Net weight not on the invoice - gross weight %s KG (%s) x 85%% = %s KG, '
                           'split by item value' % (fmt_thou(gw), sh['gw_src'], fmt_thou(round(nw_total, 3)))))
    else:
        for r in raw:
            r['nw'] = r['gw'] = None
            r['nw_src'] = ''
        issues.append(('NET WEIGHT', 'Net Weight unavailable and %s - weights left blank'
                       % ('Item Value unavailable for weight allocation' if gw else
                          'no gross weight in the invoice / air waybill - type the Gross Wt on tab 2 (NW = GW x 85 %, '
                          'split by item value)')))
    # ---- per line: master / HS / COO / UOM
    lines = []
    no_ctl = []
    auto_map = collections.OrderedDict()
    hs_missing = collections.OrderedDict()
    rpn_hits = []
    for r in raw:
        ov = sh['overrides'].get(r['key'], {})
        ms = master.get(r['item'], {})
        it = dict(r)
        it['desc'] = r['inv_desc'] or ms.get('desc', '')
        hs, hs_src = '', ''
        if r['doc_hs']:
            hs = r['doc_hs'] if re.fullmatch(r'\d{6}', str(r['doc_hs'])) else h8(r['doc_hs'])[:8]
            hs_src = 'Invoice' + (' (6-digit commodity code)' if len(hs) == 6 else '')
        elif ms.get('hs'):
            hs, hs_src = ms['hs'], 'Product master'
        hmap = ctx.get('hs_map') or {}
        doc_hs8 = hs
        if hs and hs in hmap and hmap[hs]:            # supplier HS -> Singapore HS (chosen once on tab 2, remembered)
            hs, hs_src = h8(hmap[hs])[:8], hs_src + ' > HS map'
        it['doc_hs8'] = doc_hs8
        if ms.get('hs') and hs and hs_src == 'Invoice' and ms['hs'] != hs:
            issues.append(('MISMATCH', 'Item %s HS code: invoice %s, product master %s (invoice used)' % (r['item'], hs, ms['hs'])))
        if hs and hsl and not (hsl.get(hs) or hsl.get(hs[:8])) and ctx.get('hs_auto', True) and 'hs' not in ov:
            code, how = hs_auto_map(hs, hsl, r.get('inv_desc', ''))
            if code:
                auto_map.setdefault((hs, code, how), []).append(r['item'])
                hs, hs_src = code, hs_src + ' > auto-mapped from %s' % hs
        if 'hs' in ov:                     # typed HS first: HS UOM, HS qty and CASC follow the new code
            hs, hs_src = (h8(ov['hs'])[:8] if ov['hs'] else ''), 'Edited'
        uom, uom_src = '', ''
        if hs and hsl:
            hit = hsl.get(hs) or hsl.get(hs[:8])
            it['hs_row'] = hit
            if hit:
                if hit[0] and hit[0] != hs:
                    issues.append(('CHECK', 'Item %s HS %s mapped to %s by the HS list' % (r['item'], hs, hit[0])))
                    hs, hs_src = hit[0], hs_src + ' > HS list'
                if hit[1]:
                    uom, uom_src = hit[1], 'HS list'
            else:
                hs_missing.setdefault(hs, []).append(r['item'])
                if ctx.get('blank_hs_not_in_list'):
                    hs, hs_src = '', ''
        if 'hs' in ov:
            hs, hs_src = (h8(ov['hs'])[:8] if ov['hs'] else ''), 'Edited'
        if hs and (not hs.isdigit() or len(hs) != 8):
            issues.append(('CHECK', 'Item %s HS code "%s" is not 8 digits' % (r['item'], hs)))
        if not hs:
            issues.append(('MISSING', 'HS Code missing for item %s (%s)' % (r['item'], it['desc'])))
        if not uom and ms.get('uom'):
            uom, uom_src = ms['uom'], 'Product master'
        if not uom and hs and ctx.get('default_uom'):
            uom, uom_src = ctx['default_uom'].upper(), 'Default'
        if 'uom' in ov:
            uom, uom_src = ov['uom'].upper(), 'Edited'
        coo, coo_src = '', ''
        if r['doc_coo']:
            coo, coo_src = r['doc_coo'], 'Invoice'
        elif ms.get('coo'):
            coo, coo_src = ms['coo'], 'Product master'
        if ms.get('coo') and coo and coo_src == 'Invoice' and ms['coo'] != coo:
            issues.append(('MISMATCH', 'Item %s COO: invoice %s, product master %s (invoice used)' % (r['item'], coo, ms['coo'])))
        if 'coo' in ov:
            coo, coo_src = ov['coo'].upper(), 'Edited'
        if not coo:
            issues.append(('MISSING', 'COO missing for item %s (%s)' % (r['item'], it['desc'])))
        if 'nw' in ov and ov['nw'] not in (None, ''):
            it['nw'], it['nw_src'] = fnum(ov['nw']), 'Edited'
        if 'desc' in ov and ov['desc']:
            it['desc_edit'] = ov['desc']
        # controlled item -> CASC Product Code 1: decided by the UPDATE HS CODE list control column
        why = []
        dcode = (ctx.get('control_casc') or 'MISC').upper()
        code = dcode
        row = (hsl.get(hs) or hsl.get(hs[:8])) if hs and hsl else None
        it['ca'] = ''
        if row and isinstance(row[3], dict) and HS_LIST_INFO.get('control_column'):
            info = row[3]
            cv = info['ctl_exp'] if exp else info['ctl_imp']
            ca = (info['ca_exp'] if exp else info['ca_imp']).strip(' -')
            it['ca'] = ca
            on, c = control_from_list(cv, dcode)
            if on:
                why.append('HS list %s: %s%s' % ('export' if exp else 'import', cv, (' ' + ca) if ca else ''))
                code = c
            if 'RPN' in ca.upper() and hs[:6] in HS_LIST_INFO.get('rpn', set()):
                rpn_hits.append('%s %s' % (r['item'], hs))
        elif hs and hsl and HS_LIST_INFO.get('control_column'):
            no_ctl.append(r['item'])
        eccn = (r.get('eccn') or '').upper()
        if ctx.get('control_eccn', False) and eccn and eccn not in NOT_CONTROLLED_ECCN:
            why.append('ECCN ' + r['eccn'])
        for pre in ctx.get('control_hs') or []:
            if pre and hs.startswith(pre):
                why.append('HS ' + pre)
                break
        casc, casc_src = ms.get('casc', ''), ('Product master' if ms.get('casc') else '')
        if why and not casc:
            casc, casc_src = code, 'Controlled (%s)' % ', '.join(why)
        if not ms.get('casc') and any(c[0] == 'IDAMISC' for c in product_codes(hs)):
            casc, casc_src = 'IDAMISC', 'HSCA list: HS %s has product code IDAMISC (always used)' % hs
        if 'casc' in ov:
            casc, casc_src = ov['casc'].upper(), 'Edited'
        it['pcodes'] = product_codes(hs)
        if False and casc and it['pcodes'] and casc not in {c[0] for c in it['pcodes']}:   # rule: IDAMISC else MISC
            issues.append(('MISMATCH', 'Item %s HS %s: CASC %s is not a product code of this HS in the HSCA list - pick one '
                           '(CASC cell dropdown): %s' % (r['item'], hs, casc, '; '.join(
                               '%s %s' % (c[0], c[1][:40]) for c in it['pcodes'][:6]) + (' ...' if len(it['pcodes']) > 6 else ''))))
        it.update(hs=hs, hs_src=hs_src, hs_uom=uom, uom_src=uom_src, coo=coo, coo_src=coo_src,
                  casc=casc, casc_src=casc_src, controlled=bool(why), sep=bool(ov.get('sep')),
                  brand=ov.get('brand', ms.get('brand', '') or ctx.get('brand', '')))
        lines.append(it)
    if hsl and not HS_LIST_INFO.get('control_column'):
        issues.append(('CHECK', 'UPDATE HS CODE list has no Control column - controlled items (CASC) not checked'))
    elif not hsl:
        issues.append(('CHECK', 'UPDATE HS CODE list not built in - controlled items (CASC) not checked'))
    for (a, b, how), its in auto_map.items():
        issues.append(('CHECK', 'HS %s not in the UPDATE HS CODE list - AUTO-MAPPED to %s (%s: %s) for %d line(s): %s'
                       % (a, b, how, (hsl.get(b) or ('', '', ''))[2][:60], len(its), ', '.join(list(dict.fromkeys(its))[:12]))))
    sh['hs_auto'] = [(a, b, how) for (a, b, how) in auto_map]
    if rpn_hits:
        issues.append(('CHECK', 'CHECK RPN CONTROL (HS list RPN table): %s' % ', '.join(dict.fromkeys(rpn_hits))))
    ctl = [x for x in lines if x['controlled']]
    if ctl:
        issues.append(('CHECK', '%d controlled line(s) - CASC Product Code 1 filled: %s' % (
            len(ctl), ', '.join('%s %s = %s (%s)' % (x['item'], x['hs'], x['casc'],
                                                    x['casc_src'].replace('Controlled (', '').rstrip(')')) for x in ctl))))
    sh['hs_missing'] = collections.OrderedDict((h, hs_candidates(h, hsl)) for h in hs_missing)
    for h, its in hs_missing.items():
        cands = sh['hs_missing'][h]
        issues.append(('MISSING', 'HS %s not in the UPDATE HS CODE list - HS UOM / control not checked%s (%d line(s): %s)%s'
                       % (h, ' - HS left blank' if ctx.get('blank_hs_not_in_list') else '', len(its),
                          ', '.join(dict.fromkeys(its)), ('  -> pick the SG code with "Map HS codes" on tab 2: %s'
                                                         % ', '.join(c for c, _ in cands[:6])) if cands else '')))
    dflt = [x['item'] for x in lines if x['uom_src'] == 'Default']
    if dflt:
        issues.append(('CHECK', 'HS UOM not in the HS list / product master for %d line(s) - default %s used (%s)'
                       % (len(dflt), ctx.get('default_uom', ''), ', '.join(dict.fromkeys(dflt)))))
    # ---- consolidation
    mode = ctx.get('consolidate') or 'none'
    groups = collections.OrderedDict()
    for it in lines:
        if it['sep']:                      # line kept separate (chosen on tab 2)
            k = ('SEP', it['key'])
        elif mode == 'item':
            k = (it['invoice'], it['item'], it['hs'], it['coo'], it['hs_uom'], it['uom'], it['casc'])
        elif mode == 'hs_coo':
            k = (it['hs'], it['coo'], it['hs_uom'], it['casc'])
        else:
            k = (it['key'],)
        groups.setdefault(k, []).append(it)
    tmpl = DEFAULT_DESC_TEMPLATE          # description only - no format
    dmax = int(ctx.get('desc_max') or 512)
    items = []
    for n, (k, g) in enumerate(groups.items(), 1):
        f = g[0]
        it = dict(f)
        it['keys'] = [x['key'] for x in g]
        it['lines'] = [x['line'] for x in g]
        it['n'] = n
        uoms = {x['uom'] for x in g}
        it['qty'] = sum(x['qty'] or 0 for x in g) if len(uoms) == 1 else None
        it['amount'] = round(sum(x['amount'] or 0 for x in g), 2)
        it['nw'] = sum(x['nw'] for x in g) if all(x['nw'] is not None for x in g) else None
        it['gw'] = sum(x['gw'] for x in g) if all(x['gw'] is not None for x in g) else None
        it['batches'] = list(dict.fromkeys(b for x in g for b in x['batches']))
        it['invoice'] = ', '.join(dict.fromkeys(x['invoice'] for x in g))
        it['eccn'] = ', '.join(dict.fromkeys(x['eccn'] for x in g if x['eccn']))
        it['controlled'] = any(x['controlled'] for x in g)
        it['merged'] = len(g) > 1
        # HS qty
        hq, note = None, ''
        rule = UOM_RULES.get(it['hs_uom'])
        if rule:
            basis, calc, fct = rule
            if basis == 'NET WEIGHT':
                base = it['nw']
                note = '' if base is not None else 'no net weight'
            else:
                ok = all((x['uom'] or '').upper() in COUNT_UOMS for x in g)
                base = sum(x['qty'] or 0 for x in g) if ok else None
                note = '' if ok else 'invoice UOM %s is not a piece count' % '/'.join(sorted(uoms))
                if ok and any((x['uom'] or '').upper() in PACK_UOMS for x in g):
                    issues.append(('CHECK', 'Item %s qty is in %s (packs) - used as the %s count'
                                   % (it['item'], '/'.join(sorted(uoms)), it['hs_uom'])))
            if base is not None:
                hq = base * fct if calc == 'x' else base / fct
        elif it['hs_uom']:                # other HS UOM (PCS, SET, PAR, MTR ...) = invoice qty
            if all((x['uom'] or '').upper() in COUNT_UOMS for x in g) or it['hs_uom'] in ('MTR', 'LTR', 'MTK', 'MTQ'):
                hq = sum(x['qty'] or 0 for x in g)
                issues.append(('CHECK', 'Item %s HS UOM %s has no rule - HS Qty = invoice qty %s %s'
                               % (it['item'], it['hs_uom'], fmt_plain(hq), '/'.join(sorted(uoms)))))
            else:
                note = 'no HS Qty rule for %s' % it['hs_uom']
        if hq is None and it['hs_uom']:
            issues.append(('CHECK', 'Item %s HS Qty left blank (%s %s): %s' % (it['item'], it['hs_uom'],
                                                                            rule_text(it['hs_uom']), note)))
        if hq is not None and 0 < hq < 0.00005:      # e.g. 0.01 KG in TNE - never declare 0.0000
            issues.append(('CHECK', 'Item %s HS Qty %.7f %s rounds to 0 - 0.0001 used' % (it['item'], hq, it['hs_uom'])))
            hq = 0.0001
        it['hs_qty'] = round(hq, 4) if hq is not None else None
        # description
        if len(g) == 1:
            ditem, ddesc = f['item'], f.get('desc_edit') or f['desc']
        else:
            ditem = ', '.join(dict.fromkeys(x['item'] for x in g))
            ddesc = ' / '.join(dict.fromkeys(x.get('desc_edit') or x['desc'] for x in g))
        vals = dict(ITEM=ditem, DESC=ddesc, QTY=fmt_plain(it['qty']) if it['qty'] is not None else
                    ' + '.join('%s %s' % (fmt_plain(x['qty']), x['uom']) for x in g),
                    UOM=f['uom'] if it['qty'] is not None else '', NW=fmt_thou(it['nw']), GW=fmt_thou(it['gw']),
                    HS=it['hs'], COO=it['coo'], BATCH=', '.join(it['batches']), INVOICE=it['invoice'],
                    ORDER=', '.join(dict.fromkeys(x['order'] for x in g if x['order'])),
                    ECCN=', '.join(dict.fromkeys(x['eccn'] for x in g if x['eccn'])))
        it['description'] = render_desc(tmpl, vals)
        if len(it['description']) > dmax:
            issues.append(('CHECK', 'Item %s description %d characters - cut to %d' % (it['item'], len(it['description']), dmax)))
            it['description'] = it['description'][:dmax].rstrip(' ,/')
        it['item'] = ditem
        items.append(it)
    sh['items'] = items
    sh['lines_n'] = len(lines)
    issues[:] = list(dict.fromkeys(issues))          # same message once
    sh['issues_all'] = issues
    sh['invoice_total'] = sum(i['total'] if i['total'] is not None else sum(l['amount'] or 0 for l in i['lines'])
                              for i in sh['invs']) if sh['invs'] else None
    if not sh['hawb']:
        issues.append(('MISSING', 'HAWB / HBL not found (add the B/L or customs instructions) - fill it on tab 2 or tab 3'))
    return sh


# ================================================================ excel writer
def write_excel(path, ships, items, cfg, lists, part=None, log=print):
    """One workbook from one or more resolved shipments (same HAWB). items = the rows for this file."""
    tpl = Template(cfg['template'])
    changed = {}
    s0 = ships[0]
    exp = s0['direction'] == 'EXPORT'
    ix = 1 if exp else 0
    col = lambda key: DIR_COLS[key][ix]
    permit = (cfg.get('permit_type') or '').upper()
    hawb = s0['hawb']
    cg = s0.get('cargo') or cfg.get('cargo') or {}
    hd = dict(cfg.get('header', {}))
    # cargo details fill what tab 3 left blank
    for k_cg, k_hd in (('mawb', 'mawb'), ('flight', 'flight'), ('transport_id', 'transport_id'), ('date', 'date'),
                       ('carrier_agent', 'carrier_agent'), ('loading_code', 'loading_port_code'),
                       ('loading_name', 'loading_port_name'), ('discharge_code', 'discharge_port_code'),
                       ('discharge_name', 'discharge_port_name'), ('mode', 'transport')):
        if cg.get(k_cg) and not hd.get(k_hd):
            hd[k_hd] = cg[k_cg]
    if cg.get('permit_type') in PERMIT_SCHEME and not permit:
        permit = cg['permit_type']
    locs = dict(LOCATION_NAMES)
    for side in ('receipt', 'release'):
        v = cg.get(side, '')
        if v and not hd.get(side + '_code'):
            code = v.strip().upper()
            if code in locs:
                hd[side + '_code'], hd[side + '_name'] = code, hd.get(side + '_name') or locs[code]
            elif re.fullmatch(r'[A-Z]{1,4}', code):
                hd[side + '_code'] = code
            else:
                hd[side + '_code'], hd[side + '_name'] = 'O', code
    # forwarder from the invoice "Carrier" when tab 3 is blank (full name / UEN via the forwarder table)
    if cg.get('forwarder') and not hd.get('forwarder') and cfg.get('fill_forwarder', True):   # import agent (customs instructions)
        hd['forwarder'] = cg['forwarder']
        hd['forwarder_uen'] = hd.get('forwarder_uen') or cg.get('forwarder_uen', '')
    if cg.get('carrier_agent_uen') and not hd.get('carrier_agent_uen') and hd.get('carrier_agent') == cg.get('carrier_agent'):
        hd['carrier_agent_uen'] = cg['carrier_agent_uen']
    car0 = s0['invs'][0].get('carrier') or ''
    if not car0 and s0.get('awbs'):            # scanned invoice: carrier from the air waybill issuer
        car0 = re.sub(r'\s*\(.*?\)|\b(GMBH|LTD|LIMITED|PTE|INC|CO|AG|BV|SA)\b\.?', '', s0['awbs'][0].get('issued_by', ''), flags=re.I)
        car0 = clean(car0).strip(' ,.;-').upper()
    if not hd.get('forwarder') and car0 and cfg.get('fill_forwarder', True):
        car = car0.upper()
        hd['forwarder'] = car
        for k, v in (cfg.get('forwarder_map') or {}).items():
            if k.upper() in car:
                name, _, uen = v.partition('|')
                hd['forwarder'] = name.strip() or car
                if uen.strip() and not hd.get('forwarder_uen'):
                    hd['forwarder_uen'] = uen.strip()
                break
    if not hd.get('carrier_agent') and hd.get('forwarder') and cfg.get('agent_from_forwarder', True):
        hd['carrier_agent'] = hd['forwarder']
        hd['carrier_agent_uen'] = hd.get('carrier_agent_uen') or hd.get('forwarder_uen', '')
    # ---------------- Header
    shd = tpl.sheet('Header')
    shd.clear_from(2)
    H = tpl.headers(shd)
    put = lambda h, v, st=None: _put(shd, H, h, 2, v, st)
    decl_text = hd.get('declaration_type', '')
    decl_code = lookup(lists.get('List-DeclarationType', []), decl_text) if decl_text else ''
    if not decl_text and permit in PERMIT_SCHEME:
        decl_code = PERMIT_SCHEME[permit]
        decl_text = next((t for t, c in lists.get('List-DeclarationType', []) if c == decl_code), '')
    has_cont = any(s.get('containers') for s in ships)
    cargo = hd.get('cargo_pack_type') or (('Containerized' if has_cont else 'Other Non-Containerized')
                                          if cfg.get('auto_cargo', True) else '')
    tmode = hd.get('transport') or ''
    if not tmode and has_cont and cfg.get('auto_transport', True):
        tmode = 'Sea'
    put(col('party_uen'), s0.get('party_uen', ''))
    put(col('party'), s0.get('party', ''))
    put(col('mawb'), hd.get('mawb', ''))
    put('Cargo Pack Type ' if 'cargo pack type ' in H else 'Cargo Pack Type', cargo)
    put('Cargo Pack Type Code', lookup(lists.get('List-CargoPacktype', []), cargo))
    put('Declaration Type', decl_text)
    put('Declaration Type Code', decl_code)
    put(col('carrier_agent'), hd.get('carrier_agent', ''))
    put(col('carrier_agent_uen'), hd.get('carrier_agent_uen', ''))
    put('Freight Forwarder', hd.get('forwarder', ''))
    put('Freight Forwarder Uen', hd.get('forwarder_uen', ''))
    put(col('transport'), tmode)
    put(col('transport_mode'), lookup(lists.get('List-TransportMode', []), tmode))
    if hd.get('date'):
        put(col('date'), *_date(hd['date'], tpl))
    put(col('flight'), hd.get('flight', ''))
    put(col('transport_id'), hd.get('transport_id', ''))
    if tmode == 'Sea' and not hd.get('vessel_type_list') and hd.get('transport_id') and cfg.get('auto_vessel_type', True):
        hd['vessel_type_list'] = 'CV: Cargo vessel'
    if tmode == 'Sea' and hd.get('vessel_type_list'):
        put('Vessel Type List', hd['vessel_type_list'])
        put('Vessel Type', lookup(lists.get('List-VesselType', []), hd['vessel_type_list']))
    for k, h in (('loading_port_code', 'Loading Port Code'), ('loading_port_name', 'Loading Port Name'),
                 ('discharge_port_code', 'Discharge Port Code'), ('discharge_port_name', 'Discharge Port Name'),
                 ('release_code', 'Release Location Code'), ('release_name', 'Release Location Name'),
                 ('receipt_code', 'Receipt Location Code'), ('receipt_name', 'Receipt Location Name')):
        if not exp and k.startswith('discharge'):          # import permit: discharge port is not filled
            continue
        put(h, hd.get(k, ''))
    if exp:
        put('Final Destination Country Code', s0.get('dest', ''))
    gws = [s['total_gw'] for s in ships if s.get('total_gw')]
    if part and all(it.get('gw') is not None for it in items):      # split Excel: GW of this part's items
        gws = [sum(it['gw'] for it in items)]
    if gws:
        if tmode == 'Sea':                # ocean permit: gross weight in TNE (KG / 1000)
            put('Total Gross Weight', round(sum(gws) / 1000.0, 3))
            put('Total Gross Weight Uom', 'TNE')
        else:
            put('Total Gross Weight', round(sum(gws), 3))
            put('Total Gross Weight Uom', 'KGM')
    pk = sum(s.get('packages') or 0 for s in ships)
    if pk and cfg.get('fill_outer_pack', True):
        put('Total Outer Pack Qty', pk)
        put('Total Outer Pack Uom', hd.get('outer_pack_uom') or 'PKG')
    nco = list(dict.fromkeys((c.get('no') if isinstance(c, dict) else str(c)) for s in ships for c in (s.get('containers') or [])
                            if (c.get('no') if isinstance(c, dict) else c)))
    if tmode == 'Sea' and nco and cfg.get('fill_outer_pack', True):      # sea job with containers -> UNT = number of containers
        put('Total Outer Pack Qty', len(nco))
        put('Total Outer Pack Uom', 'UNT')
    invs = [i['invoice_no'] for s in ships for i in s['invs'] if i.get('invoice_no')]
    rem = render_remarks(cfg.get('remarks_template', DEFAULT_REMARKS), dict(
        INV=', '.join(dict.fromkeys(invs)), HAWB=hawb, MAWB=hd.get('mawb', ''),
        JOB=cfg.get('job_no', '') or cg.get('job_no', ''),
        PO=', '.join(dict.fromkeys(i['customer_po'] for s in ships for i in s['invs'] if i.get('customer_po'))),
        ORDER=', '.join(dict.fromkeys(i['order_no'] for s in ships for i in s['invs'] if i.get('order_no')))))
    if part:
        rem = (rem + '\nPART %d OF %d' % part).strip()
    if cg.get('permit_remarks'):
        rem = (rem + '\n' + cg['permit_remarks']).strip()
    if rem:
        put("Trader's Remarks", rem)
    for h, v in (cfg.get('extra_header') or {}).items():       # any Header field typed on tab 3 ("More header fields")
        if h.lower() in H and str(v).strip():
            put(h, str(v).strip())
    changed[tpl.sheets['Header']] = shd.tostring()
    filled = {h for h in H if (lambda cx: bool(cx) and ('<v' in cx or '<is>' in cx))(shd.get('%s2' % H[h]))}
    check = EXPORT_HEADER_CHECK if exp else IMPORT_HEADER_CHECK
    header_blank = [h for h in check if h.lower() in H and h.lower() not in filled]
    # ---------------- Container
    if 'Container' in tpl.sheets:
        sc = tpl.sheet('Container')
        sc.clear_from(2)
        if cfg.get('fill_container', True):
            HC = tpl.headers(sc)
            r, seen = 2, set()
            for s in ships:
                for c in s.get('containers', []):
                    if c['no'] in seen:
                        continue
                    seen.add(c['no'])
                    _put(sc, HC, 'Container No', r, c['no'])
                    st = c.get('size_text') or (('FCL' + c['size']) if c.get('size') else '')
                    if st:
                        _put(sc, HC, 'Size/Type', r, st)
                    if c.get('tne') not in (None, ''):
                        _put(sc, HC, 'Weight', r, round(float(c['tne']), 3))
                    if c.get('seal'):
                        _put(sc, HC, 'Seal No', r, c['seal'])
                    r += 1
        changed[tpl.sheets['Container']] = sc.tostring()
    # ---------------- Invoice
    si = tpl.sheet('Invoice')
    si.clear_from(2)
    HI = tpl.headers(si)
    _put(si, HI, 'Invoice Number', 2, '-')
    _put(si, HI, 'Invoice Date', 2, *_date(datetime.date.today().strftime('%d/%m/%Y'), tpl))   # always today
    cur = s0['currency']
    total = sum(s['invoice_total'] or 0 for s in ships)        # always the overall shipment value (all invoices)
    chg = round(sum((s.get('charges') or 0) - (s.get('charges_in_items') or 0) for s in ships), 2)   # not added to items
    otc = bool(chg) and not part and (cfg.get('charges_other_line', False) or
                                      any((s.get('charges_mode_used') or s.get('charges_mode')) == 'otc' for s in ships))
    if otc:
        total -= chg                       # Invoice amount = goods; the charges go in "Other Taxable Charge" below
    if cur:
        _put(si, HI, 'Total Invoice Currency', 2, cur)
    if total:
        _put(si, HI, 'Total Invoice Amount', 2, round(total, 2), tpl.S2)
    if cfg.get('fill_term'):
        t = TERM_MAP.get(s0['incoterm'], s0['incoterm'])
        if t in [x for x, _ in lists.get('List-TermType', [])]:
            _put(si, HI, 'Term Type', 2, t)
    if otc:                                # charges declared on their own line (items + OTC = invoice total)
        _put(si, HI, 'Other Taxable Charge Currency', 2, cur)
        _put(si, HI, 'Other Taxable Charge Amount', 2, chg, tpl.S2)
    tt = TERM_MAP.get(s0.get('incoterm') or '', s0.get('incoterm') or '')     # charges allowed by the Term Type (TradeNet)
    allow = TERM_CHARGES.get(tt, ('frt', 'ins'))
    fa = next((s.get('freight_amount') for s in ships if s.get('freight_amount')), None)
    if fa and cfg.get('fill_freight', True) and 'frt' in allow:
        _put(si, HI, 'Freight Charge Currency', 2, next((s.get('freight_currency') for s in ships if s.get('freight_currency')), ''))
        _put(si, HI, 'Freight Charge Amount', 2, fa, tpl.S2)
    if not exp and (permit in IMPORT_PERMITS or not permit) and cfg.get('fill_insurance', True) and 'ins' in allow:
        ins = fnum(cfg.get('insurance_pct'))
        if ins:
            _put(si, HI, 'Insurance Charge %', 2, ins)
            if cfg.get('insurance_cur'):
                _put(si, HI, 'Insurance Charge Currency', 2, cfg['insurance_cur'])
    for h, v in (cfg.get('extra_invoice') or {}).items():      # any Invoice sheet field typed on tab 3
        if h.lower() in HI and str(v).strip():
            _put(si, HI, h, 2, fnum(v) if re.fullmatch(r'-?\d[\d,]*(\.\d+)?', str(v).strip()) and
                 any(w in h.lower() for w in ('amount', '%')) else str(v).strip())
    changed[tpl.sheets['Invoice']] = si.tostring()
    # ---------------- Item
    st = tpl.sheet('Item')
    st.clear_from(2)
    HT = tpl.headers(st)
    casc_rows = []
    for n, it in enumerate(items, 1):
        r = n + 1
        p = lambda h, v, s_=None: _put(st, HT, h, r, v, s_)
        p('Country of Origin', it['coo'])
        if it['hs']:
            p('HS Code', int(it['hs']) if it['hs'].isdigit() and not it['hs'].startswith('0') else it['hs'],
              None if it['hs'].isdigit() and not it['hs'].startswith('0') else tpl.STXT)
        p('Description', it['description'])
        if it['hs_qty'] is not None:
            p('HS Qty', it['hs_qty'], tpl.S4)
        if it['amount'] is not None:
            p('Total Line Amount', round(it['amount'], 2), tpl.S2)
        if it['brand']:
            p('Brand', it['brand'])
        if hawb:
            p(col('item_hawb'), hawb)
        if hd.get('mawb') and cfg.get('item_mawb', True):
            p(col('item_mawb'), hd['mawb'])
        p('HS UOM', it['hs_uom'])
        p('Invoice Number', it['invoice'])
        p('Item Currency', it['currency'])
        if it['casc']:
            p('CASC Product Code 1', it['casc'])
            if it['hs_qty'] is not None:
                p('CASC Product Code 1 Qty', it['hs_qty'], tpl.S4)
            if it['hs_uom']:
                p('CASC Product Code 1 UOM', it['hs_uom'])
            casc_rows.append((n, it['casc']))
    changed[tpl.sheets['Item']] = st.tostring()
    # ---------------- Casc sheet: not filled (user rule) - template sample rows cleared
    if 'Casc' in tpl.sheets:
        sk = tpl.sheet('Casc')
        sk.clear_from(2)
        changed[tpl.sheets['Casc']] = sk.tostring()      # Casc sheet left EMPTY (CASC is only on the Item sheet)
    if 'Items to be consolidated' in tpl.sheets:      # sample rows of the template left as they are
        pass
    tpl.save(path, changed)
    return header_blank


def file_name_for(group, cfg):
    s0 = group[0]
    if s0.get('file_name'):
        return safe_name(s0['file_name'])
    inv = '_'.join(i['invoice_no'] for s in group for i in s['invs'] if i.get('invoice_no'))
    vals = dict(HAWB=s0['hawb'], INVOICE=inv, DIR=s0.get('direction', ''))
    pat = cfg.get('file_pattern') or '{HAWB}'
    try:
        name = pat.format(**vals)
    except (KeyError, IndexError, ValueError):
        name = vals['HAWB']
    if not s0['hawb'] and '{HAWB}' in pat:
        name = 'NO_HAWB_INV_' + inv
    return safe_name(name)


def generate(ships, cfg, log=print):
    """One Excel per HAWB (split by max items) + CHECK_REPORT.xlsx + EXCEL_ALL.zip."""
    lists = template_lists(cfg['template'])
    out = cfg['out_dir']
    os.makedirs(out, exist_ok=True)
    groups = collections.OrderedDict()
    for s in ships:
        if s.get('include', True):
            groups.setdefault(s['hawb'] or ('INV:' + s['key']), []).append(s)
    mx = int(fnum(cfg.get('max_items')) or 0)
    files, used = [], set()
    for key, g in groups.items():
        items = [it for s in g for it in s['items']]
        chunks = [items[i:i + mx] for i in range(0, len(items), mx)] if mx and len(items) > mx else [items]
        base_name = file_name_for(g, cfg)
        for ci, chunk in enumerate(chunks, 1):
            name = base_name if len(chunks) == 1 else '%s_%d' % (base_name, ci)
            b, k = name, 2
            while name.upper() in used:
                name = '%s_%d' % (b, k)
                k += 1
            used.add(name.upper())
            path = os.path.join(out, name + '.xlsx')
            part = (ci, len(chunks)) if len(chunks) > 1 else None
            blank = write_excel(path, g, chunk, cfg, lists, part, log)
            amt = sum(it['amount'] or 0 for it in chunk)
            log('%s.xlsx  -  %s, %d item row(s), %s %s' % (name, g[0]['direction'], len(chunk), g[0]['currency'],
                                                          '{:,.2f}'.format(amt)))
            files.append(dict(name=name + '.xlsx', path=path, ships=g, items=chunk, header_blank=blank, part=part))
            if blank:
                log('   PLEASE FILL (Header, %s): %s' % (g[0]['direction'].lower(), ', '.join(blank)))
    rep = os.path.join(out, cfg.get('report_name') or 'CHECK_REPORT.xlsx')
    write_report(rep, files, cfg)
    log('%s  -  check report' % os.path.basename(rep))
    zpath = os.path.join(out, 'EXCEL_ALL.zip')
    with zipfile.ZipFile(zpath, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(f['path'], f['name'])
    log('EXCEL_ALL.zip  -  %d Excel file(s)' % len(files))
    return files, rep, zpath


def write_report(path, files, cfg):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    B, F = Font(bold=True, color='FFFFFF'), PatternFill('solid', fgColor='1F4E78')
    fills = {'MISMATCH': 'F8CBAD', 'MISSING': 'FFE699', 'ERROR': 'FF9999', 'NET WEIGHT': 'BDD7EE', 'CHECK': 'FFF2CC'}

    def sheet(title, cols, widths, first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill = B, F
            c.alignment = Alignment(wrap_text=True, vertical='center')
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[num2col(i)].width = w
        ws.freeze_panes = 'A2'
        return ws
    ws = sheet('Issues', ['Excel file', 'HAWB', 'Invoice(s)', 'Type', 'Detail'], [30, 16, 14, 13, 120], True)
    done = set()
    for f in files:
        for s in f['ships']:
            if id(s) in done:
                continue
            done.add(id(s))
            seen = set()
            for lvl, msg in s.get('issues_all', []):
                if (lvl, msg) in seen:
                    continue
                seen.add((lvl, msg))
                ws.append([f['name'], s['hawb'], ', '.join(i['invoice_no'] for i in s['invs']), lvl, msg])
                ws.cell(ws.max_row, 4).fill = PatternFill('solid', fgColor=fills.get(lvl, 'FFFFFF'))
    for fn, lvl, msg in cfg.get('global_issues', []):
        ws.append([fn, '', '', lvl, msg])
        ws.cell(ws.max_row, 4).fill = PatternFill('solid', fgColor=fills.get(lvl, 'FFFFFF'))
    if ws.max_row == 1:
        ws.append(['', '', '', 'OK', 'No missing data or mismatches found'])
    ws.auto_filter.ref = ws.dimensions
    ws = sheet('Header Check', ['Excel file', 'Header field left blank', 'Why'], [30, 34, 80])
    why = {'Importer Uen': 'Not on the invoice (only GST no.) - type it on tab 3 (remembered)',
           'Exporter UEN': 'Not on the invoice - type it on tab 3 (remembered)',
           'Arrival Date': 'Not on the invoice - paste the pre-alert / AWB on tab 1 or type on tab 3',
           'Departure Date': 'Not on the invoice - tab 3',
           'Inward MAWB': 'MAWB not on the invoice - paste the pre-alert / AWB on tab 1 or tab 3',
           'Outward MAWB': 'MAWB not on the invoice - tab 3',
           'In Flight Number': 'Flight not on the invoice - cargo details or tab 3',
           'In Transport Identifier': 'Flight / vessel not on the invoice - cargo details or tab 3',
           'Inward Transport': 'Air / Sea not on the invoice - choose on tab 3',
           'Loading Port Code': 'Origin port / airport not on the invoice - cargo details or tab 3',
           'Declaration Type': 'Choose Permit type / Declaration Type on tab 3'}
    for f in files:
        for h in f.get('header_blank', []):
            ws.append([f['name'], h, why.get(h, 'Not in the documents - fill on tab 3 if required')])
    if ws.max_row == 1:
        ws.append(['', 'All header fields filled'])
    ws = sheet('Summary', ['Excel file', 'Direction', 'HAWB', 'Importer / Exporter', 'UEN', 'Invoice(s)', 'Invoice date',
                           'Incoterm', 'Currency', 'Invoice total', 'Item value (this file)', 'Item rows', 'Invoice lines',
                           'Gross Wt (KG)', 'Net Wt (KG)', 'Boxes', 'Carrier', 'Ship from', 'Ship to', 'Source files'],
               [26, 9, 15, 42, 13, 14, 11, 8, 8, 15, 16, 8, 8, 12, 12, 7, 22, 40, 40, 40])
    for f in files:
        s0 = f['ships'][0]
        iv = sum(i['amount'] or 0 for i in f['items'])
        nw = sum(i['nw'] or 0 for i in f['items'])
        invs = [i for s in f['ships'] for i in s['invs']]
        ws.append([f['name'], s0['direction'], s0['hawb'], s0.get('party', ''), s0.get('party_uen', ''),
                   ', '.join(i['invoice_no'] for i in invs), ', '.join(dict.fromkeys(i['invoice_date'] for i in invs)),
                   s0['incoterm'], s0['currency'], sum(s['invoice_total'] or 0 for s in f['ships']), round(iv, 2),
                   len(f['items']), sum(s['lines_n'] for s in f['ships']),
                   sum(s.get('total_gw') or 0 for s in f['ships']), round(nw, 3),
                   sum(s.get('packages') or 0 for s in f['ships']), invs[0].get('carrier', ''),
                   invs[0].get('ship_from', ''), invs[0].get('ship_to', ''),
                   ', '.join(dict.fromkeys(i['file'] for i in invs))])
        for c in (10, 11):
            ws.cell(ws.max_row, c).number_format = '#,##0.00'
        for c in (14, 15):
            ws.cell(ws.max_row, c).number_format = '#,##0.000'
    ws = sheet('Items', ['Excel file', 'Row', 'Item No', 'Invoice No', 'Line No', 'Description (Excel)', 'Inv Qty',
                         'Inv UOM', 'Amount', 'Currency', 'Net Wt (KG)', 'NW source', 'Gross Wt (KG)', 'HS Code',
                         'HS source', 'HS UOM', 'UOM source', 'HS Qty', 'COO', 'COO source', 'ECCN', 'Order',
                         'CASC Product Code 1', 'CASC source', 'CA agency (HS list)', 'Lot'],
               [24, 5, 14, 11, 8, 70, 9, 7, 14, 8, 11, 22, 11, 10, 16, 7, 14, 11, 5, 12, 9, 12, 14, 30, 14, 18])
    Y = PatternFill('solid', fgColor='FFE699')
    for f in files:
        for row, it in enumerate(f['items'], 2):
            ws.append([f['name'], row, it['item'], it['invoice'], ', '.join(it['lines']), it['description'], it['qty'],
                       it['uom'], it['amount'], it['currency'], round(it['nw'], 3) if it['nw'] is not None else None,
                       it['nw_src'], round(it['gw'], 3) if it['gw'] is not None else None, it['hs'], it['hs_src'],
                       it['hs_uom'], it['uom_src'], it['hs_qty'], it['coo'], it['coo_src'], it.get('eccn', ''),
                       it.get('order', ''), it['casc'], it.get('casc_src', ''), it.get('ca', ''), ', '.join(it['batches'])])
            ws.cell(ws.max_row, 9).number_format = '#,##0.00'
            ws.cell(ws.max_row, 18).number_format = '0.0000'
            for c, bad in ((14, not it['hs']), (19, not it['coo']), (18, it['hs_qty'] is None),
                           (16, it['uom_src'] == 'Default')):
                if bad:
                    ws.cell(ws.max_row, c).fill = Y
    ws = sheet('UOM Rules', ['HS UOM', 'Based on', 'Calc', 'Factor', 'HS Qty ='], [10, 14, 8, 10, 30])
    for u, (basis, calc, f) in UOM_RULES.items():
        ws.append([u, basis, calc, f, rule_text(u)])
    ws.append(['other', 'QTY', 'x', 1, 'invoice qty (reported)'])
    ws.append([])
    ws.append(['UPDATE HS CODE list:', '', '', '', '%s - %s HS codes, columns %s' % (
        HS_LIST_INFO.get('source', '-'), HS_LIST_INFO.get('count', 0), HS_LIST_INFO.get('columns', {}))])
    wb.save(path)


def make_ctx(cfg, lists):
    return dict(master=load_master(cfg.get('master')), hs_list=load_hs_list(cfg.get('hs_list')),
                countries=Countries(lists.get('List-Country', [])), default_uom=cfg.get('default_uom', ''),
                blank_hs_not_in_list=cfg.get('blank_hs_not_in_list'), desc_template=cfg.get('desc_template'),
                desc_max=cfg.get('desc_max', 512), brand=cfg.get('brand', ''), cargo=cfg.get('cargo') or {},
                party_name=cfg.get('party_name', ''), party_uen=cfg.get('party_uen', ''),
                uen_book=cfg.get('uen_book') or {}, consolidate=cfg.get('consolidate', 'none'),
                direction_mode=cfg.get('direction', 'Auto'), control_casc=cfg.get('control_casc', 'MISC'),
                control_eccn=cfg.get('control_eccn', False), control_hs=cfg.get('control_hs') or [], cont_tne_dec=cfg.get('cont_tne_dec', 2), hs_auto=cfg.get('hs_auto', True),
                hs_map=cfg.get('hs_map') or {}, type_in_desc=cfg.get('type_in_desc', True),
                charges_as_other=cfg.get('charges_as_other', True), charges_mode=cfg.get('charges_mode', 'other'), ignore_line_nw=cfg.get('ignore_line_nw', False))


def run_folder(pdfs, cfg, log=print):
    """Command-line helper: parse PDFs, resolve, generate."""
    lists = template_lists(cfg['template'])
    ctx = make_ctx(cfg, lists)
    docs = [d for p in pdfs for d in parse_pdf(p, ctx['countries'])]
    ships, gi = build_shipments(docs)
    ctx['ships_n'] = list(ships)
    for s in ships.values():
        resolve(s, ctx)
    cfg = dict(cfg, global_issues=gi)
    return generate(list(ships.values()), cfg, log), ships, gi


# ================================================================ raw data (everything read from the documents)
def write_raw(path, docs, ships=None):
    """RAW_DATA.xlsx: every value the tool read, before any rule - Documents, Invoice lines, Packing list lines,
    Cargo (B/L / customs instructions / air waybill) and, when given, the final Items."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    def sheet(name, rows, first=False):
        ws = wb.active if first else wb.create_sheet(name)
        ws.title = name
        if not rows:
            ws.append(['(nothing)'])
            return
        cols = []
        for r in rows:
            for k in r:
                if k not in cols:
                    cols.append(k)
        ws.append(cols)
        for c in ws[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='2E3036')
        for r in rows:
            ws.append([_raw_val(r.get(k)) for k in cols])
        ws.freeze_panes = 'A2'
        for i, k in enumerate(cols, 1):
            ws.column_dimensions[num2col(i)].width = min(45, max(9, len(str(k)) + 2))
    skip = {'lines', 'path', 'words', 'pages', 'cargo', 'pkgs', 'img', 'chars', 'invoice_to_block', 'ship_to_block',
            'ship_from_block', 'invoice_from_block', 'carrier_block', 'packages_ids'}
    heads, inv_lines, pl_lines, cargo = [], [], [], []
    for d in docs:
        h = {'file': d.get('file', ''), 'kind': d.get('kind', ''), 'reader': d.get('reader', d.get('doc_type', ''))}
        h.update({k: v for k, v in d.items() if k not in skip and k not in h})
        heads.append(h)
        for l in d.get('lines') or []:
            row = {'file': d.get('file', ''), 'invoice_no': d.get('invoice_no', '')}
            row.update(l)
            (inv_lines if d.get('kind') == 'INV' else pl_lines).append(row)
        if d.get('kind') == 'AWB':
            c = {'file': d.get('file', ''), 'type': d.get('doc_type', 'air waybill'), 'hawb': d.get('hawb', ''),
                 'mawb': d.get('mawb', '')}
            c.update(d.get('cargo') or {})
            cargo.append(c)
    sheet('Invoice lines', inv_lines, first=True)
    sheet('Packing list lines', pl_lines)
    sheet('Cargo', cargo)
    sheet('Documents', heads)
    if ships:
        items = []
        for sh in ships:
            for it in sh.get('items', []):
                items.append({'HAWB': sh.get('hawb', ''), 'row': it.get('n'), 'item': it.get('item'),
                              'invoice': it.get('invoice'), 'lines': ', '.join(it.get('lines') or []),
                              'description': it.get('description'), 'qty': it.get('qty'), 'uom': it.get('uom'),
                              'amount': it.get('amount'), 'currency': it.get('currency'), 'nw': it.get('nw'),
                              'gw': it.get('gw'), 'nw_source': it.get('nw_src'), 'invoice_hs': it.get('doc_hs8') or it.get('doc_hs'),
                              'hs': it.get('hs'), 'hs_source': it.get('hs_src'), 'hs_uom': it.get('hs_uom'),
                              'hs_qty': it.get('hs_qty'), 'coo': it.get('coo'), 'casc': it.get('casc'),
                              'casc_source': it.get('casc_src'), 'ca': it.get('ca')})
        sheet('Final items', items)
    wb.save(path)
    return path


def _raw_val(v):
    if isinstance(v, (list, tuple, set)):
        return ', '.join(str(x) for x in v)
    if isinstance(v, dict):
        return '; '.join('%s=%s' % kv for kv in v.items())
    return v
