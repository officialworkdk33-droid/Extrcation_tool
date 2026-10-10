# -*- coding: utf-8 -*-
"""AGILENT PERMIT TOOL - engine
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
    txt = isinstance(v, str)
    if isinstance(v, float):
        v = int(v)
    s = str(v).strip().replace('.', '').replace(' ', '')
    if s.isdigit():
        if txt and len(s) == 6:                # "903300" printed on a document = HS 6 digits -> 90330000
            return s + '00'
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
    if sum(1 for r in out if any(v not in (None, '') for v in r)) <= 1:
        # file written without sheet dimensions (read-only mode sees one cell) -> read it normally
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        out = [list(r) for r in ws.iter_rows(values_only=True)]
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
# Singapore Customs units of quantity (HS UOM) -> meaning
HS_UNITS = {'CEN': 'Unit of One Hundred', 'CTM': 'Carat', 'GWH': 'Gigawatt-hour', 'KGM': 'Kilogram(s)', 'LTR': 'Litre(s)',
            'MIL': 'Unit of One Thousand', 'MTK': 'Square metre(s)', 'MTQ': 'Cubic metre(s)', 'MTR': 'Metre(s)',
            'NMB': 'Piece(s)', 'TEN': 'Unit of One Ten', 'TNE': 'Tonne', 'TPR': 'Unit of One Ten pairs',
            'VAL': 'Unit of Value of One Singapore Dollar'}
PAIR_UOMS = {'PR', 'PRS', 'PAIR', 'PAIRS', 'PAR', 'PA'}
# invoice unit -> (HS unit, factor): invoice qty x factor = qty in the HS unit
MEASURE_UOMS = {
    'MTR': {'M': 1, 'MTR': 1, 'MT': 1, 'METER': 1, 'METERS': 1, 'METRE': 1, 'METRES': 1, 'LM': 1, 'CM': 0.01, 'CMT': 0.01,
            'MM': 0.001, 'MMT': 0.001, 'KM': 1000, 'KMT': 1000, 'FT': 0.3048, 'FOT': 0.3048, 'FEET': 0.3048, 'IN': 0.0254,
            'INH': 0.0254, 'INCH': 0.0254, 'YD': 0.9144, 'YRD': 0.9144},
    'LTR': {'L': 1, 'LT': 1, 'LTR': 1, 'LTRS': 1, 'LITRE': 1, 'LITER': 1, 'LITRES': 1, 'ML': 0.001, 'MLT': 0.001, 'CL': 0.01,
            'KL': 1000, 'M3': 1000, 'MTQ': 1000, 'CBM': 1000, 'GAL': 3.785411784, 'GLL': 3.785411784, 'USG': 3.785411784},
    'MTK': {'M2': 1, 'MTK': 1, 'SQM': 1, 'SQ.M': 1, 'SQMT': 1, 'CM2': 0.0001, 'SQFT': 0.09290304, 'FT2': 0.09290304, 'SF': 0.09290304},
    'MTQ': {'M3': 1, 'MTQ': 1, 'CBM': 1, 'L': 0.001, 'LTR': 0.001, 'LITRE': 0.001, 'ML': 0.000001, 'CFT': 0.0283168466, 'FT3': 0.0283168466},
    'CTM': {'CT': 1, 'CTM': 1, 'CARAT': 1, 'CARATS': 1, 'G': 5, 'GRM': 5, 'GR': 5},
    'GWH': {'GWH': 1, 'MWH': 0.001, 'KWH': 0.000001, 'TWH': 1000},
}
# invoice UOMs that are a plain piece count (QTY based HS UOM rules only use these)
COUNT_UOMS = {'EA', 'EACH', 'PC', 'PCS', 'PCE', 'NMB', 'UN', 'UNIT', 'UNITS', 'NOS', 'NO', 'ST', 'SET'}
KG_UOMS = {'KG', 'KGS', 'KGM'}

DEFAULT_DESC_TEMPLATE = '{ITEM} , {DESC} , QTY : {QTY} {UOM} , NW : {NW} KGM , GW : {GW} KGM'
DESC_VARS = ('ITEM', 'DESC', 'QTY', 'UOM', 'NW', 'GW', 'HS', 'COO', 'BATCH', 'INVOICE')
REMARK_VARS = ('INV', 'HAWB', 'DELIVERY', 'JOB', 'PO')

# Permit type -> Declaration type code (master prompt)
PERMIT_SCHEME = collections.OrderedDict([('IG', 'GST'), ('ME', 'APS'), ('IT', 'SFZ'), ('IM', 'SFZ'),
                                         ('OD', 'DRT'), ('OT', 'DRT'), ('ID', 'DNG')])
EXTRA_DECL_TYPES = [('Duty and GST', 'DNG')]   # IPT DNG (dutiable liquor / tobacco for local use) - not in the template list
INSURANCE_PERMITS = ('ME', 'IM', 'IT', 'IG')
HS_KEYWORDS_DEFAULT = [('CAPACITOR', '85322400'), ('INDUCTOR', '85049090')]     # Settings > HS by description word
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

# Header / Item columns that belong to one direction only: an export permit shows + fills only the export ones,
# an import permit only the import ones (the web pages use the same patterns)
EXPORT_ONLY_RE = (r'^(exporter|outward|out |departure date|discharge port|final destination|consignee|end user|co type|'
                  r'certificate|additional certificate|co certificate|co manufacturing|co invoice|co percentage|co origin)')
IMPORT_ONLY_RE = r'^(importer|inward|in (?!pack)|arrival date)'


def other_direction_field(col, exp):
    """True when the column is for the other direction (e.g. 'Inward MAWB' on an export permit)."""
    return bool(re.match(IMPORT_ONLY_RE if exp else EXPORT_ONLY_RE, str(col or '').strip(), re.I))


# ---------------------------------------------------------------- customer rules ("Customer rules" tab of a job)
# Each rule belongs to one customer / supplier (found in the documents). After reading, the job lists the rules that
# apply; every rule is ticked (used) - untick one to leave it out for that job (form rules_off).
SANDVIK_WHO = r'SANDVIK\s+(?:SMC|MINING\s*(?:&|AND)\s*CONSTRUCTION)\s+LOGISTICS'
DOW_WHO = r'DOW\s+CHEMICAL\s+PACIFIC'
RS_WHO = r'\bRS\s+COMPONENTS\s+(?:LIMITED|LTD)'
DDI_WHO = r'DUPONT\s+DIVERSIFIED\s+INDUSTRIALS'
GIVAUDAN_WHO = r'\bGIVAUDAN\b'
VF_WHO = r'\bVF\s+(?:SINGAPORE\s+OVERSEAS|ASIA|ENTERPRISES|CORPORATION)|VF\s*-\s*INVOICE'
CUSTOMER_RULES = [
    dict(id='sandvik_export', customer='SANDVIK', who=SANDVIK_WHO, title='Export permit only',
         detail='Goods delivered from the Sandvik SMC Logistics Singapore hub: always an EXPORT permit (no import permit).'),
    dict(id='sandvik_desc', customer='SANDVIK', who=SANDVIK_WHO,
         title='Description = TOOLS OR PARTS INV NO : ( invoice no )',
         detail='Every item: "TOOLS OR PARTS INV NO : ( <invoice number> )", e.g. "TOOLS OR PARTS INV NO : ( 92898223 )".'),
    dict(id='dow_export', customer='DOW CHEMICAL PACIFIC', who=DOW_WHO, title='Export job',
         detail='DOW Chemical Pacific shipments are export permits.'),
    dict(id='dow_model', customer='DOW CHEMICAL PACIFIC', who=DOW_WHO, title='Model = shipment number',
         detail='Item sheet "Model" = the shipment number on the DOW invoice.'),
    dict(id='dow_brand', customer='DOW CHEMICAL PACIFIC', who=DOW_WHO, title='Brand = DOW',
         detail='Every item without a brand in the documents gets brand DOW.'),
    dict(id='dow_desc', customer='DOW CHEMICAL PACIFIC', who=DOW_WHO, title='Description + Our Reference',
         detail='Item description = the description followed by "Our Reference"; only the description when there is none.'),
    dict(id='dow_invno', customer='DOW CHEMICAL PACIFIC', who=DOW_WHO, title='Invoice number on the Invoice sheet',
         detail='Invoice sheet "Invoice Number" = the invoice number (e.g. 986049056) instead of "-".'),
    dict(id='rs_consolidate', customer='RS COMPONENTS', who=RS_WHO, title='Items combined by HS code + COO (descriptions merged, max 500)',
         detail='Invoice lines with the same HS code + country of origin = one item; descriptions joined with ", ", '
                'cut at a whole word so the description is never longer than 500 characters.'),
    dict(id='rs_casc', customer='RS COMPONENTS', who=RS_WHO, title='CASC Product Code 1 from the RS product code master (by HS)',
         detail='CASC Product Code 1 = the code of the HS code in MASTER LIST FOR PRODUCT CODE (Settings > Product masters); '
                'CASC qty / UOM = HS qty / UOM. HS codes missing from the master are listed - update the master.'),
    dict(id='rs_split', customer='RS COMPONENTS', who=RS_WHO, title='Max 50 items per Excel, IDAMISC / PCD in separate files',
         detail='GENERAL items in files of 50 items; items with CASC IDAMISC or PCD... each in their own file(s).'),
    dict(id='rs_remarks', customer='RS COMPONENTS', who=RS_WHO, title="Trader's Remarks: J-n / shipment / GBP@rate / (INV: ...) / container & OBL",
         detail="J-{J} / {SHIPMENT} / {CUR}@{RATE} / (INV:{INV}) / PARTS OF CONTAINER NO#{CONTAINER} & OBL NO#{OBL} - "
                "J counts up per Excel, INV = the sales invoices of that Excel. Type the exchange rate (and J start) below.",
         inputs=[('exchange_rate', 'Exchange rate (GBP@)'), ('j_start', 'J start (1, T8, J-07 ...)')]),
    dict(id='rs_invno', customer='RS COMPONENTS', who=RS_WHO, title='Invoice number on the Invoice sheet',
         detail='Invoice sheet "Invoice Number" = the consolidated invoice number (e.g. Z200826SIN20260819).'),
    dict(id='ddi_vendor', customer='DUPONT DIVERSIFIED INDUSTRIALS', who=DDI_WHO, side='buyer', title='Vendor product list: HS + CA/SC code + description by material number',
         detail='Material number of each invoice line -> the row of the chosen vendor list (Settings > Customer product lists): '
                'HS Code, CA/SC Code (= CASC Product Code 1) and Material description. Auto = every list.',
         inputs=[('cust_list', 'Vendor / product list', 'custlist')]),
    dict(id='ddi_desc', customer='DUPONT DIVERSIFIED INDUSTRIALS', who=DDI_WHO, side='buyer',
         title='Description = Customer Order Number : material number - description (master list)',
         detail='e.g. "4241363167 : 4018227 - MOLYKOTE LONGTERM 2/78 G GREASE 180KG". Customer Order Number = "Our Reference" of the '
                'line; description from the vendor list (invoice description when the material is not in the list).'),
    dict(id='ddi_casc', customer='DUPONT DIVERSIFIED INDUSTRIALS', who=DDI_WHO, side='buyer',
         title='Casc sheet: Item Number / Product Code Seq Number 1 / Code 1 = Control Substances',
         detail='For every item with a CASC product code: a Casc sheet row - Item Number (1, 2, 3 ... = the item row), '
                'Product Code Seq Number 1, Code 1 = "Control Substances" of the material in the vendor list.'),
    dict(id='givaudan_hs', customer='GIVAUDAN', who=GIVAUDAN_WHO, side='buyer',
         title='HS code from the Givaudan material master (by material number)',
         detail='Material number of each invoice line -> HS code (and CA/SC code when the list has one) from the chosen '
                'list (Settings > Customer product lists - upload the Givaudan master).',
         inputs=[('cust_list', 'Material master', 'custlist')]),
    dict(id='vf_brand', customer='VF', who=VF_WHO, side='buyer', title='Brand from the invoice',
         detail='Every item gets the brand printed on the VF invoice (e.g. TIMBERLAND, THE NORTH FACE).'),
    dict(id='vf_consolidate', customer='VF', who=VF_WHO, side='buyer', title='Items combined by HS code + COO + brand',
         detail='Invoice lines with the same HS code, country of origin and brand = one item (qty / amount / weights added).'),
]
for _r in CUSTOMER_RULES:                 # every customer rule so far is for the DGF OCEAN team (other teams: normal)
    _r.setdefault('teams', ('DGF OCEAN',))
RULE_BY_ID = {r['id']: r for r in CUSTOMER_RULES}
RS_REMARKS = "J-{J}\n{SHIPMENT}\n{CUR}@{RATE}\n(INV:{INV})\nPARTS OF CONTAINER NO#{CONTAINER} & OBL NO#{OBL}"


_CASC_HERE = {}


def _casc_master_here():
    """Desktop tool: RS product code master next to the program (masters/casc_by_hs/*.xlsx) -> {HS: CASC}."""
    if 'm' not in _CASC_HERE:
        out = {}
        import sys
        for base in (os.path.dirname(os.path.abspath(sys.argv[0] or '.')), getattr(sys, '_MEIPASS', ''), os.path.dirname(os.path.abspath(__file__))):
            d = os.path.join(base, 'masters', 'casc_by_hs')
            fs = sorted(f for f in os.listdir(d) if f.lower().endswith(('.xlsx', '.xlsm'))) if base and os.path.isdir(d) else []
            if fs:
                try:
                    import openpyxl
                    wb = openpyxl.load_workbook(os.path.join(d, fs[0]), read_only=True, data_only=True)
                    rows = list(wb.worksheets[0].iter_rows(values_only=True))
                    wb.close()
                    names = [str(v or '').strip().lower() for v in rows[0]]
                    ch, cc = names.index('hs code'), next(k for k, n in enumerate(names) if 'casc' in n)
                    for r in rows[1:]:
                        h = re.sub(r'\D', '', str(r[ch] if r[ch] is not None else ''))
                        if h:
                            out[h.zfill(8)[:8]] = str(r[cc]).strip().upper() if r[cc] not in (None, '') else ''
                except Exception:
                    out = {}
                break
        _CASC_HERE['m'] = out
    return _CASC_HERE['m']


def cut_words(text, n):
    """One line, single spaces, cut at a whole word to at most n characters (no trailing comma / space)."""
    text = re.sub(r'\s+', ' ', str(text or '')).strip()
    n = int(n)
    if len(text) <= n:
        return text
    cut = text[:n + 1]
    sp = cut.rfind(' ')
    cut = cut[:sp] if sp > n * 0.6 else text[:n]
    return cut.rstrip(' ,;-/').strip()[:n]


def j_label(start, idx):
    """J of Excel file idx (0-based): '1' -> 1, 2, 3 ; 'T8' -> T8, T9 ; 'J-07' -> J-07, J-08."""
    start = str(start or '1').strip() or '1'
    m = re.match(r'^(.*?)(\d+)$', start)
    if not m:
        return start
    head, digits = m.groups()
    return head + str(int(digits) + idx).zfill(len(digits) if head else 0)


def customers_in(text):
    """Customers whose rules apply to this text: ['SANDVIK', ...]."""
    out = []
    for r in CUSTOMER_RULES:
        if r['customer'] not in out and re.search(r['who'], text or '', re.I):
            out.append(r['customer'])
    return out


def tag_customers(docs, path):
    """Mark every invoice with the customers found in its document (rules), plus the values those rules need."""
    txt = ''
    if str(path).lower().endswith('.pdf'):
        try:
            import pdfplumber
            with pdfplumber.open(path) as pdf:
                txt = '\n'.join((pg.extract_text() or '') for pg in pdf.pages[:6])
        except Exception:
            txt = ''
    for d in docs:
        if d.get('kind') != 'INV':
            continue
        own = ' '.join(str(d.get(k) or '') for k in ('supplier', 'invoice_from', 'ship_from', 'reader', 'title')) + ' ' + \
            ' '.join(' '.join(d.get(k) or []) for k in ('invoice_from_block', 'ship_from_block'))
        # the SELLER decides (supplier fields + the page head) - an import consigned TO Sandvik gets no Sandvik rules
        d['customers'] = customers_in(own + '\n' + txt[:500]) if d.get('direction') != 'IMPORT' or own.strip() else []
        if d.get('direction') == 'IMPORT' and d['customers'] and not customers_in(own):
            d['customers'] = []
        d['customers'] = [c for c in d['customers'] if not any(r['customer'] == c and r.get('side') == 'buyer' for r in CUSTOMER_RULES)]
        if d.get('no_rules'):                  # e.g. Sandvik invoice INTO Singapore: the export rules of the supplier do not apply
            d['customers'] = []
        buyer = ' '.join(str(d.get(k) or '') for k in ('ship_to', 'invoice_to')) + ' ' + \
            ' '.join(' '.join(d.get(k) or []) for k in ('invoice_to_block', 'ship_to_block'))
        for r in CUSTOMER_RULES:                # rules of the importer / buyer (e.g. DuPont Diversified Industrials)
            if r.get('side') == 'buyer' and r['customer'] not in d['customers'] and re.search(r['who'], buyer, re.I):
                d['customers'].append(r['customer'])
        if 'DOW CHEMICAL PACIFIC' in d['customers']:
            t = txt or own
            m = re.search(r'Shipment\s*(?:No\.?|Number|#)\s*[:.]?\s*([A-Z0-9][A-Z0-9/-]{4,})', t, re.I)
            if m and not d.get('shipment_no'):
                d['shipment_no'] = m.group(1)
            m = re.search(r'Our\s*Ref(?:erence)?\.?\s*(?:No\.?|Number)?\s*[:.]?\s*([A-Z0-9][A-Z0-9/-]{2,})', t, re.I)
            if m and not d.get('our_ref'):
                d['our_ref'] = m.group(1)
    return docs


# ---------------------------------------------------------------- HS lists by DESCRIPTION (per customer, e.g. PANASONIC FACTORY)
# Settings > Product masters > "HS lists by description": an Excel (description + HS code) for one customer. When an
# item has no HS on the invoice / in the product master, its description is looked up: exact first, then the nearest
# description (word overlap weighted by how rare each word is). No good match -> HS stays empty.
HSD_STOP = {'W', 'WITH', 'FOR', 'OF', 'THE', 'AND', 'A', 'NS', 'NSB', 'MM', 'CM', 'M', 'PCS', 'PC', 'TYPE', 'NO', 'NET',
            'PACK', 'PIECES', 'PIECE', 'MAIN', 'MATERIAL'}


def hsd_norm(s):
    return re.sub(r'[^A-Z0-9]+', ' ', str(s or '').upper().replace('™', '')).strip()


def hsd_toks(s):
    out = []
    for t in hsd_norm(s).split():
        if t in HSD_STOP or (any(c.isdigit() for c in t) and len(t) <= 8):
            continue                     # sizes / numbers (4 5MM, 20W, 1PACK ...) do not decide the HS
        out.append(t)
    return out


def hsd_core(s):
    """Description without the extras in brackets / after ':' / NET= / 1PACK = ..."""
    c = re.split(r'\(|:|\bNET\s*=|\b\d+\s*PACK\b|\bMATERIAL\b', str(s or ''), flags=re.I)[0]
    return c.strip(' ,;')


class HSDescIndex:
    def __init__(self, pairs):
        self.exact = collections.defaultdict(collections.Counter)
        for d, h in pairs:
            h = re.sub(r'\D', '', str(h or ''))[:8]
            if d and len(h) >= 6:
                self.exact[hsd_norm(d)][h.ljust(8, '0')] += 1
        self.descs = list(self.exact)
        self.tok = [frozenset(hsd_toks(d)) for d in self.descs]
        df = collections.Counter(t for ts in self.tok for t in ts)
        n = len(self.descs) or 1
        self.idf = {t: math.log((n + 1) / (c + 0.5)) for t, c in df.items()}
        self.inv = collections.defaultdict(list)
        for i, ts in enumerate(self.tok):
            for t in ts:
                self.inv[t].append(i)

    def __len__(self):
        return len(self.descs)

    def _near(self, desc, min_score):
        q = frozenset(hsd_toks(desc))
        if not q:
            return None
        head = hsd_toks(desc)[0]
        w = lambda ts: sum(self.idf.get(t, 3.0) for t in ts)
        best = []
        for i in set(j for t in q for j in self.inv.get(t, ())):
            ts = self.tok[i]
            sc = w(q & ts) / (w(q | ts) or 1)
            if head in ts:
                sc += 0.1
            elif sc < 0.75:              # the main word of the description must be there (VANE is not a VACUUM PUMP)
                continue
            best.append((min(sc, 1.0), sum(self.exact[self.descs[i]].values()), i))
        if not best:
            return None
        best.sort(reverse=True)
        sc, _, i = best[0]
        if sc < min_score:
            return None
        top = collections.Counter()
        for s2, _c, j in best[:5]:
            if s2 >= sc - 0.02:
                top.update(self.exact[self.descs[j]])
        return top.most_common(1)[0][0], 'near', self.descs[i], round(sc, 2)

    def match(self, desc, min_score=0.55):
        """-> (hs, 'exact' / 'near', matched description, score) or None."""
        core = hsd_core(desc)
        for d in (desc, core):
            n = hsd_norm(d)
            if n in self.exact:
                return self.exact[n].most_common(1)[0][0], 'exact', n, 1.0
        return self._near(desc, min_score) or (self._near(core, min_score) if core != desc else None)


def load_hs_desc_list(path):
    """Excel / CSV with a description column and an HS code column (first sheet, heading row found by its words)."""
    pairs = []
    if str(path).lower().endswith('.csv'):
        import csv
        with open(path, encoding='utf-8-sig', errors='replace') as f:
            rows = list(csv.reader(f))
    else:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
        wb.close()
    di = hi = None
    start = 0
    for k, r in enumerate(rows[:20]):
        cells = [str(c or '').strip().lower() for c in r]
        for j, c in enumerate(cells):
            if di is None and 'desc' in c:
                di = j
            if hi is None and re.search(r'hs|commodity|tariff|code', c) and 'desc' not in c:
                hi = j
        if di is not None and hi is not None:
            start = k + 1
            break
    if di is None or hi is None:
        di, hi, start = 0, 1, 0
    for r in rows[start:]:
        if len(r) > max(di, hi) and r[di] and r[hi]:
            pairs.append((r[di], r[hi]))
    return pairs


# sea port name (as printed on B/L / arrival notices) -> UN/LOCODE, used when the document gives only the name
SEA_PORTS = {
    'SINGAPORE': 'SGSIN', 'NHAVA SHEVA': 'INNSA', 'JAWAHARLAL NEHRU': 'INNSA', 'MUNDRA': 'INMUN', 'CHENNAI': 'INMAA',
    'KOLKATA': 'INCCU', 'COCHIN': 'INCOK', 'VISAKHAPATNAM': 'INVTZ', 'SAN ANTONIO': 'CLSAI', 'VALPARAISO': 'CLVAP',
    'SHANGHAI': 'CNSHA', 'TIANJIN XINGANG PT': 'CNTXG', 'TIANJIN XINGANG': 'CNTXG', 'XINGANG': 'CNTXG', 'NINGBO': 'CNNGB',
    'QINGDAO': 'CNTAO', 'SHENZHEN': 'CNSZX', 'YANTIAN': 'CNYTN', 'SHEKOU': 'CNSHK', 'XIAMEN': 'CNXMN', 'DALIAN': 'CNDLC',
    'GUANGZHOU': 'CNCAN', 'HONG KONG': 'HKHKG', 'PORT KLANG': 'MYPKG', 'PASIR GUDANG': 'MYPGU', 'PENANG': 'MYPEN',
    'JAKARTA': 'IDJKT', 'TANJUNG PRIOK': 'IDTPP', 'SURABAYA': 'IDSUB', 'BANGKOK': 'THBKK', 'LAEM CHABANG': 'THLCH',
    'HO CHI MINH': 'VNSGN', 'CAT LAI': 'VNCLI', 'HAIPHONG': 'VNHPH', 'MANILA': 'PHMNL', 'BUSAN': 'KRPUS', 'TOKYO': 'JPTYO',
    'YOKOHAMA': 'JPYOK', 'KOBE': 'JPUKB', 'NAGOYA': 'JPNGO', 'OSAKA': 'JPOSA', 'KAOHSIUNG': 'TWKHH', 'KEELUNG': 'TWKEL',
    'SYDNEY': 'AUSYD', 'MELBOURNE': 'AUMEL', 'BRISBANE': 'AUBNE', 'FREMANTLE': 'AUFRE', 'AUCKLAND': 'NZAKL',
    'CHATTOGRAM': 'BDCGP', 'CHITTAGONG': 'BDCGP', 'COLOMBO': 'LKCMB', 'KARACHI': 'PKKHI', 'JEBEL ALI': 'AEJEA',
    'DUBAI': 'AEJEA', 'ROTTERDAM': 'NLRTM', 'HAMBURG': 'DEHAM', 'ANTWERP': 'BEANR', 'FELIXSTOWE': 'GBFXT',
    'LOS ANGELES': 'USLAX', 'LONG BEACH': 'USLGB', 'NEW YORK': 'USNYC', 'HOUSTON': 'USHOU', 'DURBAN': 'ZADUR',
    'CALLAO': 'PECLL', 'SANTOS': 'BRSSZ', 'MANZANILLO': 'MXZLO'}


def sea_port_code(name):
    n = re.sub(r'\s+', ' ', re.sub(r',.*$', '', str(name or '')).upper()).strip()
    return SEA_PORTS.get(n) or next((c for k, c in SEA_PORTS.items() if n.startswith(k + ' ') or n.endswith(' ' + k)), '')


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
    words = pg.get('hwords') or pg['words']
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
    t = re.sub(r'(?<=\s)[|]+(?=\s)|[|]+$', ' ', t)
    t = re.sub(r'[ \t]{2,}', ' ', t)
    t = re.sub(r'\bEAR[QSOo09]{2,3}\b', 'EAR99', t)                 # EARQ9 / EARS9 -> EAR99
    t = re.sub(r'\bIll\b', 'III', t)
    t = re.sub(r'\bIl\b', 'II', t)
    t = re.sub(r'\|(?=-?[A-Za-z])', 'I', t)                         # |-AS -> I-AS
    t = t.replace('\u00a7', '5')                                     # §500-1193 -> 5500-1193
    t = re.sub(r'(?<=\d)m!', 'ml', t)
    return t


def remove_lines(img, dpi=300):
    """Whiten long horizontal / vertical ruling lines (table grids) so Tesseract reads the cells."""
    try:
        import numpy as np
    except ImportError:
        return img
    from PIL import Image
    a = np.asarray(img).copy()
    dark = a < 150
    def runs(mask, L, axis):
        m = mask if axis == 1 else mask.T
        c = np.cumsum(np.pad(m, ((0, 0), (1, 0))).astype(np.int32), axis=1)
        full = (c[:, L:] - c[:, :-L]) == L                      # window [j, j+L) fully dark
        cov = np.zeros(m.shape, dtype=np.int32)
        f = np.pad(full.astype(np.int32), ((0, 0), (0, L - 1)))[:, :m.shape[1]]
        cs = np.cumsum(np.pad(f, ((0, 0), (1, 0))), axis=1)
        idx = np.arange(m.shape[1])
        lo = np.clip(idx - L + 1, 0, None)
        cov = cs[:, idx + 1] - cs[:, lo]
        out = cov > 0
        return out if axis == 1 else out.T
    Lh, Lv = max(30, int(dpi * 0.35)), max(24, int(dpi * 0.2))      # ~9 mm horizontal, ~5 mm vertical
    lines = runs(dark, Lh, 1) | runs(dark, Lv, 0)
    # widen the mask 2 px so the grey edges of the lines go too
    m = lines.copy()
    for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1), (2, 0), (-2, 0), (0, 2), (0, -2)):
        m |= np.roll(np.roll(lines, dy, axis=0), dx, axis=1)
    a[m & (a < 200)] = 255
    return Image.fromarray(a)


def _ocr_score(d):
    """How readable an OCR result is: confident real words (rotated / upside-down scans score near 0)."""
    n = 0
    for t, c in zip(d['text'], d['conf']):
        t = (t or '').strip()
        try:
            c = float(c)
        except ValueError:
            c = -1
        if c >= 70 and len(t) >= 3 and re.fullmatch(r"[A-Za-z][A-Za-z.,:'()/-]*", t):
            n += 1
    return n


def _ocr_text_words(img, k, d=None):
    """Tesseract block mode: page text (lines) + word phrases with positions (for column reading)."""
    import ocr_runner
    if d is None:
        d = ocr_runner.image_to_data(img, 6)
    lines = collections.OrderedDict()
    for i, t in enumerate(d['text']):
        t = (t or '').strip()
        if t:
            lines.setdefault((d['block_num'][i], d['par_num'][i], d['line_num'][i]), []).append(t)
    text = ocr_fix('\n'.join(' '.join(v) for v in lines.values()))
    return text, _phrases(d, k)


def _phrases(d, k):
    toks = []
    for i, t in enumerate(d['text']):
        t = (t or '').strip().strip('|_')
        if not t or float(d['conf'][i]) < 20:
            continue
        toks.append(dict(text=ocr_fix(t), x0=d['left'][i] * k, x1=(d['left'][i] + d['width'][i]) * k, top=d['top'][i] * k,
                         key=(d['block_num'][i], d['par_num'][i], d['line_num'][i])))
    out = []
    for w in toks:
        if out and out[-1]['key'] == w['key'] and w['x0'] - out[-1]['x1'] < 6.5:
            out[-1]['text'] += ' ' + w['text']
            out[-1]['x1'] = w['x1']
        else:
            out.append(dict(w))
    return [dict(text=w['text'], x0=w['x0'], x1=w['x1'], top=w['top']) for w in out]


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
        turn = {'deg': 0}                       # sideways scans: the angle that worked is tried first on the next pages

        def upright(orig):
            """OCR data of the page turned the readable way up (0 / 90 / 270 / 180 degrees)."""
            tried = []
            for deg in [turn['deg']] + [a for a in (0, 90, 270, 180) if a != turn['deg']]:
                im = orig.rotate(deg, expand=True) if deg else orig
                cl = remove_lines(im, OCR['dpi'])
                d = ocr_runner.image_to_data(cl, 6)
                s = _ocr_score(d)
                tried.append((s, deg, im, cl, d))
                if s >= 25:
                    break
            s, deg, im, cl, d = max(tried, key=lambda x: x[0])
            if deg != turn['deg'] and s >= 25:
                turn['deg'] = deg
                if log:
                    log('   OCR %s: pages are turned - reading them rotated %d degrees' % (os.path.basename(path), deg))
            return im, cl, d

        def one(i):
            orig, img, d = upright(imgs.pop(i))
            text, words = _ocr_text_words(img, k, d)
            first = i == 0 or re.search(r'Page\s*1\s*of\s*\d+', text)
            return i, dict(text=text, words=words, hwords=(_ocr_words(orig, k) + words) if first else words, chars=None,
                           ocr=True, img=orig if first else None, k=k)
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


def parse_pdf(path, countries=None, log=None, profiles=None):
    """Any supplier file (PDF / scan / Excel) -> list of documents (see readers.py)."""
    import readers
    return tag_customers(readers.read_file(path, countries, profiles, log), path)


def parse_pdf_pages(path, allp, countries=None):
    """Agilent invoice reader (one PDF can hold several invoices)."""
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
            d = parse_invoice_pages([(p['text'], p.get('hwords') or p['words']) for p in pgs], path, countries)
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


DHL_SYS_FIELDS = [  # DHL system block "LABEL:old/new" (e.g. "MASTER WAYBILL:-/HLCUSIN260918760") - the new value is used
    ('hawb', r'HOUSE WAYBILL'), ('mawb', r'MASTER WAYBILL'), ('mode', r'(?:OUTWARD|INWARD) TRANSPORT MODE'),
    ('discharge_code', r'DISCHARGE PORT CODE'), ('loading_code', r'LOADING PORT CODE'),
    ('flight', r'(?:OUTWARD|INWARD) FLIGHT/VOYAGE NUMBER'), ('transport_id', r'(?:OUTWARD|INWARD) AIRCRAFT REGISTRATION NUMBER/VESSEL NAME'),
    ('term_type', r'TERM TYPE'), ('packages_unit', r'TOTAL OUTER PACK UOM'), ('packages_qty', r'TOTAL OUTER PACK'),
    ('gw_uom', r'TOTAL GROSS WEIGHT UOM'), ('gw_kg', r'TOTAL GROSS WEIGHT'), ('direction', r'PERMIT TYPE'),
    ('etd_time', r'ESTIMATED DEPARTURE TIME'), ('eta_time', r'ESTIMATED ARRIVAL TIME')]
MONTH_N = {m: i for i, m in enumerate(('JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'), 1)}


def _dhl_instructions(text, out):
    """DHL e-mail instructions: the system block (LABEL:old/new), booking ref, CO type, special instructions
    (SOB date, '6 digits HS', 'include description: ...', gross weight with pallets). Returns the remaining text."""
    keep = []
    lines = text.splitlines()
    i = 0
    sys_hit = False
    while i < len(lines):
        ln = lines[i].strip()
        m = re.match(r'^(.+?)\s*:\s*([^/]*?)\s*/\s*(.*)$', ln)
        hit = None
        if m:
            for key, lab in DHL_SYS_FIELDS:
                if re.fullmatch(lab, m.group(1).strip().upper()):
                    hit = key
                    break
        if hit:
            v = clean(m.group(3)) if clean(m.group(3)) not in ('', '-') else clean(m.group(2)).strip('-')
            if hit == 'hawb' and m.group(2).strip() not in ('', '-'):
                out['dhl_ref'] = m.group(2).strip()
            if v:
                out[hit] = v
            sys_hit = True
            i += 1
            continue
        m = re.match(r'(CARRIER\s*BOOKING\s*REF)[^:]*:\s*(.*)$', ln, re.I)
        if m:
            if clean(m.group(2)):
                out['booking_ref'] = clean(m.group(2))
            i += 1
            continue
        m = re.match(r'(CO\s*TYPE)[^:]*:\s*(.*)$', ln, re.I)
        if m:
            if clean(m.group(2)):
                out['co_type'] = clean(m.group(2)).upper()
            i += 1
            continue
        m = re.match(r'SPECIAL\s*INSTRUCTIONS?\s*:\s*(.*)$', ln, re.I)
        if m:
            sp = [m.group(1)]
            i += 1
            while i < len(lines) and lines[i].strip() and not re.match(r'^[A-Z][A-Z /()\-]{2,40}:\s*\S', lines[i].strip()) \
                    or (i < len(lines) and re.match(r'^(USE|THE|PLS|PLEASE)\b', lines[i].strip(), re.I)):
                sp.append(lines[i].strip())
                i += 1
            out['special'] = clean(' '.join(x for x in sp if x))
            continue
        keep.append(lines[i])
        i += 1
    sp = (out.get('special') or '').upper()
    if sp:
        m = re.search(r'\bSOB\s*:?\s*(\d{1,2})\s*([A-Z]{3})[A-Z]*\b(?:\s*(\d{4}))?', sp)
        if m and m.group(2) in MONTH_N:
            today = datetime.date.today()
            y = int(m.group(3)) if m.group(3) else today.year
            d = datetime.date(y, MONTH_N[m.group(2)], int(m.group(1)))
            if not m.group(3) and (today - d).days > 180:
                d = datetime.date(y + 1, d.month, d.day)
            out['date'] = d.strftime('%d/%m/%Y')
            out['date_note'] = 'departure date = shipped on board (SOB) date in the instructions'
        if re.search(r'\b6\s*DIGITS?\s*HS', sp):
            out['co_hs6'] = True
        m = re.search(r'INCLUDE\s+(?:THE\s+)?DESCRIPTION\s*:?\s*([A-Z0-9 ,.&/-]+?)(?=\s+THE\s+GROSS|\s+PLS\b|\s+PLEASE\b|$)', sp)
        if m:
            out['desc_include'] = clean(m.group(1)).strip(' .,')
        if re.search(r'GROSS WEIGHT.*WITH PALLET', sp):
            out['gw_with_pallets'] = True
    if sys_hit:
        out['instructions'] = True
    if out.get('direction'):
        out['direction'] = 'EXPORT' if 'EXP' in out['direction'].upper() else 'IMPORT' if 'IMP' in out['direction'].upper() else ''
    if out.get('mode'):
        out['mode'] = out['mode'].title()
    if out.get('packages_qty'):
        try:
            out['packages_qty'] = int(fnum(out['packages_qty']))
        except (TypeError, ValueError):
            out.pop('packages_qty')
    if out.get('gw_kg'):
        out['gw_kg'] = fnum(out['gw_kg'])
    return '\n'.join(keep)


def parse_cargo(text):
    """Cargo details from a pasted pre-alert / AWB / B/L / e-mail ("LABEL: value" lines)."""
    out = {}
    if not text or not text.strip():
        return out
    text = _dhl_instructions(text, out)
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
    if out.get('carrier_agent'):                   # "HAPAG LLOYD AG HAPAG LLOYD" -> "HAPAG LLOYD AG"
        w = out['carrier_agent'].split()
        for k in range(len(w) // 2, 1, -1):
            if w[-k:] == w[:k]:
                out['carrier_agent'] = ' '.join(w[:-k])
                break
    if not out.get('mawb') and out.get('booking_ref') and out.get('direction') == 'EXPORT':
        out['mawb'] = out['booking_ref']          # export: carrier booking ref = Outward MAWB / OBL when no master B/L
    conts = list(dict.fromkeys(re.findall(r'\b([A-Z]{4}\d{7})\b', U)))
    if conts:
        out['containers'] = conts
    if 'packages' in out:
        m = re.search(r'(\d[\d,]*)\s*([A-Z]{2,5})?', out['packages'].upper())
        out['packages_qty'] = int(fnum(m.group(1))) if m else None
        out['packages_unit'] = (m.group(2) or '') if m else ''
    if 'gw' in out and not out.get('gw_kg'):
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
def _detail_copies(docs, gi):
    """An invoice also given as a detail list (e.g. Clasquin 'detail_hscode' Excel of the Millesima PDF invoice): the PDF
    invoice is used, the detail list only checks it (line amount / litres / alcohol %) and is then left out."""
    invs = [d for d in docs if d.get('kind') == 'INV']
    for dc in [d for d in invs if d.get('detail_copy')]:
        main = next((d for d in invs if d is not dc and not d.get('detail_copy') and d.get('invoice_no') == dc.get('invoice_no')), None)
        if not main:
            continue
        by = {l['line']: l for l in main['lines']}
        diff = []
        for l in dc['lines']:
            m = by.get(l['line'])
            if not m:
                diff.append('line %s not on the invoice' % l['line'])
                continue
            for f, a, b in (('amount', l.get('amount'), m.get('amount')), ('litres', l.get('ltr'), m.get('ltr')),
                            ('alcohol %', (l.get('liquor') or {}).get('abv'), (m.get('liquor') or {}).get('abv'))):
                if a is not None and b is not None and abs(a - b) > 0.01:
                    diff.append('line %s %s %s / invoice %s' % (l['line'], f, a, b))
        if len(dc['lines']) != len(main['lines']):
            diff.append('%d lines / invoice %d lines' % (len(dc['lines']), len(main['lines'])))
        gi.append((dc['file'], 'MISMATCH' if diff else 'CHECK', 'Same invoice %s as %s - %s' % (
            dc['invoice_no'], main['file'], ('differs: ' + '; '.join(diff[:8])) if diff else
            'all %d lines agree (amount, litres, alcohol %%); the PDF invoice is used' % len(dc['lines']))))
        dc['kind'] = 'OTHER'
        dc['notes'] = (dc.get('notes') or []) + ['detail list of invoice %s - used to check it, not as a second invoice' % dc['invoice_no']]


def build_shipments(docs):
    """Invoices + packing lists + air waybills / B/L -> shipments (one per HAWB)."""
    ships = collections.OrderedDict()
    gi = []
    for d in docs:                             # common keys for every reader
        if d.get('kind') == 'INV':
            for k, v in (('mawb', ''), ('uen', ''), ('coo_text', ''), ('hs_global', ''), ('package_unit', ''),
                         ('reader', 'Agilent'), ('delivery', '')):
                d.setdefault(k, v)
            for l in d['lines']:
                l.setdefault('nw', None)
                l.setdefault('gw', None)
    _detail_copies(docs, gi)
    awbs = [d for d in docs if d.get('kind') == 'AWB']
    pls = [d for d in docs if d.get('kind') == 'PL']
    invs = [d for d in docs if d.get('kind') == 'INV']
    # the same invoice added twice (e.g. full pack + split PDF, or a copy) -> used once
    seen, keep = {}, []
    for d in invs:
        k = (squash(d.get('invoice_no') or ''), round(sum(l['amount'] or 0 for l in d['lines']), 2), len(d['lines']))
        if k[0] and k in seen:
            gi.append((d['file'], 'CHECK', 'invoice %s already read from %s - skipped (same invoice twice)'
                       % (d['invoice_no'], seen[k])))
            continue
        seen[k] = d['file']
        keep.append(d)
    invs = keep
    seen, keep = {}, []
    for p in pls:                              # the same packing list twice -> used once (weights not doubled)
        k = (tuple(sorted(squash(r) for r in p.get('invoice_refs', []))), squash(p.get('delivery') or ''),
             tuple(p.get('containers') or []), p.get('total_gw'), p.get('total_nw'), len(p['lines']))
        if (k[0] or k[1]) and k in seen:
            gi.append((p['file'], 'CHECK', 'packing list already read from %s - skipped (same document twice)' % seen[k]))
            continue
        seen[k] = p['file']
        keep.append(p)
    pls = keep
    for d in docs:
        if d.get('kind') not in ('INV', 'PL', 'AWB'):
            gi.append((d['file'], 'CHECK', '; '.join(d.get('notes') or ['not an invoice'])))
    # several transport documents of one shipment: the house B/L number wins over a forwarder reference of it
    # (DSV shipment no. "S" + B/L no., e.g. SRAF0583645 = B/L RAF0583645)
    awbs.sort(key=lambda a: -a.get('priority', 0))
    canon = []
    for a in awbs:
        h = squash(a.get('hawb') or '')
        if not h:
            continue
        hit = next((c for c in canon if c == h or (len(h) == len(c) + 1 and h[0] == 'S' and h[1:] == c)
                    or (len(c) == len(h) + 1 and c[0] == 'S' and c[1:] == h)), None)
        if hit and hit != h:
            a['notes'].append('house B/L %s here is the forwarder reference of B/L %s - %s used' % (a['hawb'], hit, hit))
            a['hawb'] = hit
            if a.get('cargo', {}).get('hawb'):
                a['cargo']['hawb'] = hit
            if a.get('cargo', {}).get('mawb') and squash(a['cargo']['mawb'])[-len(hit):] == hit:
                a['cargo'].pop('mawb')
        elif not hit:
            canon.append(h)
    # several house B/Ls listed on one arrival notice / customs instructions under one ocean B/L -> one permit
    hbl_group = {}
    for a in awbs:
        hl = [squash(h) for h in (a.get('house_list') or []) if squash(h)]
        if len(hl) > 1:
            gk = 'OBL:' + (squash(a.get('mawb') or '') or hl[0])
            for h in hl:
                hbl_group.setdefault(h, gk)
    if hbl_group:
        gks = {}
        for a in awbs:
            if a.get('mawb') and len(a.get('house_list') or []) > 1:
                gks.setdefault('OBL:' + squash(a['mawb']), a['mawb'])
        for d in invs + pls:
            if d.get('hawb') and squash(d['hawb']) in hbl_group:
                for l in d.get('lines', []):
                    l.setdefault('hawb', d['hawb'])          # the house B/L stays on every line (In HAWB)
                d['house_hawb'] = d['hawb']
                gk = hbl_group[squash(d['hawb'])]
                d['hawb'], d['hawb_from'] = gks.get(gk, gk[4:]), 'ocean B/L of the house B/Ls'
        for a in awbs:
            if a.get('hawb') and squash(a['hawb']) in hbl_group:
                gk = hbl_group[squash(a['hawb'])]
                a['house_hawb'] = a['hawb']
                a['hawb'] = gks.get(gk, gk[4:])
                if a.get('cargo', {}).get('hawb'):
                    a['cargo'].pop('hawb')
            elif len(a.get('house_list') or []) > 1:
                gk = 'OBL:' + (squash(a.get('mawb') or '') or squash(a['house_list'][0]))
                a['hawb'] = gks.get(gk, gk[4:])
    # one air waybill / B/L and invoices without HAWB -> they belong to it
    awb_h = [a['hawb'] for a in awbs if a.get('hawb')]
    if len(set(awb_h)) == 1 and not any(d.get('hawb') for d in invs):
        for d in invs + pls:
            if not d.get('hawb'):
                d['hawb'], d['hawb_from'] = awb_h[0], 'air waybill / B/L'
    # invoices without HAWB that one packing list names together -> same group
    group_of = {}
    for p in pls:
        refs = [squash(r) for r in p.get('invoice_refs', [])]
        if len(refs) > 1:
            for r in refs:
                group_of[r] = 'PL:' + p['file']
    one_pack = not any(a.get('hawb') for a in awbs) and not any(d.get('hawb') for d in invs)
    for d in invs:
        key = d['hawb'] or group_of.get(squash(d['invoice_no'])) or (
            # no HAWB / B/L anywhere: invoices of one supplier to one consignee = one shipment (one permit)
            ('PACK:%s>%s' % (squash(d.get('supplier') or d.get('invoice_from') or '')[:24], squash(d.get('ship_to') or '')[:24]))
            if one_pack and (d.get('supplier') or d.get('invoice_from')) else ('INV:' + (d['invoice_no'] or d['file'])))
        sh = ships.get(key)
        if not sh:
            sh = ships[key] = dict(key=key, hawb=d['hawb'], hawb_src=('Invoice' if not d.get('hawb_from') else d['hawb_from'])
                                   if d['hawb'] else '', invs=[], pls=[], issues=[], overrides={}, include=True,
                                   direction=d['direction'], party='', party_src='', party_uen='', party_uen_src='',
                                   dest='', dest_src='', awbs=[])
        sh['invs'].append(d)
        if d['direction'] != sh['direction']:
            sh['issues'].append(('MISMATCH', '%s is %s but other invoices of this HAWB are %s'
                                 % (d['invoice_no'], d['direction'], sh['direction'])))
        for n in d['notes']:
            sh['issues'].append(('ERROR' if 'no item lines' in n else 'CHECK', '%s: %s' % (d['file'], n)))
    for p in pls:
        refs = {squash(r) for r in p.get('invoice_refs', [])}
        hit = [sh for sh in ships.values() if p.get('hawb') and squash(sh['hawb']) == squash(p['hawb'])] or \
              [sh for sh in ships.values() if refs & {squash(i['invoice_no']) for i in sh['invs']}] or \
              [sh for sh in ships.values() if p.get('delivery') and any(i.get('delivery') == p['delivery'] for i in sh['invs'])]
        if not hit and len(ships) == 1:
            hit = list(ships.values())
        if hit:
            hit[0]['pls'].append(p)
            for n in p['notes']:
                hit[0]['issues'].append(('CHECK', '%s: %s' % (p['file'], n)))
        else:
            gi.append((p['file'], 'CHECK', 'Packing list could not be matched to an invoice (no HAWB / invoice no.)'))
    for a in awbs:
        hit = [sh for sh in ships.values() if a['hawb'] and squash(sh['hawb']) == squash(a['hawb'])]
        if not hit and len(ships) == 1:
            hit = list(ships.values())
            if a['hawb']:
                hit[0]['issues'].append(('MISMATCH', 'Air waybill / B/L %s: HAWB %s, invoice HAWB %s'
                                         % (a['file'], a['hawb'], hit[0]['hawb'] or '-')))
        if hit:
            hit[0]['awbs'].append(a)
        else:
            gi.append((a['file'], 'CHECK', 'Air waybill / B/L HAWB %s does not match any invoice' % (a['hawb'] or '?')))
    for sh in ships.values():
        if sh['pls']:                      # packing list added -> "add the packing list" notes are done
            sh['issues'] = [x for x in sh['issues'] if 'add the Taiyo Yuden packing list' not in x[1]]
            if any(p.get('total_nw') or any(l.get('nw') for l in p['lines']) for p in sh['pls']):
                sh['issues'] = [x for x in sh['issues'] if 'no net weight on the invoice' not in x[1]]
        nos = [i['invoice_no'] for i in sh['invs']]
        dup = [n for n, c in collections.Counter(nos).items() if c > 1 and n]
        for n in dup:
            sh['issues'].append(('MISMATCH', 'Invoice %s added more than once (same PDF twice?)' % n))
    return ships, gi


def ship_label(sh):
    return sh['hawb'] or ', '.join(i['invoice_no'] for i in sh['invs'])


PERMIT_DIRECTION = {'OD': 'EXPORT', 'OT': 'EXPORT', 'IG': 'IMPORT', 'ME': 'IMPORT', 'IT': 'IMPORT', 'IM': 'IMPORT', 'ID': 'IMPORT'}


def _party_default(sh, ctx):
    """Importer / exporter = the Singapore company of the documents (also when the permit type turned the direction)."""
    i0 = sh['invs'][0]
    if (sh.get('doc_direction') or sh['direction']) == 'IMPORT':
        name = i0['ship_to'] or i0['invoice_to']
    else:
        name = i0['invoice_from'] or i0['ship_from']
    return name.upper()


CUST_NO_CODE = {'', 'NA', 'N/A', '#N/A', 'TBA', 'NIL', 'NONE', '-', 'NO', 'N'}


def cust_lookup(cust, r):
    """Customer product list hit for an invoice line: the item no., else a product code found in the line text
    ("D14289560" also when printed "D 14289560"). -> (key, {hs, casc, desc, file})"""
    for k in (r.get('item'), r.get('cust_part'), r.get('maker_part')):
        k = squash(k)
        for k2 in (k, k.lstrip('0')):
            if k2 and k2 in cust:
                return k2, cust[k2]
    dsc = str(r.get('inv_desc') or r.get('desc') or '').upper()
    dh = re.sub(r'\D', '', str(r.get('doc_hs') or ''))
    m = re.match(r'\s*([A-Z]{1,2})\b', dsc)
    if dh and m and m.group(1) + dh in cust:           # "D" + "14289560" read apart (item code taken as HS)
        return m.group(1) + dh, cust[m.group(1) + dh]
    txt = re.sub(r'\b([A-Z]{1,2})\s+(\d{5,})\b', r'\1\2', ' '.join((str(r.get('item') or ''), dsc)).upper())
    for tok in re.findall(r'[A-Z0-9][A-Z0-9-]{3,}', txt):      # a code with letters + digits in the line text
        k = squash(tok)
        if len(k) >= 6 and re.search(r'\d', k) and re.search(r'[A-Z]', k) and k in cust:
            return k, cust[k]
    return '', {}


def resolve(sh, ctx):
    """(Re)build the items of one shipment.
    ctx: master, hs_list, countries, default_uom, blank_hs_not_in_list, desc_template, desc_max, brand,
         cargo{}, party_name, party_uen, uen_book{}, consolidate ('none'|'item'|'hs_coo'), direction_mode"""
    master, hsl = ctx.get('master', {}), ctx.get('hs_list', {})
    issues = list(sh['issues'])
    box = ctx.get('cargo') or {}
    cg = {}
    ci_cons = next((a['cargo']['party_name'] for a in sh.get('awbs', []) if a['cargo'].get('permit_type')
                    and a['cargo'].get('party_name')), '')
    for a in sorted(sh.get('awbs', []), key=lambda a: -a.get('priority', 0)):
        for k, v in a['cargo'].items():
            cg.setdefault(k, v)
        for n in a['notes']:
            issues.append(('CHECK', '%s: %s' % (a['file'], n)))
        if a['cargo'].get('date_note'):
            issues.append(('CHECK', '%s %s %s' % ('Departure Date' if sh.get('direction') == 'EXPORT' else 'Arrival Date',
                                                  a['cargo'].get('date', ''), a['cargo']['date_note'])))
        if a.get('awb_hs'):
            issues.append(('CHECK', '%s shows one HS %s for the whole shipment - item HS codes from the invoice used'
                           % ('B/L / customs instructions' if cg.get('mode') == 'Sea' or a['cargo'].get('mode') == 'Sea' else 'Air waybill', a['awb_hs'])))
        imp = ci_cons or _party_default(sh, ctx)
        if a.get('consignee') and sh['direction'] == 'IMPORT' and squash(a['consignee'])[:20] != squash(imp)[:20]:
            issues.append(('CHECK', 'Air waybill consignee "%s" vs importer "%s"' % (a['consignee'], imp)))
    if not cg.get('mode'):                 # transport mode the invoice reader worked out (e.g. SG -> MY = Road)
        tm = next((i.get('transport') for i in sh['invs'] if i.get('transport') in ('Sea', 'Air', 'Road', 'Rail')), '')
        if tm:
            cg['mode'] = tm
    if (cg.get('mode') == 'Air' or cg.get('flight')) and cg.get('flight') and not cg.get('transport_id'):
        al = AIRLINES.get(re.sub(r'\d.*$', '', cg['flight'])[:3]) or AIRLINES.get(cg['flight'][:2])
        if al:
            cg['transport_id'] = al
            issues.append(('CHECK', '%s Transport Identifier %s from flight %s - check'
                           % ('Out' if sh.get('direction') == 'EXPORT' else 'In', al, cg['flight'])))
    for side in ('loading', 'discharge'):  # port name only -> UN/LOCODE from the sea port table
        if cg.get(side + '_name') and not cg.get(side + '_code') and cg.get('mode', 'Sea') == 'Sea':
            code = sea_port_code(cg[side + '_name'])
            if code:
                cg[side + '_code'] = code
                issues.append(('CHECK', '%s port code %s from the port name %s on the B/L - check'
                               % (side.title(), code, cg[side + '_name'])))
    awb_cg = dict(cg)
    cg.update(box)                        # cargo details box (typed / pasted) wins over the air waybill
    sh['cargo'] = cg
    # customer rules of this shipment (Customer rules tab: ticked = used)
    custs = [c for i in sh['invs'] for c in (i.get('customers') or [])]
    team = str(ctx.get('team') or '').strip().upper()
    rules_here = [r['id'] for r in CUSTOMER_RULES if r['customer'] in custs and (not team or team in r.get('teams', ()))]
    sh['team'] = team
    off = set(ctx.get('rules_off') or [])
    sh['rules'] = [dict(id=r, on=r not in off) for r in rules_here]
    rule_on = lambda rid: rid in rules_here and rid not in off
    sh['rule_on'] = {r: rule_on(r) for r in rules_here}
    dm = (ctx.get('direction_mode') or 'Auto').upper()
    i0 = sh['invs'][0]
    if dm in ('IMPORT', 'EXPORT'):
        sh['direction'] = dm
    elif rule_on('sandvik_export') or rule_on('dow_export'):
        sh['direction'] = 'EXPORT'
    elif ('sandvik_export' in rules_here or 'dow_export' in rules_here) and i0.get('ship_to_country') == 'SG':
        sh['direction'] = 'IMPORT'           # export rule unticked: direction from the documents (goods to Singapore)
    else:
        sh['direction'] = i0['direction']
    sh['doc_direction'] = sh['direction']
    pt_ = (ctx.get('permit_type') or cg.get('permit_type') or '').upper()[:2]
    if dm not in ('IMPORT', 'EXPORT') and PERMIT_DIRECTION.get(pt_) and PERMIT_DIRECTION[pt_] != sh['direction']:
        sh['direction'] = PERMIT_DIRECTION[pt_]   # the permit type decides: OD / OT = export, IG / ME / IT / IM / ID = import
        issues.append(('MISMATCH', 'Permit type %s = %s permit, but the documents look like an %s (%s) - the permit is made as '
                                   '%s (%s fields only). If this shipment is an %s, change the permit type on the job.'
                       % (pt_, sh['direction'], sh['doc_direction'], 'goods to Singapore' if sh['doc_direction'] == 'IMPORT'
                          else 'goods from Singapore', sh['direction'],
                          'outward transport, exporter, Out HAWB' if sh['direction'] == 'EXPORT' else 'inward transport, importer, In HAWB',
                          sh['doc_direction'])))
        # the documents describe the other leg (e.g. the import voyage): its transport details are not this permit's
        leg = [k for k in ('mawb', 'flight', 'transport_id', 'date', 'eta', 'etd', 'loading_code', 'loading_name', 'discharge_code',
                           'discharge_name', 'release', 'receipt', 'containers', 'container_info', 'carrier_agent',
                           'carrier_agent_uen', 'vessel_type') if cg.get(k) and k not in box]
        for k in leg:
            cg.pop(k, None)
        if leg:
            issues.append(('MISSING', 'Permit type %s: the %s transport of the documents (%s) is not used - type the %s details '
                                      '(%s) on the Cargo / Header tabs or in Cargo details' % (
                               pt_, 'inward' if sh['direction'] == 'EXPORT' else 'outward',
                               ', '.join(sorted({k.split('_')[0] for k in leg})),
                               'outward' if sh['direction'] == 'EXPORT' else 'inward',
                               'vessel / flight, departure date, ports, OBL / MAWB, containers, release / receipt'
                               if sh['direction'] == 'EXPORT' else 'vessel / flight, arrival date, loading port, OBL / MAWB, containers')))
        sh['dir_flipped'] = True
    exp = sh['direction'] == 'EXPORT'
    if not box.get('date'):               # documents: departure date for an export, arrival date for an import
        if exp and cg.get('etd'):
            cg['date'] = cg['etd']
        elif not exp and cg.get('eta'):
            cg['date'] = cg['eta']
        elif exp and cg.get('date') and cg.get('eta') == cg['date']:
            cg.pop('date')
            issues.append(('CHECK', 'Departure Date: the documents give only the arrival date (ETA %s) - type the '
                                    'departure / on-board date' % cg['eta']))
    # ---- party (importer for IMPORT, exporter for EXPORT)
    if sh.get('party_src') != 'Entered':
        ci_party = (cg.get('party_name') or '').upper() if cg.get('permit_type') and cg.get('party_name') else ''
        sh['party'] = (ctx.get('party_name') or '').upper() or ci_party or _party_default(sh, ctx)
        sh['party_src'] = 'Tab 3' if ctx.get('party_name') else 'Customs instructions' if ci_party else 'Invoice'
    if sh.get('party_uen_src') != 'Entered':
        book = {k.upper(): v for k, v in (ctx.get('uen_book') or {}).items()}
        doc_uen = next((i.get('uen') for i in sh['invs'] if i.get('uen')), '')
        uen = ctx.get('party_uen') or cg.get('uen') or doc_uen or book.get(sh['party'].upper(), '')
        sh['party_uen'], sh['party_uen_src'] = uen, ('Tab 3' if ctx.get('party_uen') else 'Cargo details'
                                                      if cg.get('uen') else 'Invoice' if doc_uen else 'Saved' if uen else '')
    un = (ctx.get('uen_names') or {}).get(str(sh.get('party_uen') or '').upper()) if ctx.get('uen_names_on', True) else None
    sh['party_address'] = []
    if un and un.get('name') and sh.get('party_src') not in ('Entered', 'Tab 3'):
        if squash(un['name']) != squash(sh.get('party') or ''):
            issues.append(('CHECK', '%s name from the UEN register (ACRA) for %s: "%s" (documents: "%s")'
                           % ('Exporter' if exp else 'Importer', sh['party_uen'], un['name'], sh.get('party') or '-')))
        sh['party'], sh['party_src'] = un['name'], 'ACRA register (UEN %s)' % sh['party_uen']
        sh['party_address'] = list(un.get('address') or [])
    if not sh['party_uen']:
        issues.append(('MISSING', '%s UEN not in the documents - type it on tab 3 (remembered) or tab 2'
                       % ('Exporter' if exp else 'Importer')))
    if exp and sh.get('dest_src') != 'Entered':
        sh['dest'] = sh['invs'][0]['ship_to_country']
        if sh['dest'] == 'SG':                 # an export goes out of Singapore - SG is never the final destination
            sh['dest'] = ''
            issues.append(('MISSING', 'Final destination country: the documents ship to Singapore - type the export '
                                      'destination on tab 2'))
        elif not sh['dest']:
            issues.append(('MISSING', 'Final destination country not found in Ship To - type it on tab 2'))
    # ---- HAWB
    if sh.get('hawb_src') != 'Entered':
        if not sh['hawb'] and cg.get('hawb'):
            sh['hawb'], sh['hawb_src'] = cg['hawb'], 'Cargo details'
        elif sh['hawb'] and cg.get('hawb') and squash(cg['hawb']) != squash(sh['hawb']):
            issues.append(('MISMATCH', 'HAWB: invoice Carrier Reference No. %s, cargo details %s (invoice used)'
                           % (sh['hawb'], cg['hawb'])))
    # ---- weights / packages
    pls = sh.get('pls', [])
    gw_inv = sum(i['total_gw'] or 0 for i in sh['invs']) or sum(p['total_gw'] or 0 for p in pls) or None
    nw_inv = sum(i['total_nw'] or 0 for i in sh['invs']) or sum(p['total_nw'] or 0 for p in pls) or None
    boxes = sum(i['boxes'] or 0 for i in sh['invs']) or sum(p['packages'] or 0 for p in pls) or None
    if not boxes and pls:
        boxes = None
    gw_pl = sum(p['total_gw'] or 0 for p in pls) or None
    inv_del = [l.get('order') for i in sh['invs'] for l in i['lines'] if l.get('order')]
    pl_del = {p.get('delivery') for p in pls if p.get('delivery')}
    partial = bool(pl_del) and bool(inv_del) and not set(inv_del) <= pl_del and pl_del <= set(inv_del)
    if partial:
        miss_d = [x for x in dict.fromkeys(inv_del) if x not in pl_del]
        issues.append(('CHECK', 'Packing lists / delivery notes for %d of %d deliveries - not added: %s (weights from the invoice '
                                'for those)' % (len(pl_del), len(set(inv_del)), ', '.join(miss_d[:20]))))
    if gw_pl and not partial and sum(i['total_gw'] or 0 for i in sh['invs']) and abs(gw_pl - sum(i['total_gw'] or 0 for i in sh['invs'])) > max(0.05, gw_pl * 0.00005):
        issues.append(('MISMATCH', 'Gross weight: invoice %s KG, packing list %s KG' % (
            fmt_thou(sum(i['total_gw'] or 0 for i in sh['invs'])), fmt_thou(gw_pl))))
    for p in pls:
        for c in ([] if sh.get('dir_flipped') and 'containers' not in box else p.get('containers', [])):
            cg.setdefault('containers', [])
            if c not in cg['containers']:
                cg['containers'] = cg['containers'] + [c]
    sh['gw_doc'], sh['boxes_doc'] = gw_inv, boxes
    # gross weight: typed on tab 2 > invoice > air waybill > cargo details box
    one = len(ctx.get('ships_n', [1])) <= 1
    if sh.get('gw_override'):
        sh['total_gw'], sh['gw_src'] = sh['gw_override'], 'typed on tab 2'
    elif box.get('instructions') and box.get('gw_kg') and one:
        sh['total_gw'], sh['gw_src'] = box['gw_kg'], 'instructions (pasted)'
        if gw_inv and abs(gw_inv - box['gw_kg']) > 0.5:
            issues.append(('CHECK', 'Gross weight %s KG from the instructions%s - documents show %s KG'
                           % (fmt_thou(box['gw_kg']), ' (with pallets)' if box.get('gw_with_pallets') else '', fmt_thou(gw_inv))))
    elif gw_inv:
        sh['total_gw'], sh['gw_src'] = gw_inv, 'invoice' if sum(i['total_gw'] or 0 for i in sh['invs']) else 'packing list'
    elif awb_cg.get('gw_kg'):
        sh['total_gw'], sh['gw_src'] = awb_cg['gw_kg'], 'air waybill'
    elif cg.get('gw_kg'):
        sh['total_gw'], sh['gw_src'] = cg['gw_kg'], 'cargo details box'
    else:
        sh['total_gw'], sh['gw_src'] = None, ''
    sh['packages'] = sh.get('pkg_override') or (box.get('packages_qty') if box.get('instructions') and one else None) \
        or boxes or cg.get('packages_qty')
    for src, c in (('air waybill', awb_cg), ('cargo details', box if len(ctx.get('ships_n', [1])) <= 1 else {})):
        if gw_inv and c.get('gw_kg') and abs(gw_inv - c['gw_kg']) > max(0.05, c['gw_kg'] * 0.00005):
            issues.append(('MISMATCH', 'Gross weight: invoice %s KG, %s %s KG' % (fmt_thou(gw_inv), src, fmt_thou(c['gw_kg']))))
        cu = (c.get('packages_unit') or '').upper()
        if boxes and c.get('packages_qty') and boxes != c['packages_qty'] and cu not in ('UNT', 'UNIT', 'CNTR', 'CONTAINER'):
            issues.append(('MISMATCH', 'Packages: invoice %s boxes, %s %s %s'
                           % (boxes, src, c['packages_qty'], c.get('packages_unit', ''))))
    sh['freight_amount'], sh['freight_currency'] = None, ''
    if cg.get('freight_amount') and not exp:
        inc = next((i['incoterm'] for i in sh['invs'] if i.get('incoterm')), '')
        if inc in ('', 'EXW', 'FCA', 'FOB', 'FAS'):
            sh['freight_amount'], sh['freight_currency'] = cg['freight_amount'], cg.get('freight_currency', '')
            issues.append(('CHECK', 'Freight %s %s from the customs instructions (%s) - written as Freight Charge (invoice %s)'
                           % (sh['freight_currency'], fmt_plain(sh['freight_amount']), cg.get('freight_note', ''), inc or 'no incoterm')))
        else:
            issues.append(('CHECK', 'Freight %s %s in the customs instructions not written - invoice term %s includes freight'
                           % (cg.get('freight_currency', ''), fmt_plain(cg['freight_amount']), inc)))
    if not sh['total_gw']:
        issues.append(('MISSING', 'Total Gross Weight not found (invoice / air waybill) - type it on tab 2 "Gross Wt (KG)"'))
    if not sh['packages']:
        issues.append(('MISSING', 'Number of boxes / packages not on the invoice - type it on tab 2'))
    cinfo = {}
    for c in cg.get('container_info') or []:
        if isinstance(c, dict) and c.get('no'):
            cinfo.setdefault(c['no'], c)
    sh['containers'] = [dict(no=c, size=(cinfo.get(c) or {}).get('size', ''), size_text=(cinfo.get(c) or {}).get('size_text', ''),
                             tne=(cinfo.get(c) or {}).get('tne'), seal=(cinfo.get(c) or {}).get('seal', ''))
                        for c in cg.get('containers', [])]
    dec = 0                   # container weight (TNE) always a round figure: .5 and above up (5.5 -> 6), below down (5.49 -> 5)
    gw_c = sh.get('total_gw')
    for c in sh['containers']:
        raw_t, how = c.get('tne'), 'B/L / customs instructions'
        if raw_t not in (None, '') and gw_c and len(sh['containers']) == 1 and abs(float(raw_t) * 1000 - gw_c) > 0.5:
            raw_t, how = gw_c / 1000.0, 'gross weight %s KG / 1000' % fmt_thou(gw_c)
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
    sh['incoterm'] = next((i['incoterm'] for i in sh['invs'] if i['incoterm']), '')
    curs = {i['currency'] for i in sh['invs'] if i['currency']}
    if len(curs) > 1:
        issues.append(('CHECK', 'More than one invoice currency: %s - the Invoice sheet gets one invoice per currency (items linked by invoice number)' % ', '.join(sorted(curs))))
    incs = {i['incoterm'] for i in sh['invs'] if i['incoterm']}
    if len(incs) > 1:
        issues.append(('CHECK', 'More than one Incoterm: %s (first used)' % ', '.join(sorted(incs))))
    # ---- raw lines + document checks
    raw = []
    hs_by_desc, hs_by_order = {}, {}            # HS printed on the B/L / packing list for a product (invoice has none)
    for a_ in sh.get('awbs', []):
        for dsc, h_ in (a_.get('hs_by_desc') or {}).items():
            hs_by_desc.setdefault(squash(dsc), (h_, 'B/L'))
    for p_ in sh.get('pls', []):
        for l_ in p_['lines']:
            if l_.get('hs'):
                hs_by_desc.setdefault(squash(l_['desc']), (l_['hs'], 'Packing list'))
                if l_.get('order'):
                    hs_by_order.setdefault(squash(l_['order']), (l_['hs'], 'Packing list'))
    for inv in sh['invs']:
        if inv.get('ocr'):
            issues.append(('CHECK', 'Invoice %s: %s' % (inv['invoice_no'] or inv['file'], inv.get('notes_ocr') or 'scanned - read by OCR, check the values')))
        for kk in ('notes_box', 'notes_total'):
            if inv.get(kk):
                issues.append(('CHECK', 'Invoice %s: %s' % (inv['invoice_no'], inv[kk])))
        for kk in ('notes_ai', 'notes_coo'):
            if inv.get(kk):
                issues.append(('CHECK', 'Invoice %s: %s' % (inv['invoice_no'] or inv['file'], inv[kk])))
        if inv.get('hawb_from_name'):
            issues.append(('CHECK', 'Invoice %s: HAWB %s taken from the file name (not readable on the scan)'
                           % (inv['invoice_no'], inv['hawb'])))
        line_sum = round(sum(l['amount'] or 0 for l in inv['lines']) + (inv.get('charges') or 0), 2)
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
            raw.append(dict(key='%s|%s' % (inv['invoice_no'], l['line']), line=l['line'], item=l['item'],
                            inv_desc=l['desc'], qty=l['qty'], uom=l['uom'], amount=l['amount'], price=l['price'],
                            invoice=inv['invoice_no'], currency=inv['currency'] or l['currency'], doc_hs=l['hs'],
                            doc_coo=l['coo'] or (inv.get('coo_global') or ''), eccn=l['eccn'], order=l['order'],
                            batches=list(l['lots']), boxes=list(l['boxes']), src=inv['file'],
                            doc_nw=l.get('nw'), doc_gw=l.get('gw'), desc_rule=inv.get('desc_rule', ''),
                            pl_first=bool(inv.get('pl_first')), brand=(l.get('brand') or inv.get('brand') or '').upper(),
                            our_ref=l.get('our_ref') or inv.get('our_ref') or '', shipment=inv.get('shipment_no') or '',
                            sinv=l.get('sales_inv') or '', bl_hs=l.get('bl_hs') or '', charges=l.get('charges'),
                            line_hawb=l.get('hawb') or inv.get('hawb') or '', ltr=l.get('ltr'),
                            liquor=l.get('liquor'), eu_hs=l.get('eu_hs') or ''))
            if not l['hs'] and inv.get('hs_global'):
                raw[-1]['doc_hs'] = inv['hs_global']
            if inv.get('pl_first'):            # supplier rule: HS from the packing list / delivery note of that delivery
                hd_ = hs_by_order.get(squash(l.get('order') or ''))
                if hd_:
                    raw[-1]['doc_hs'], raw[-1]['hs_from'] = hd_[0], 'Packing list (delivery %s)' % l.get('order')
                elif not raw[-1]['doc_hs']:
                    hd_ = hs_by_desc.get(squash(l['desc']))
                    if hd_:
                        raw[-1]['doc_hs'], raw[-1]['hs_from'] = hd_[0], hd_[1] + ' (delivery note not added)'
                        issues.append(('CHECK', 'Delivery %s (%s): delivery note not added - HS %s from the %s'
                                       % (l.get('order'), l['desc'], hd_[0], hd_[1])))
            elif not raw[-1]['doc_hs']:
                hd_ = hs_by_desc.get(squash(l['desc'])) or hs_by_order.get(squash(l.get('order') or ''))
                if hd_:
                    raw[-1]['doc_hs'], raw[-1]['hs_from'] = hd_
        if inv.get('charges') and cmode == 'last' and inv['lines'] and raw and raw[-1]['invoice'] == inv['invoice_no']:
            raw[-1]['amount'] = round((raw[-1]['amount'] or 0) + inv['charges'], 2)     # charges in the last item
            raw[-1]['charge_add'] = inv['charges']
    sh['charges'] = round(sum(i.get('charges') or 0 for i in sh['invs']), 2)
    sh['charges_in_items'] = round(sum(r.get('charge_add') or 0 for r in raw), 2)
    # ---- weights: invoice line NW > packing list (by item) > total NW split by value > GW x 85 % split by value
    gw = sh['total_gw']
    for r in raw:
        r['nw'], r['gw'], r['nw_src'] = None, None, ''
        if r.get('doc_nw') is not None:
            r['nw'], r['gw'], r['nw_src'] = r['doc_nw'], r.get('doc_gw'), 'Invoice line'
    plm = collections.OrderedDict()
    for p in pls:
        for l in p['lines']:
            k = squash(l['item']) or squash(l['desc'])[:25]
            if not k:
                continue
            m = plm.setdefault(k, dict(qty=0.0, nw=0.0, gw=0.0, has_nw=False, has_gw=False))
            m['qty'] += l['qty'] or 0
            if l.get('nw') is not None:
                m['nw'] += l['nw']
                m['has_nw'] = True
            if l.get('gw') is not None:
                m['gw'] += l['gw']
                m['has_gw'] = True
    for r in raw:
        if r.get('pl_first') and r.get('order') and not any(squash(p.get('delivery') or '') == squash(r['order']) for p in pls):
            issues.append(('CHECK', 'Delivery %s (%s): delivery note not added - net weight %s KG from the invoice'
                           % (r['order'], r['item'], fmt_thou(r.get('doc_nw')))))
    plref = {}
    for p in pls:                          # packing list rows that name the invoice line (PO line no.) -> exact weights
        for l in p['lines']:
            if l.get('order') and (l.get('nw') is not None or l.get('gw') is not None):
                plref.setdefault(squash(l['order']) + '|' + squash(l['item']), l)
    if plref:
        by_order = collections.defaultdict(list)
        for k_, l_ in plref.items():
            by_order[k_.split('|')[0]].append(l_)
        for r in raw:
            l = plref.get(squash(r.get('order') or '') + '|' + squash(r['item']))
            if not l and r.get('order') and len(by_order.get(squash(r['order']), [])) == 1:
                l = by_order[squash(r['order'])][0]      # same delivery / order no., different item code
            if r.get('pl_first') and l and l.get('nw') is not None:     # supplier rule: net weight from the packing list
                if r.get('doc_nw') is not None and abs(r['doc_nw'] - l['nw']) > 0.5:
                    issues.append(('MISMATCH', 'Delivery %s (%s) net weight: delivery note %s KG, invoice %s KG (delivery '
                                               'note used)' % (r['order'], r['item'], fmt_thou(l['nw']), fmt_thou(r['doc_nw']))))
                r['nw'], r['gw'], r['nw_src'] = l['nw'], l.get('gw') or r.get('gw'), 'Packing list (delivery %s)' % r['order']
            elif r.get('pl_first') and r['nw'] is not None:
                r['nw_src'] = 'Invoice line (delivery note not added)'
            elif r['nw'] is None and l:
                r['nw'], r['gw'], r['nw_src'] = l.get('nw'), l.get('gw'), 'Packing list (PO line)'
                if l.get('qty') and r['qty'] and abs(l['qty'] - r['qty']) > 0.0005:
                    issues.append(('MISMATCH', 'Item %s PO line %s quantity: invoice %s, packing list %s'
                                   % (r['item'], r['order'], fmt_plain(r['qty']), fmt_plain(l['qty']))))
    pl_gw_only = [0]
    if plm:
        by_item = collections.defaultdict(list)
        for r in raw:
            if r['nw'] is None:
                by_item[squash(r['item']) or squash(r['inv_desc'])[:25]].append(r)
        used = set()
        for k, rs in by_item.items():
            m = plm.get(k)
            if not m or not (m['has_nw'] or m['has_gw']):
                continue
            used.add(k)
            tq = sum(r['qty'] or 0 for r in rs)
            for r in rs:
                share = (r['qty'] or 0) / tq if tq else 1.0 / len(rs)
                if m['has_nw']:
                    r['nw'], r['nw_src'] = m['nw'] * share, 'Packing list'
                else:                               # packing list shows box gross weight only -> GW x 85 %
                    r['nw'], r['nw_src'] = m['gw'] * share * 0.85, 'Packing list box GW x 85%'
                    pl_gw_only[0] += 1
                r['gw'] = m['gw'] * share if m['has_gw'] else None
            if tq and m['qty'] and abs(tq - m['qty']) > 0.0005:
                issues.append(('MISMATCH', 'Item %s quantity: invoice %s, packing list %s' % (rs[0]['item'], fmt_plain(tq), fmt_plain(m['qty']))))
        if pl_gw_only[0]:
            issues.append(('NET WEIGHT', 'Net weight not on the documents - packing list box gross weights x 85%% for %d line(s)'
                           % pl_gw_only[0]))
        miss_pl = [k for k in plm if k not in used and k not in by_item and not any(
            (squash(r['item']) or squash(r['inv_desc'])[:25]) == k for r in raw)]
        if miss_pl and len(miss_pl) < len(plm):
            issues.append(('CHECK', 'Packing list items not on the invoice: %s' % ', '.join(miss_pl[:15])))
    miss = [r for r in raw if r['nw'] is None]
    if miss:
        known_nw = sum(r['nw'] for r in raw if r['nw'] is not None)
        known_gw = sum(r['gw'] or 0 for r in raw if r['nw'] is not None)
        tv = sum(r['amount'] or 0 for r in miss)
        if nw_inv:
            rest_nw, how = nw_inv - known_nw, 'total NW (%s), split by value' % ('invoice' if sum(i['total_nw'] or 0 for i in sh['invs']) else 'packing list')
        elif gw:
            rest_nw, how = gw * 0.85 - known_nw, '85% of GW, split by value'
        else:
            rest_nw, how = None, ''
        rest_gw = (gw - known_gw) if gw else None
        if rest_nw is not None and rest_nw > 0 and tv > 0:
            for r in miss:
                f = (r['amount'] or 0) / tv
                r['nw'], r['gw'], r['nw_src'] = rest_nw * f, (rest_gw * f if rest_gw else None), how
            if not nw_inv:
                issues.append(('NET WEIGHT', 'Net weight not on the documents for %d line(s) - gross weight %s KG (%s) x 85%% '
                               '= %s KG%s, split by item value' % (len(miss), fmt_thou(gw), sh['gw_src'],
                                                                  fmt_thou(round(gw * 0.85, 3)),
                                                                  (' less %s KG known' % fmt_thou(round(known_nw, 3))) if known_nw else '')))
        else:
            issues.append(('NET WEIGHT', 'Net Weight unavailable and %s - weights left blank'
                           % ('Item Value unavailable for weight allocation' if gw and tv <= 0 else
                              'net weight already used up by the known lines' if rest_nw is not None and rest_nw <= 0 else
                              'no gross weight in the invoice / packing list / air waybill - type the Gross Wt on tab 2 '
                              '(NW = GW x 85 %, split by item value)')))
    nw_all = sum(r['nw'] or 0 for r in raw)
    if nw_inv and raw and all(r['nw'] is not None for r in raw) and abs(nw_all - nw_inv) > max(0.05, nw_inv * 0.01):
        issues.append(('MISMATCH', 'Item net weights add up to %s KG, document total NW %s KG' % (fmt_thou(nw_all), fmt_thou(nw_inv))))
    # ---- per line: master / HS / COO / UOM
    lines = []
    no_ctl = []
    auto_map = collections.OrderedDict()
    hs_missing = collections.OrderedDict()
    rpn_hits = []
    ptxt = ' '.join(str(x or '') for i in sh['invs'] for x in (i.get('supplier'), i.get('invoice_from'), i.get('ship_to'),
                                                              i.get('invoice_to'))) + ' ' + str(sh.get('party') or '')
    hsd_here = [(nm_, ix_) for nm_, rx_, ix_ in ctx.get('hs_desc') or [] if rx_ and re.search(rx_, ptxt, re.I)]
    hsd_near = []
    hs_kw = []
    rs_nomaster = collections.OrderedDict()
    cust = ctx.get('cust_master') or {}
    sel = str(ctx.get('cust_list') or '').strip()
    if sel and (rule_on('ddi_vendor') or rule_on('givaudan_hs') or not ({'ddi_vendor', 'givaudan_hs'} & set(rules_here))):
        if sel == '-':
            cust = {}
        else:
            fn_, _, sh_ = sel.partition('::')
            cust = {k: v for k, v in ((ctx.get('cust_lists') or {}).get(fn_) or {}).items() if not sh_ or v.get('sheet', '').strip() == sh_.strip()}
            if not cust:
                issues.append(('MISSING', 'Vendor product list "%s" not found (Settings > Product masters) - no list used' % sel))
    cust_miss = []
    mci = re.search(r'\bHS\s*(?:CODE|NO\.?)?\s*[:=\-]?\s*(\d{4}[ .]?\d{2}(?:[ .]?\d{2})?)\b',
                    ' '.join(str(cg.get(k) or '') for k in ('special', 'permit_remarks')), re.I)
    ci_hs = h8(re.sub(r'\D', '', mci.group(1)))[:8] if mci else ''
    if ci_hs and len(re.sub(r'\D', '', mci.group(1))) == 6:
        ci_hs = re.sub(r'\D', '', mci.group(1)) + '00'
    ci_hs_diff = []
    ahs = {re.sub(r'\D', '', a.get('awb_hs') or '') for a in sh.get('awbs', []) if a.get('awb_hs')}
    ahs = {h for h in ahs if len(h) >= 6}
    awb_hs1 = ''
    if len({h[:6] for h in ahs}) == 1:
        h_ = max(ahs, key=len)
        awb_hs1 = (h_ + '00')[:8] if len(h_) == 6 else h_[:8]
    hs_bad = collections.OrderedDict()       # HS on the documents but not in the UPDATE HS CODE list -> left empty
    liq_eu, liq_coo, liq_master = [], [], []
    def hs_ok(h):
        return bool(h) and (not hsl or bool(hsl.get(h) or hsl.get(h[:8])))
    for r in raw:
        ov = sh['overrides'].get(r['key'], {})
        ms = master.get(r['item'], {})
        ck, cm = cust_lookup(cust, r) if cust else ('', {})
        it = dict(r)
        it['desc'] = r['inv_desc'] or ms.get('desc', '') or cm.get('desc', '')
        it['cust_key'], it['cust_desc'], it['cust_item'] = ck, cm.get('desc', ''), cm.get('item_raw', '')
        it['cust_control'] = cm.get('control', '')
        if cust and not ck and (rule_on('ddi_vendor') or rule_on('givaudan_hs') or sel):
            cust_miss.append(r['item'])
        if ck and not it.get('item'):
            it['item'] = ck
        hs, hs_src = '', ''
        # HS: only a code that is in the UPDATE HS CODE list counts - invoice > product master > customer product list
        cands = []
        liq_res = None
        if r.get('liquor'):                # wine / champagne / spirits: Singapore HS + product code by the liquor rules
            import liquor_rules as LQ
            liq_res = LQ.classify(r['liquor'], product_codes)
            it['liq_res'] = liq_res
            for n_ in liq_res['notes']:
                issues.append(('CHECK', 'Liquor: ' + n_))
            if liq_res['hs']:
                cands.append((liq_res['hs'], 'Liquor rule (%s)' % liq_res['why']))
            if r.get('eu_hs'):
                liq_eu.append('%s %s' % (r['line'], r['eu_hs']))
        if ci_hs:                          # HS in the customs instructions (Special Instructions "HS CODE: 8414 5191") wins
            cands.append((ci_hs, 'Customs instructions'))
            if r['doc_hs'] and h8(r['doc_hs'])[:8] != ci_hs:
                ci_hs_diff.append('%s: invoice %s' % (r['item'] or (r.get('inv_desc') or '')[:25], h8(r['doc_hs'])[:8]))
        if r['doc_hs']:
            cands.append((h8(r['doc_hs'])[:8], r.get('hs_from') or 'Invoice'))
        if ms.get('hs'):
            cands.append((ms['hs'], 'Product master'))
        if cm.get('hs'):
            cands.append((cm['hs'], 'Customer product list %s%s (%s)' % (cm.get('file', ''), (' / ' + cm['sheet'].strip()) if cm.get('sheet') else '', ck)))
        if r.get('bl_hs'):
            cands.append((h8(r['bl_hs'])[:8], 'B/L HS-NO (same net weight)'))
        if not cands and awb_hs1:          # nothing else: the one HS of the B/L / customs instructions ("Harmonized Code")
            cands.append((awb_hs1, 'B/L / customs instructions Harmonized Code'))
        for h_, src_ in cands:
            if hs_ok(h_):
                hs, hs_src = h_, src_
                break
        if not hs and hsl and ctx.get('hs_auto', True) and 'hs' not in ov:
            for h_, src_ in cands:             # a real HS (its 6-digit subheading is in the list): Singapore code of it
                code, how = hs_auto_map(h_, hsl, r.get('inv_desc', ''))
                if code:
                    auto_map.setdefault((h_, code, how), []).append(r['item'] or ck)
                    hs, hs_src = code, src_ + ' > auto-mapped from %s' % h_
                    break
        for h_, src_ in cands:
            if not hs_ok(h_) and not (hs and hs_src.endswith('from %s' % h_)):
                hs_bad.setdefault((h_, src_.split(' (')[0]), []).append((r['key'], r['item'] or ck or (r.get('inv_desc') or '')[:30]))
        if hs and len(cands) > 1:
            for h_, src_ in cands:
                if hs_ok(h_) and h_ != hs:
                    issues.append(('MISMATCH', 'Item %s HS code: %s %s, %s %s (%s used)' % (r['item'], hs_src.split(' (')[0], hs, src_.split(' (')[0], h_, hs_src.split(' (')[0])))
                    break
        if not hs and not cands and hsd_here and 'hs' not in ov:      # customer HS list by description (e.g. PANASONIC FACTORY)
            for nm_, ix_ in hsd_here:
                mt = ix_.match(r.get('inv_desc') or it['desc'])
                if mt and hs_ok(mt[0]):
                    hs = mt[0]
                    hs_src = '%s HS list (%s: %s%s)' % (nm_, mt[1], mt[2][:40], '' if mt[1] == 'exact' else ' %.2f' % mt[3])
                    if mt[1] == 'near':
                        hsd_near.append('%s -> %s (%s)' % (r['item'], hs, mt[2][:30]))
                    break
        if not hs and not cands and 'hs' not in ov:    # no HS anywhere: HS by a word of the description (Settings, e.g. CAPACITOR)
            dsc = (r.get('inv_desc') or it['desc'] or '').upper()
            for kw, code in HS_KEYWORDS_DEFAULT if ctx.get('hs_keywords') is None else ctx['hs_keywords']:
                if kw and re.search(r'\b' + re.escape(kw.upper()) + r'S?\b', dsc) and hs_ok(h8(code)[:8]):
                    hs, hs_src = h8(code)[:8], 'Description word "%s" (Settings > HS by description word)' % kw.upper()
                    hs_kw.append('%s -> %s (%s)' % (r['item'], hs, kw.upper()))
                    break
        if cm and not ms:
            ms = dict(casc=cm.get('casc', ''), casc_src='Customer product list %s (%s)' % (cm.get('file', ''), ck))
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
            elif 'hs' in ov:
                hs_missing.setdefault(hs, []).append(r['item'])
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
        casc, casc_src = ms.get('casc', ''), ((ms.get('casc_src') or 'Product master') if ms.get('casc') else '')
        if why and not casc:
            casc, casc_src = code, 'Controlled (%s)' % ', '.join(why)
        if not ms.get('casc') and any(c[0] == 'IDAMISC' for c in product_codes(hs)):
            casc, casc_src = 'IDAMISC', 'HSCA list: HS %s has product code IDAMISC (always used)' % hs
        if rule_on('rs_casc'):            # RS: CASC from the product code master by HS (blank in the master = no CASC)
            cm = ctx.get('casc_hs') if ctx.get('casc_hs') is not None else _casc_master_here()
            if hs in cm:
                casc, casc_src = cm[hs], ('RS product code master' if cm[hs] else 'RS product code master (no code)')
            elif hs:
                casc, casc_src = '', 'Not in the RS product code master'
                rs_nomaster.setdefault(hs, []).append(r['item'])
        if liq_res is not None:            # liquor: the product code of the official description (never MISC)
            casc, casc_src = (liq_res['code'], 'Liquor rule: %s' % liq_res['code_desc']) if liq_res['code'] and hs == liq_res['hs'] \
                else ('', 'Liquor rule: no code decided - pick one')
            mpath = ctx.get('liquor_master')
            if mpath:
                import liquor_rules as LQ
                past = LQ.load_master(mpath).get(LQ.norm_name(r['liquor'].get('name')))
                if past and (hs, casc) not in past:
                    liq_master.append('%s %s: now %s / %s, past permit %s' % (r['line'], (r['liquor'].get('name') or '')[:40], hs, casc or '-',
                                                                              ', '.join('%s / %s' % x for x in sorted(past))))
            hint = r['liquor'].get('country_hint')
            if hint and coo and hint != coo:
                liq_coo.append('%s %s (%s -> %s)' % (r['line'], (r['liquor'].get('name') or '')[:35], coo, hint))
        if 'casc' in ov:
            casc, casc_src = ov['casc'].upper(), 'Edited'
        it['pcodes'] = product_codes(hs)
        if False and casc and it['pcodes'] and casc not in {c[0] for c in it['pcodes']}:   # rule: IDAMISC else MISC
            issues.append(('MISMATCH', 'Item %s HS %s: CASC %s is not a product code of this HS in the HSCA list - pick one '
                           '(CASC cell dropdown): %s' % (r['item'], hs, casc, '; '.join(
                               '%s %s' % (c[0], c[1][:40]) for c in it['pcodes'][:6]) + (' ...' if len(it['pcodes']) > 6 else ''))))
        it.update(hs=hs, hs_src=hs_src, hs_uom=uom, uom_src=uom_src, coo=coo, coo_src=coo_src,
                  casc=casc, casc_src=casc_src, controlled=bool(why), sep=bool(ov.get('sep')),
                  brand=ov.get('brand', ms.get('brand', '') or r.get('brand', '') or ctx.get('brand', '')))
        if rule_on('vf_brand') and r.get('brand') and 'brand' not in ov:
            it['brand'], it['brand_src'] = r['brand'], 'Invoice (customer rule)'
        lines.append(it)
    if hsl and not HS_LIST_INFO.get('control_column'):
        issues.append(('CHECK', 'UPDATE HS CODE list has no Control column - controlled items (CASC) not checked'))
    elif not hsl:
        issues.append(('CHECK', 'UPDATE HS CODE list not built in - controlled items (CASC) not checked'))
    if rs_nomaster:
        issues.append(('MISSING', '%d HS code(s) not in the RS product code master - CASC left blank, update the master '
                       '(Settings > Product masters): %s' % (len(rs_nomaster), ', '.join(
                           '%s (%s)' % (h, ', '.join(list(dict.fromkeys(v))[:3])) for h, v in list(rs_nomaster.items())[:20]))))
    if rule_on('rs_casc') and not (ctx.get('casc_hs') if ctx.get('casc_hs') is not None else _casc_master_here()):
        issues.append(('MISSING', 'RS product code master not found - CASC not filled (Settings > Product masters)'))
    if hs_kw:
        issues.append(('CHECK', 'No HS on the documents - HS from a word of the description for %d item(s): %s'
                       % (len(hs_kw), '; '.join(list(dict.fromkeys(hs_kw))[:12]))))
    if hsd_near:
        issues.append(('CHECK', '%d item(s): HS from the NEAREST description of the %s HS list (not an exact match) - check: %s'
                       % (len(hsd_near), hsd_here[0][0] if hsd_here else '', '; '.join(hsd_near[:12]))))
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
    if ci_hs:
        issues = [x for x in issues if 'shows one HS' not in x[1]]
        issues.append(('CHECK', 'HS %s from the customs instructions used for the items ("%s")' % (ci_hs, mci.group(0).strip())))
    if ci_hs_diff:
        issues.append(('MISMATCH', 'HS on the invoice differs from the customs instructions HS %s - %s (customs instructions '
                                   'used, check)' % (ci_hs, '; '.join(ci_hs_diff[:15]))))
    if cust_miss:
        issues.append(('MISSING', 'Material not in the vendor product list (%s) for %d line(s): %s - add them to the list'
                       % (sel or 'all lists', len(cust_miss), ', '.join(list(dict.fromkeys(cust_miss))[:20]))))
    for (h, src), its in hs_bad.items():
        keys = {k for k, _ in its}
        used = [x for x in lines if x['key'] in keys and x['hs']]
        sug = ''
        issues.append(('MISSING' if len(used) < len(keys) else 'CHECK', 'HS %s (%s) is not in the UPDATE HS CODE list - not used: %s (%d line(s): %s)%s'
                       % (h, src, ('HS %s from %s' % (used[0]['hs'], used[0]['hs_src'].split(' (')[0])) if used else 'HS left empty',
                          len(keys), ', '.join(list(dict.fromkeys(n for _, n in its))[:12]), (' - nearest in the list: %s, check and type it' % sug) if sug else '')))
    for h, its in hs_missing.items():
        issues.append(('MISSING', 'HS %s not in the UPDATE HS CODE list - HS UOM / control not checked%s (%d line(s): %s)'
                       % (h, ' - HS left blank' if ctx.get('blank_hs_not_in_list') else '', len(its),
                          ', '.join(dict.fromkeys(its)))))
    dflt = [x['item'] for x in lines if x['uom_src'] == 'Default']
    if dflt:
        issues.append(('CHECK', 'HS UOM not in the HS list / product master for %d line(s) - default %s used (%s)'
                       % (len(dflt), ctx.get('default_uom', ''), ', '.join(dict.fromkeys(dflt)))))
    if liq_eu:
        issues.append(('CHECK', 'Liquor: the invoice HS codes are EU tariff codes (%s ...) - Singapore HS + product code chosen by '
                                'the liquor rules (type, colour, bottle size, alcohol %%) from the official HS list descriptions'
                       % ', '.join(liq_eu[:3])))
    if liq_coo:
        issues.append(('CHECK', 'Origin: the invoice says %s for every line, but the appellation / producer points to another '
                                'country for %d line(s): %s - the documents\' origin is kept; change it on the Item page if needed'
                       % ('FRANCE', len(liq_coo), '; '.join(liq_coo[:12]))))
    if liq_master:
        issues.append(('CHECK', 'MASTER LIST (past permits) used other codes for %d line(s): %s - the liquor rules follow the '
                                'official descriptions' % (len(liq_master), '; '.join(liq_master[:8]))))
    # ---- consolidation
    mode = ctx.get('consolidate') or 'none'
    if rule_on('rs_consolidate'):
        mode = 'hs_coo'
    if rule_on('vf_consolidate'):
        mode = 'hs_coo_brand'
    if mode == 'none' and not ctx.get('consolidate_explicit') and any(i.get('consolidate_default') for i in sh['invs']):   # supplier default (Satair: HAWB + HS + COO)
        mode = next(i['consolidate_default'] for i in sh['invs'] if i.get('consolidate_default'))
        issues.append(('CHECK', 'Item rows: %s (supplier default) - change "Item rows" on the Item page to use another'
                       % {'hawb_hs_coo': 'same HAWB + HS code + origin'}.get(mode, mode)))
    sh['consolidate_used'] = mode
    if len({x.get('line_hawb') for x in lines}) > 1:      # several house B/Ls: items in HAWB order (files are read in parallel)
        lines = sorted(lines, key=lambda x: str(x.get('line_hawb') or ''))
    groups = collections.OrderedDict()
    for it in lines:
        if it['sep']:                      # line kept separate (chosen on tab 2)
            k = ('SEP', it['key'])
        elif mode == 'item':
            k = (it['invoice'], it['item'], it['hs'], it['coo'], it['hs_uom'], it['uom'], it['casc'], it['currency'])
        elif mode == 'hs_coo':
            k = (it['hs'], it['coo'], it['hs_uom'], it['casc'], it['currency'])
        elif mode == 'hs_coo_brand':
            k = (it['hs'], it['coo'], it['hs_uom'], it['casc'], (it.get('brand') or '').upper(), it['currency'])
        elif mode == 'hawb_hs_coo':                # one row per house B/L + HS + origin (e.g. Satair: 16 HBLs, 1 permit)
            k = (it.get('line_hawb') or '', it['hs'], it['coo'], it['hs_uom'], it['casc'], it['currency'])
        else:
            k = (it['key'],)
        if it.get('liquor') and k[0] != 'SEP' and len(k) > 1:       # liquor: never merge other strengths / bottle sizes
            k = k + (it['liquor'].get('abv'), it['liquor'].get('bottle_l'))
        groups.setdefault(k, []).append(it)
    tmpl = DEFAULT_DESC_TEMPLATE          # description only - no format
    dmax = int(ctx.get('desc_max') or 512)
    if rule_on('rs_consolidate'):
        dmax = min(dmax, 500)
    n_ = 0
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
        it['sinvs'] = sorted({x.get('sinv') for x in g if x.get('sinv')})
        # HS qty
        hq, note = None, ''
        rule = UOM_RULES.get(it['hs_uom'])
        if rule:
            basis, calc, fct = rule
            if basis == 'NET WEIGHT':
                base = it['nw']
                note = '' if base is not None else 'no net weight'
            else:
                ok = all((x['uom'] or '').upper() in COUNT_UOMS for x in g) or rule_on('rs_consolidate')   # RS: invoice qty
                if not ok and it['hs_uom'] == 'TPR' and all((x['uom'] or '').upper() in PAIR_UOMS for x in g):
                    ok = True                      # TPR = unit of ten pairs: invoice pairs / 10
                base = sum(x['qty'] or 0 for x in g) if ok else None
                note = '' if ok else 'invoice UOM %s is not a piece count' % '/'.join(sorted(uoms))
                if ok and any((x['uom'] or '').upper() in PACK_UOMS for x in g) and not rule_on('rs_consolidate'):
                    issues.append(('CHECK', 'Item %s qty is in %s (packs) - used as the %s count'
                                   % (it['item'], '/'.join(sorted(uoms)), it['hs_uom'])))
            if base is not None:
                hq = base * fct if calc == 'x' else base / fct
        elif it['hs_uom'] in MEASURE_UOMS:    # metres / litres / m2 / m3 / carat / GWh: invoice qty converted to the HS unit
            conv = MEASURE_UOMS[it['hs_uom']]
            if all(re.sub(r'[^A-Z0-9.]', '', (x['uom'] or '').upper()) in conv for x in g):
                hq = sum((x['qty'] or 0) * conv[re.sub(r'[^A-Z0-9.]', '', (x['uom'] or '').upper())] for x in g)
                if any(conv[re.sub(r'[^A-Z0-9.]', '', (x['uom'] or '').upper())] != 1 for x in g):
                    issues.append(('CHECK', 'Item %s HS Qty %s %s (%s) converted from the invoice qty %s'
                                   % (it['item'], fmt_plain(round(hq, 4)), it['hs_uom'], HS_UNITS.get(it['hs_uom'], ''),
                                      ' + '.join('%s %s' % (fmt_plain(x['qty']), x['uom']) for x in g))))
            elif it['hs_uom'] == 'LTR' and all(x.get('liquor') and x.get('ltr') for x in g):   # liquor: the litres of the bottles
                hq = sum(x['ltr'] for x in g)
            elif it['hs_uom'] == 'LTR' and not (all((x['uom'] or '').upper() in KG_UOMS for x in g) or it['nw'] is not None) \
                    and all(x.get('ltr') for x in g):          # no kg at all: litres printed on the document
                hq = sum(x['ltr'] for x in g)
                issues.append(('CHECK', 'Item %s HS Qty %s LTR from the litres on the invoice' % (it['item'], fmt_plain(round(hq, 4)))))
            elif it['hs_uom'] == 'LTR' and (all((x['uom'] or '').upper() in KG_UOMS for x in g) or it['nw'] is not None):
                # user rule: LTR from KGM - 1 kg = 1 litre (invoice qty in KG, else the net weight)
                kg_inv = all((x['uom'] or '').upper() in KG_UOMS for x in g)
                hq = sum(x['qty'] or 0 for x in g) if kg_inv else it['nw']
                issues.append(('CHECK', 'Item %s HS Qty %s LTR = %s KG (%s; rule LTR - KGM, 1 kg = 1 litre)'
                               % (it['item'], fmt_plain(round(hq, 4)), fmt_plain(round(hq, 4)),
                                  'invoice qty' if kg_inv else 'net weight')))
            else:
                note = 'HS UOM %s (%s) - invoice qty is in %s, not convertible (type the %s)' % (
                    it['hs_uom'], HS_UNITS.get(it['hs_uom'], ''), '/'.join(sorted(u or '?' for u in uoms)),
                    HS_UNITS.get(it['hs_uom'], '').lower())
        elif it['hs_uom'] == 'VAL':           # unit of value of one Singapore dollar
            if all((x.get('currency') or '').upper() == 'SGD' for x in g):
                hq = round(sum(x['amount'] or 0 for x in g), 2)
            elif ctx.get('exchange_rates', {}).get((g[0].get('currency') or '').upper()):
                hq = round(sum(x['amount'] or 0 for x in g) * ctx['exchange_rates'][(g[0].get('currency') or '').upper()], 2)
                issues.append(('CHECK', 'Item %s HS Qty VAL = %s SGD (value x exchange rate)' % (it['item'], fmt_plain(hq))))
            else:
                note = 'VAL = value in Singapore dollars - item is in %s, type the SGD value' % (g[0].get('currency') or '?')
        elif it['hs_uom']:                # other HS UOM (SET, PAR ...) = invoice qty
            if all((x['uom'] or '').upper() in COUNT_UOMS for x in g):
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
        if rule_on('rs_consolidate') and len(g) > 1 and not f.get('desc_edit'):     # RS: descriptions joined with ", "
            it['description'] = cut_words(', '.join(dict.fromkeys(re.sub(r'\s+', ' ', x.get('desc_edit') or x['desc']).strip()
                                                                for x in g if (x.get('desc_edit') or x['desc']))), dmax)
        if box.get('co_hs6') and n_ == 0:
            issues.append(('CHECK', 'Instructions: 6-digit HS for the certificate of origin%s - "HS Code on Certificate" = first 6 '
                           'digits, HS Code stays 8 digits' % (' (CO TYPE %s)' % box['co_type'] if box.get('co_type') else '')))
        n_ += 1
        inc = (box.get('desc_include') or '').upper()
        if inc and not f.get('desc_edit') and inc not in it['description'].upper():
            it['description'] = '%s, %s' % (inc, it['description']) if it['description'] else inc
        if rule_on('dow_desc') and not f.get('desc_edit'):              # DOW: description + Our Reference
            ref = ', '.join(dict.fromkeys(x.get('our_ref') for x in g if x.get('our_ref')))
            if ref and ref not in it['description']:
                it['description'] = ('%s %s' % (it['description'], ref)).strip()
        if rule_on('ddi_desc') and not f.get('desc_edit'):     # DuPont DDI: PO : material - master description
            parts = []
            for x in g:
                mat = (x.get('cust_item') or x['item'] or '').lstrip('0') or x['item']      # 00004018227 -> 4018227
                parts.append('%s : %s - %s' % (x.get('our_ref') or x.get('order') or '', mat, (x.get('cust_desc') or x['desc'] or '').strip()))
            it['description'] = cut_words(' / '.join(dict.fromkeys(parts)), dmax)
        if rule_on('dow_model'):
            it['model'] = ', '.join(dict.fromkeys(x.get('shipment') for x in g if x.get('shipment')))
            if not it['model'] and n_ == 1:
                issues.append(('CHECK', 'Customer rule "Model = shipment number": no shipment number found on the invoice - type it'))
        if rule_on('dow_brand') and not it.get('brand'):
            it['brand'], it['brand_src'] = 'DOW', 'Customer rule'
        if ((f.get('desc_rule') == 'SANDVIK' and ('sandvik_desc' not in rules_here or rule_on('sandvik_desc')))
                or rule_on('sandvik_desc')) and not f.get('desc_edit'):     # Sandvik export: TOOLS OR PARTS INV NO : ( no )
            it['description'] = 'TOOLS OR PARTS INV NO : ( %s )' % it['invoice']      # fixed text + invoice number
        if len(it['description']) > dmax:
            issues.append(('CHECK', 'Item %s description %d characters - cut to %d' % (it['item'], len(it['description']), dmax)))
            it['description'] = cut_words(it['description'], dmax) if rule_on('rs_consolidate') else it['description'][:dmax].rstrip(' ,/')
        it['item'] = ditem
        items.append(it)
    sh['items'] = items
    sh['lines_n'] = len(lines)
    issues[:] = list(dict.fromkeys(issues))          # same message once
    sh['issues_all'] = issues
    sh['invoice_total'] = sum(i['total'] if i['total'] is not None else sum(l['amount'] or 0 for l in i['lines'])
                              for i in sh['invs']) if sh['invs'] else None
    if not sh['hawb']:
        issues.append(('MISSING', 'HAWB (Carrier Reference No.) not found - fill it on tab 2'))
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
                       ('discharge_name', 'discharge_port_name'), ('mode', 'transport'), ('forwarder', 'forwarder'),
                       ('forwarder_uen', 'forwarder_uen'), ('carrier_agent_uen', 'carrier_agent_uen')):
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
        decl_text = next((t for t, c in list(lists.get('List-DeclarationType', [])) + EXTRA_DECL_TYPES if c == decl_code), '')
    if decl_text and not decl_code:
        decl_code = next((c for t, c in EXTRA_DECL_TYPES if t.upper() == decl_text.upper()), '')
    has_cont = any(s.get('containers') for s in ships)
    cargo = hd.get('cargo_pack_type') or (('Containerized' if has_cont else 'Other Non-Containerized')
                                          if cfg.get('auto_cargo', True) else '')
    tmode = hd.get('transport') or ''
    if not tmode and has_cont and cfg.get('auto_transport', True):
        tmode = 'Sea'
    put(col('party_uen'), s0.get('party_uen', ''))
    put(col('party'), s0.get('party', ''))
    if exp and s0.get('party_address'):          # exporter address from the UEN register (ACRA)
        for k, v in zip(('Exporter Address 1', 'Exporter Address 2', 'Exporter Address 3'), s0['party_address']):
            put(k, v[:35])
        mpc = re.search(r'SINGAPORE\s+(\d{6})', ' '.join(s0['party_address']))
        if mpc:
            put('Exporter Postal Code', mpc.group(1))
            put('Exporter City Name', 'SINGAPORE')
            put('Exporter Country Code', 'SG')
    un = cfg.get('uen_names') or {}
    for kn, ku in (('forwarder', 'forwarder_uen'), ('carrier_agent', 'carrier_agent_uen')):     # names by UEN (ACRA)
        u_ = str(hd.get(ku) or '').upper()
        if u_ in un and un[u_].get('name') and cfg.get('uen_names_on', True):
            hd[kn] = un[u_]['name']
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
        pu = (hd.get('outer_pack_uom') or 'PKG').upper()
        if pu == 'PKG':                     # default: use the unit on the documents when it is a pallet / carton / bag
            du = (cg.get('packages_unit') or next((p.get('package_unit') for s in ships for p in s.get('pls', []) if p.get('package_unit')), '')
                  or next((i.get('package_unit') for s in ships for i in s['invs'] if i.get('package_unit')), '')).upper()
            pu = {'PAT': 'PAT', 'PLT': 'PAT', 'PAL': 'PAT', 'PALLET': 'PAT', 'PALLETS': 'PAT', 'CTN': 'CTN', 'CTNS': 'CTN',
                  'CARTON': 'CTN', 'CARTONS': 'CTN', 'BAG': 'BAG', 'BAGS': 'BAG', 'UNT': 'UNT', 'UNIT': 'UNT', 'UNITS': 'UNT',
                  'DRM': 'DRM', 'DRUM': 'DRM', 'DRUMS': 'DRM'}.get(du, 'PKG')
        put('Total Outer Pack Uom', pu)
    nco = list(dict.fromkeys((c.get('no') if isinstance(c, dict) else str(c)) for s in ships for c in (s.get('containers') or [])
                            if (c.get('no') if isinstance(c, dict) else c)))
    if tmode == 'Sea' and nco and cfg.get('fill_outer_pack', True):      # sea job with containers -> UNT = number of containers
        put('Total Outer Pack Qty', len(nco))
        put('Total Outer Pack Uom', 'UNT')
    invs = [i['invoice_no'] for s in ships for i in s['invs'] if i.get('invoice_no')]
    rs_on = lambda rid: (s0.get('rule_on') or {}).get(rid)
    if rs_on('rs_remarks'):               # RS: J-n / shipment / GBP@rate / (INV:sales invoices of this Excel) / container & OBL
        rate = str(cfg.get('exchange_rate') or '').strip()
        rem = render_remarks(RS_REMARKS, dict(
            J=cfg.get('file_j') or j_label(cfg.get('j_start') or '1', (part[0] - 1) if part else 0),
            SHIPMENT=cg.get('shipment_no') or cfg.get('job_no', '') or cg.get('job_no', ''), CUR=s0.get('currency') or '',
            RATE=rate or ' ', INV=', '.join(sorted({x for it in items for x in (it.get('sinvs') or [])}) or invs),
            CONTAINER=', '.join(c['no'] for c in (s0.get('containers') or []) if c.get('no')) or ' ',
            OBL=hd.get('mawb') or cg.get('mawb') or ' ')).replace('@ \n', '@\n')
        if rem:
            put("Trader's Remarks", rem)
        part, rs_rem = None, True
    else:
        rs_rem = False
    rem = '' if rs_rem else render_remarks(cfg.get('remarks_template', DEFAULT_REMARKS), dict(
        INV=', '.join(dict.fromkeys(invs)), HAWB=hawb, MAWB=hd.get('mawb', ''),
        JOB=cfg.get('job_no', '') or cg.get('job_no', ''),
        PO=', '.join(dict.fromkeys(i['customer_po'] for s in ships for i in s['invs'] if i.get('customer_po'))),
        ORDER=', '.join(dict.fromkeys(i['order_no'] for s in ships for i in s['invs'] if i.get('order_no')))))
    if part:
        rem = (rem + '\nPART %d OF %d' % part).strip()
    if cg.get('permit_remarks') and not rs_rem:
        rem = (rem + '\n' + cg['permit_remarks']).strip()
    if rem:
        put("Trader's Remarks", rem)
    for h, v in (cfg.get('extra_header') or {}).items():       # any Header field typed by the user ("More header fields")
        if h.lower() in H and str(v).strip() and not other_direction_field(h, exp):
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
    inv_nos = list(dict.fromkeys(i['invoice_no'] for s in ships for i in s['invs'] if i.get('invoice_no')))
    if any((s.get('rule_on') or {}).get('dow_invno') or (s.get('rule_on') or {}).get('rs_invno') for s in ships) and inv_nos:
        _put(si, HI, 'Invoice Number', 2, ', '.join(inv_nos)[:35])        # customer rule: invoice number on the Invoice sheet
    else:
        _put(si, HI, 'Invoice Number', 2, '-')
    _put(si, HI, 'Invoice Date', 2, *_date(datetime.date.today().strftime('%d/%m/%Y'), tpl))   # always today
    cur = s0['currency']
    total = sum(s['invoice_total'] or 0 for s in ships)        # always the overall shipment value (all invoices)
    chg = round(sum((s.get('charges') or 0) - (s.get('charges_in_items') or 0) for s in ships), 2)   # not added to items
    otc = bool(chg) and not part and (cfg.get('charges_other_line', False) or
                                      any((s.get('charges_mode_used') or s.get('charges_mode')) == 'otc' for s in ships))
    if otc:
        total -= chg                       # Invoice amount = goods; the charges go in "Other Taxable Charge" below
    curs = list(dict.fromkeys(it.get('currency') for it in items if it.get('currency')))
    inv_rows = {}                          # currency -> Invoice sheet row (several currencies = one invoice row each)
    if len(curs) > 1:
        tot_c = collections.OrderedDict((c, 0.0) for c in curs)
        for s in ships:
            for i in s['invs']:
                c = i.get('currency') or ''
                if c in tot_c:
                    tot_c[c] += i.get('total') if i.get('total') is not None else sum(l.get('amount') or 0 for l in i['lines'])
        if part:                           # split Excel: this part's items only
            tot_c = collections.OrderedDict((c, sum(it['amount'] or 0 for it in items if it.get('currency') == c)) for c in curs)
        for r, c in enumerate(curs, 2):
            nos = list(dict.fromkeys(x for it in items if it.get('currency') == c for x in str(it['invoice']).split(', ') if x))
            label = (', '.join(nos) if len(', '.join(nos)) <= 35 else '%s-%s' % (c, nos[0][:30])) or c
            inv_rows[c] = (r, label)
            _put(si, HI, 'Invoice Number', r, label)
            _put(si, HI, 'Invoice Date', r, *_date(datetime.date.today().strftime('%d/%m/%Y'), tpl))
            _put(si, HI, 'Total Invoice Currency', r, c)
            _put(si, HI, 'Total Invoice Amount', r, round(tot_c[c], 2), tpl.S2)
        total, cur = tot_c[curs[0]], curs[0]
    else:
        if cur:
            _put(si, HI, 'Total Invoice Currency', 2, cur)
        if total:
            _put(si, HI, 'Total Invoice Amount', 2, round(total, 2), tpl.S2)
    if cfg.get('fill_term') and not exp:       # export permit: Term Type is not captured (typed by the user when needed)
        t = (cg.get('term_type') or '').upper() or TERM_MAP.get(s0['incoterm'], s0['incoterm'])
        if t in [x for x, _ in lists.get('List-TermType', [])]:
            for r in ([x[0] for x in inv_rows.values()] or [2]):
                _put(si, HI, 'Term Type', r, t)
    if otc:                                # charges declared on their own line (items + OTC = invoice total)
        _put(si, HI, 'Other Taxable Charge Currency', 2, cur)
        _put(si, HI, 'Other Taxable Charge Amount', 2, chg, tpl.S2)
    # charges allowed by the Term Type (as TradeNet): CIF none, CFR insurance, CNI freight, EXW / FOB / FAS both
    tt = (cg.get('term_type') or '').upper() or TERM_MAP.get(s0.get('incoterm') or '', s0.get('incoterm') or '')
    allow = TERM_CHARGES.get(tt, ('frt', 'ins'))
    fa = next((s.get('freight_amount') for s in ships if s.get('freight_amount')), None)
    if fa and 'frt' not in allow:
        fa = None
    if fa and cfg.get('fill_freight', True) and not part:
        fcur = next((s.get('freight_currency') for s in ships if s.get('freight_currency')), cur)
        fr_row = inv_rows.get(fcur, (2,))[0]
        _put(si, HI, 'Freight Charge Currency', fr_row, fcur)
        _put(si, HI, 'Freight Charge Amount', fr_row, fa, tpl.S2)
    if not exp and (permit in IMPORT_PERMITS or not permit) and cfg.get('fill_insurance', True) and 'ins' in allow:
        ins = fnum(cfg.get('insurance_pct'))
        if ins:
            for r in ([x[0] for x in inv_rows.values()] or [2]):      # every invoice row
                _put(si, HI, 'Insurance Charge %', r, ins)
                if cfg.get('insurance_cur'):
                    _put(si, HI, 'Insurance Charge Currency', r, cfg['insurance_cur'])
    for h, v in (cfg.get('extra_invoice') or {}).items():      # any Invoice sheet field typed by the user
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
        if it.get('model'):
            p('Model', it['model'])
        if it.get('line_hawb') or hawb:            # house B/L of the item (several HBLs under one ocean B/L)
            p(col('item_hawb'), it.get('line_hawb') or hawb)
        if hd.get('mawb') and cfg.get('item_mawb', True):
            p(col('item_mawb'), hd['mawb'])
        p('HS UOM', it['hs_uom'])
        p('Invoice Number', inv_rows[it['currency']][1] if it.get('currency') in inv_rows else it['invoice'])
        if cg.get('co_hs6') and it['hs']:                 # instructions "use 6 digits HS codes" -> certificate of origin
            p('HS Code on Certificate', it['hs'][:6])
            if cg.get('desc_include'):
                p('CO Certificate Description', cg['desc_include'])
        p('Item Currency', it['currency'])
        if it['casc']:
            p('CASC Product Code 1', it['casc'])
            if it['hs_qty'] is not None:
                p('CASC Product Code 1 Qty', it['hs_qty'], tpl.S4)
            if it['hs_uom']:
                p('CASC Product Code 1 UOM', it['hs_uom'])
        liq = it.get('liquor')
        if liq:                    # liquor (user's permit layout): alcohol %, bottle size as dutiable qty, bottles as outer pack
            if liq.get('abv') is not None:
                p('Alcohol Percentage', liq['abv'])
            if liq.get('bottle_l'):
                p('Dutiable Qty', liq['bottle_l'], tpl.S4)
                p('Dutiable UOM', 'LTR')
            nb = it.get('qty') if (it.get('uom') or '').upper() == 'BTL' else liq.get('bottles')
            if nb:
                p('Outer Pack Qty', int(round(nb)))
                p('Outer Pack UOM', 'BOT')
            if not it['brand']:
                p('Brand', 'UNBRANDED')
        casc_rows.append((n, it['casc'], it.get('cust_control') or '',
                          ((it['coo'] or '') + '99999') if liq and it['coo'] else ''))   # liquor: Casc Code 1 = COO + 99999
    changed[tpl.sheets['Item']] = st.tostring()
    # ---------------- Casc sheet: empty (user rule) - except customer rule DDI: Item Number / Seq 1 / Code 1 = Control Substances
    if 'Casc' in tpl.sheets:
        sk = tpl.sheet('Casc')
        sk.clear_from(2)
        if any(x[3] for x in casc_rows):           # liquor items: Code 1 = country of origin + 99999 (e.g. FR99999)
            HK = tpl.headers(sk)
            rn = 2
            for n, code, ctl, lc in casc_rows:
                if lc:
                    _put(sk, HK, 'Item Number', rn, n)
                    _put(sk, HK, 'Product Code Seq Number', rn, 1)
                    _put(sk, HK, 'Code 1', rn, lc)
                    rn += 1
        elif any((s.get('rule_on') or {}).get('ddi_casc') for s in ships):
            HK = tpl.headers(sk)
            for rn, (n, code, ctl, _lc) in enumerate(casc_rows, 2):
                ctl = '' if str(ctl).strip().upper() in ('', '#N/A', 'N/A', 'NA', '0', 'NONE', 'TBA') else ctl
                _put(sk, HK, 'Item Number', rn, n)
                _put(sk, HK, 'Product Code Seq Number', rn, 1)
                if ctl:
                    _put(sk, HK, 'Code 1', rn, re.sub(r'\s+', ' ', str(ctl)).strip()[:200])
        changed[tpl.sheets['Casc']] = sk.tostring()
    if 'Items to be consolidated' in tpl.sheets:      # sample rows of the template left as they are
        pass
    tpl.save(path, changed)
    return header_blank


def file_name_for(group, cfg):
    s0 = group[0]
    if s0.get('file_name'):
        return safe_name(s0['file_name'])
    nos = list(dict.fromkeys(i['invoice_no'] for s in group for i in s['invs'] if i.get('invoice_no')))
    inv = '_'.join(nos)
    if len(inv) > 60:                          # many invoices (e.g. 29 in one PDF): first + count, file names stay short
        inv = '%s_AND_%d_MORE' % (nos[0], len(nos) - 1)
    vals = dict(HAWB=s0['hawb'], INVOICE=inv, DIR=s0.get('direction', ''))
    pat = cfg.get('file_pattern') or '{HAWB}'
    try:
        name = pat.format(**vals)
    except (KeyError, IndexError, ValueError):
        name = vals['HAWB']
    if not s0['hawb'] and '{HAWB}' in pat:
        name = 'NO_HAWB_INV_' + inv
    return safe_name(name)[:120]


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
    jn = 0                                     # Excel file number (RS remarks J-n)
    for key, g in groups.items():
        items = [it for s in g for it in s['items']]
        rs = any((s.get('rule_on') or {}).get('rs_split') for s in g) and not cfg.get('no_split')
        if rs:                                 # RS: GENERAL / IDAMISC / PCD... in their own files, max 50 items each
            m_ = mx or 50
            grp = collections.OrderedDict([('GENERAL', [])])
            for it in items:
                c = (it.get('casc') or '').upper()
                grp.setdefault(c if c.startswith('IDAMISC') or c.startswith('PCD') else 'GENERAL', []).append(it)
            chunks, tags = [], []
            for gname, its in grp.items():
                its = sorted(its, key=lambda x: (x.get('hs') or '', x.get('coo') or ''))     # as the RS COMPO TOOL
                for i in range(0, len(its), m_):
                    chunks.append(sorted(its[i:i + m_], key=lambda x: (x.get('coo') or '', x.get('hs') or '')))
                    tags.append(gname)
        elif mx and len(items) > mx and not cfg.get('no_split') and cfg.get('split_hs'):
            # max items per Excel, rows of one HS code kept in the same Excel (an HS with more rows than the max is cut)
            by_hs = collections.OrderedDict()
            for it in items:
                by_hs.setdefault(it.get('hs') or '', []).append(it)
            chunks, cur_ = [], []
            for hs_, its in by_hs.items():
                while len(its) > mx:
                    if cur_:
                        chunks.append(cur_)
                        cur_ = []
                    chunks.append(its[:mx])
                    its = its[mx:]
                if len(cur_) + len(its) > mx:
                    chunks.append(cur_)
                    cur_ = []
                cur_ = cur_ + its
            if cur_:
                chunks.append(cur_)
            tags = [''] * len(chunks)
        else:
            chunks = [items[i:i + mx] for i in range(0, len(items), mx)] if mx and len(items) > mx and not cfg.get('no_split') else [items]
            tags = [''] * len(chunks)
        base_name = file_name_for(g, cfg)
        for ci, chunk in enumerate(chunks, 1):
            if rs:
                jl = j_label(cfg.get('j_start') or '1', jn)
                jt = 'J-%02d' % int(jl) if str(jl).isdigit() else re.sub(r'[\\/:*?"<>|\s]+', '_', str(jl))
                name = '%s_%s_%s' % (base_name, jt, re.sub(r'[^A-Z0-9_-]+', '', tags[ci - 1]))
            else:
                name = base_name if len(chunks) == 1 else '%s_%d' % (base_name, ci)
            b, k = name, 2
            while name.upper() in used:
                name = '%s_%d' % (b, k)
                k += 1
            used.add(name.upper())
            path = os.path.join(out, name + '.xlsx')
            part = (ci, len(chunks)) if len(chunks) > 1 else None
            blank = write_excel(path, g, chunk, dict(cfg, file_j=j_label(cfg.get('j_start') or '1', jn)) if rs else cfg,
                                lists, part, log)
            jn += 1
            byc = collections.OrderedDict()
            for it in chunk:
                byc[it.get('currency') or g[0]['currency']] = byc.get(it.get('currency') or g[0]['currency'], 0) + (it['amount'] or 0)
            log('%s.xlsx  -  %s, %d item row(s), %s%s' % (name, g[0]['direction'], len(chunk),
                                                         ' + '.join('%s %s' % (c, '{:,.2f}'.format(a)) for c, a in byc.items()),
                                                         ' (%d invoices on the Invoice sheet, one per currency)' % len(byc) if len(byc) > 1 else ''))
            files.append(dict(name=name + '.xlsx', path=path, ships=g, items=chunk, header_blank=blank, part=part,
                              base=base_name + '.xlsx'))
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
                control_eccn=cfg.get('control_eccn', False), control_hs=cfg.get('control_hs') or [], cont_tne_dec=cfg.get('cont_tne_dec', 2), hs_auto=cfg.get('hs_auto', True))


def run_folder(pdfs, cfg, log=print):
    """Command-line helper: parse PDFs, resolve, generate."""
    lists = template_lists(cfg['template'])
    ctx = make_ctx(cfg, lists)
    docs = [d for p in pdfs for d in parse_pdf(p, ctx['countries'], log=None, profiles=cfg.get('profiles'))]
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
                items.append({'HAWB': it.get('line_hawb') or sh.get('hawb', ''), 'row': it.get('n'), 'item': it.get('item'),
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
