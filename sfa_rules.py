# -*- coding: utf-8 -*-
"""SFA TEAM RULES - Singapore import permits for Japanese perishables (SFA project).

Used by the Japan food (GLT) tool when the job's team is SFA (New job > Team = SFA).

  HS code / CASC product code, in this order:
    1. codes saved on the Item tab ("Save to product master")          - user confirmed
    2. SFA overrides (masters/sfa/sfa_overrides.csv)                    - decisions taken on earlier permits
    3. SFA MASTER DATA (the team's own master list)
    4. past permits (RPT "HS CODE & PRODUCT CODE" reports) - most used pair that is valid in the UPDATE HS CODE list
    5. nearest description (word match)                                 - always CHECK
  Every HS | product code pair is checked against the server's UPDATE HS CODE list.

  Net weight (KG): importers in qty_as_net_weight_importers.txt (e.g. GINZA SUSHI ICHI) = the invoice (Quantity)
                   column for every line;  other importers: KG lines = quantity, PK / PC / BE ... = quantity x standard
  HS Qty by HS UOM, CASC Product Code 1 Qty by PRODUCT UOM:
        NMB = qty    KGM = net weight    TNE = net weight / 1000    TEN = qty / 10
  Description = NAME**PREFECTURE**, Casc sheet Code 1 = prefecture code (only rows with a CASC product code).

Data folder: PERMIT_SERVER_DATA/masters/sfa
    sfa_master.db        built from the files in source/ (Settings > Product masters > SFA team > Rebuild)
    sfa_overrides.csv    invoice_desc,hs_code,product_code,note   (product_code "-" = non-controlled, no CASC code)
    source/              SFA MASTER DATA + RPT reports (.xlsx / .xls)
"""
import csv, glob, json, os, re, shutil, sqlite3, unicodedata

TEAM = 'SFA'
DIR_NAME = 'sfa'
DB_NAME = 'sfa_master.db'
OVR_NAME = 'sfa_overrides.csv'
RULES_NAME = 'SFA_RULES.md'
SRC_NAME = 'source'
OVR_HEADER = ['invoice_desc', 'hs_code', 'product_code', 'note']
QTYNW_NAME = 'qty_as_net_weight_importers.txt'    # importers whose invoice (Quantity) = net weight, one per line
QTYNW_DEFAULT = ['GINZA SUSHI ICHI PTE. LTD.']
STOP = {'FRESH', 'FROZEN', 'CHILLED', 'REITOU', 'JPN', 'WHOLE', 'WILD', 'THE', 'OF', 'AND', 'FISH', 'PK', 'KG'}
JP_PREF = ('AICHI JPACH|AKITA JPAKT|AOMORI JPAMR|CHIBA JPCHB|EHIME JPEHM|FUKUOKA JPFKK|FUKUSHIMA JPFKS|FUKUI JPFUK|'
           'GIFU JPGIF|GUNMA JPGUM|HOKKAIDO JPHKD|HIROSHIMA JPHRS|HYOGO JPHYG|IBARAKI JPIBR|ISHIKAWA JPISK|IWATE JPIWT|'
           'KAGOSHIMA JPKGS|KAGAWA JPKGW|KUMAMOTO JPKMM|KANAGAWA JPKNG|KOCHI JPKOC|KYOTO JPKYT|MIE JPMIE|MIYAGI JPMYG|'
           'MIYAZAKI JPMYZ|NARA JPNAR|NAGANO JPNGN|NAGASAKI JPNGS|NIIGATA JPNGT|OITA JPOIT|OKINAWA JPOKN|OKAYAMA JPOKY|'
           'OSAKA JPOSK|SAGA JPSAG|SHIGA JPSHG|SHIMANE JPSMN|SAITAMA JPSTM|SHIZUOKA JPSZO|TOCHIGI JPTCG|TOTTORI JPTGR|'
           'TOKUSHIMA JPTKS|TOKYO JPTOK|TOYAMA JPTYM|WAKAYAMA JPWKY|YAMAGUCHI JPYMG|YAMANASHI JPYMN|YAMAGATA JPYMT')


# ---------------------------------------------------------------- helpers
def s(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return unicodedata.normalize('NFKC', str(v)).strip()


def base_desc(d):
    """Description without the **PREFECTURE** tag, upper case, single spaces."""
    d = re.sub(r'\*\*.*?\*\*', '', s(d))
    return re.sub(r'\s+', ' ', d.replace('*', '')).strip().upper()


def pref_of(d):
    m = re.search(r'\*\*(.*?)\*\*', s(d))
    return m.group(1).strip().upper() if m else ''


def key(d):
    """Match key: letters only (spaces / brackets / full-width characters do not matter)."""
    return re.sub(r'[^A-Z]', '', base_desc(d))


def name_key(d):
    """Key of the Japanese name only: 'KINMEDAI (ALFONSINO)' -> 'KINMEDAI' ('' when there is no bracket)."""
    b = base_desc(d)
    return re.sub(r'[^A-Z]', '', b.split('(')[0]) if '(' in b else ''


def hs8(v):
    v = re.sub(r'\D', '', s(v))
    return v.zfill(8)[:8] if v else ''


def conv(uom, qty, nw):
    """Quantity in a UOM: NMB = qty, KGM = net weight, TNE = net weight / 1000, TEN = qty / 10."""
    uom = s(uom).upper()
    if uom == 'NMB':
        return qty
    if uom == 'KGM':
        return nw
    if uom == 'TNE':
        return None if nw is None else nw / 1000.0
    if uom == 'TEN':
        return None if qty is None else qty / 10.0
    return None


def net_kg(qty, unit, std, std_unit, qty_as_nw=False):
    """Net weight in KG of one invoice line. qty_as_nw (importer rule): the (Quantity) column.
    Otherwise KG lines = quantity, PK / PC / BE ... = quantity x standard (250 G/PK x 1 = 0.25 kg). Returns (kg, note)."""
    if qty is None:
        return None, 'no quantity'
    if qty_as_nw:
        return qty, ''
    unit = s(unit).upper()
    su = s(std_unit).upper().replace(' ', '')
    if qty is None:
        return None, 'no quantity'
    if unit in ('KG', 'KGS', 'KGM'):
        return qty, ''
    if unit in ('G', 'GR', 'GRM'):
        return qty / 1000.0, ''
    if std:
        m = re.match(r'^(KG|KGS|G|GR|GRM)\b', su) or re.match(r'^(KG|KGS|G|GR|GRM)', su)
        if m:
            f = 1.0 if m.group(1).startswith('K') else 0.001
            return qty * std * f, ''
    return qty, 'no standard weight for %s - net weight = quantity, CHECK' % (unit or 'line')


def _rows(path):
    """All rows of every sheet of an .xlsx / .xls file (xls needs xlrd - in the exe)."""
    if path.lower().endswith('.xls'):
        import xlrd
        wb = xlrd.open_workbook(path)
        for sh in wb.sheets():
            for i in range(sh.nrows):
                yield sh.name, sh.row_values(i)
        return
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            try:
                ws.reset_dimensions()
            except Exception:
                pass
            for r in ws.iter_rows(values_only=True):
                yield ws.title, list(r)
    finally:
        wb.close()


def file_kind(path):
    """'MASTER' (SFA MASTER DATA: sheet DATA, description | hs | product code ... prefecture | code),
    'RPT' (permit report with 'Description Goods' + 'HS Code'), or ''."""
    try:
        for n, (sheet, r) in enumerate(_rows(path)):
            up = [s(v).upper() for v in r]
            if sheet.strip().upper() == 'DATA' and n == 0 and any('HS' in c for c in up) and any('PREF' in c for c in up):
                return 'MASTER'
            if 'DESCRIPTION GOODS' in up and 'HS CODE' in up:
                return 'RPT'
            if n > 60:
                break
    except Exception:
        return ''
    return ''


# ---------------------------------------------------------------- HS list (the server's UPDATE HS CODE list)
class HSList:
    def __init__(self, hs_json):
        self.uom, self.pc = {}, {}
        try:
            with open(hs_json, encoding='utf-8') as f:
                d = json.load(f)
            for k, r in (d.get('hs') or {}).items():
                self.uom[k] = s(r[1]).upper()
            for k, lst in (d.get('pcodes') or {}).items():
                self.pc[k] = {s(p[0]).upper(): (s(p[1]), s(p[2]).upper()) for p in lst}
        except Exception:
            pass

    def ok(self):
        return len(self.uom) > 1000

    def pair(self, hs, code):
        """(product description, product UOM) or None."""
        return (self.pc.get(hs) or {}).get(s(code).upper())


# ---------------------------------------------------------------- database
def build_db(src_dir, out_db, hs_json):
    """Build sfa_master.db from SFA MASTER DATA + RPT reports in src_dir. Returns counts."""
    hsl = HSList(hs_json)
    tmp = out_db + '.tmp'
    if os.path.exists(tmp):
        os.remove(tmp)
    db = sqlite3.connect(tmp)
    db.executescript('''
    CREATE TABLE sfa_master(description TEXT, base_desc TEXT, k TEXT, hs_code TEXT, product_code TEXT, source TEXT);
    CREATE TABLE prefectures(prefecture TEXT PRIMARY KEY, code TEXT);
    CREATE TABLE permit_history(source TEXT, permit_no TEXT, importer TEXT, description TEXT, base_desc TEXT, k TEXT,
                                prefecture TEXT, hs_code TEXT, product_code TEXT, product_code2 TEXT);
    CREATE TABLE info(name TEXT PRIMARY KEY, value TEXT);''')
    files, bad = [], []
    for p in sorted(glob.glob(os.path.join(src_dir, '*'))):
        if not p.lower().endswith(('.xlsx', '.xls', '.xlsm')) or os.path.basename(p).startswith('~$'):
            continue
        kind = file_kind(p)
        if not kind:
            bad.append(os.path.basename(p))
            continue
        files.append((os.path.basename(p), kind))
        if kind == 'MASTER':
            for n, (sheet, r) in enumerate(_rows(p)):
                if sheet.strip().upper() != 'DATA' or n == 0:
                    continue
                r = list(r) + [None] * 7
                if s(r[0]) and s(r[1]):
                    db.execute('INSERT INTO sfa_master VALUES(?,?,?,?,?,?)',
                               (s(r[0]), base_desc(r[0]), key(r[0]), hs8(r[1]), s(r[2]).upper(), os.path.basename(p)))
                if s(r[5]) and s(r[6]) and s(r[5]).upper() != 'PREFECTURE':
                    db.execute('INSERT OR REPLACE INTO prefectures VALUES(?,?)', (s(r[5]).upper(), s(r[6]).upper()))
        else:
            hdr = None
            for sheet, r in _rows(p):
                up = [s(v) for v in r]
                if 'Description Goods' in up and 'HS Code' in up:
                    hdr = {v: i for i, v in enumerate(up) if v}
                    continue
                if not hdr:
                    continue
                g = lambda k_: s(r[hdr[k_]]) if k_ in hdr and hdr[k_] < len(r) else ''
                d, h = g('Description Goods'), g('HS Code')
                if d and h and h[:1].isdigit():
                    db.execute('INSERT INTO permit_history VALUES(?,?,?,?,?,?,?,?,?,?)',
                               (os.path.basename(p), g('Permit Number'), g('Importer Name'), d, base_desc(d), key(d),
                                pref_of(d), hs8(h), g('CA/SC Product Code 1').upper(), g('CA/SC Product Code 2').upper()))
    for pair in JP_PREF.split('|'):
        n_, c = pair.rsplit(' ', 1)
        db.execute('INSERT OR IGNORE INTO prefectures VALUES(?,?)', (n_, c))
    # one best answer per description key: valid pair first, then master before history, then most used
    cand = {}
    for src, pri, sql in (('MASTER', 1, 'SELECT k, base_desc, hs_code, product_code, COUNT(*) FROM sfa_master '
                                        "WHERE product_code<>'' GROUP BY 1,3,4"),
                          ('HISTORY', 2, 'SELECT k, base_desc, hs_code, product_code, COUNT(*) FROM permit_history '
                                         "WHERE product_code<>'' GROUP BY 1,3,4")):
        for k_, bd, h, c, n_ in db.execute(sql):
            valid = 1 if (not hsl.ok() or hsl.pair(h, c)) else 0
            cand.setdefault(k_, []).append((-valid, pri, -n_, bd, h, c, src, n_, valid))
    db.execute('CREATE TABLE desc_lookup(k TEXT PRIMARY KEY, base_desc TEXT, hs_code TEXT, product_code TEXT, source TEXT, '
               'times_used INTEGER, valid INTEGER)')
    for k_, lst in cand.items():
        lst.sort()
        _, _, _, bd, h, c, src, n_, valid = lst[0]
        db.execute('INSERT INTO desc_lookup VALUES(?,?,?,?,?,?,?)', (k_, bd, h, c, src, n_, valid))
    cnt = {t: db.execute('SELECT COUNT(*) FROM %s' % t).fetchone()[0]
           for t in ('sfa_master', 'permit_history', 'prefectures', 'desc_lookup')}
    db.execute('INSERT INTO info VALUES(?,?)', ('files', json.dumps(files)))
    db.execute('INSERT INTO info VALUES(?,?)', ('counts', json.dumps(cnt)))
    db.commit()
    db.close()
    os.replace(tmp, out_db)
    return dict(cnt, files=files, not_read=bad)


# ---------------------------------------------------------------- rules object
class Rules:
    def __init__(self, folder, hs_json):
        self.folder = folder
        self.hs = HSList(hs_json)
        self.look, self.all, self.pref, self.ov, self.counts, self.names = {}, [], {}, {}, {}, {}
        dbp = os.path.join(folder, DB_NAME)
        if os.path.exists(dbp):
            db = sqlite3.connect(dbp)
            for k_, bd, h, c, src, n_, v in db.execute('SELECT * FROM desc_lookup'):
                self.look[k_] = (h, c, src, n_, v)
                if v:
                    self.all.append((bd, h, c, n_))
                    nk = name_key(bd)
                    if len(nk) >= 3:
                        self.names.setdefault(nk, {}).setdefault((h, c), [0, bd])[0] += n_
            self.pref = dict(db.execute('SELECT prefecture, code FROM prefectures'))
            try:
                self.counts = json.loads(db.execute("SELECT value FROM info WHERE name='counts'").fetchone()[0])
            except Exception:
                pass
            db.close()
        for pair in JP_PREF.split('|'):
            n_, c = pair.rsplit(' ', 1)
            self.pref.setdefault(n_, c)
        self.qty_nw = read_importers(os.path.join(folder, QTYNW_NAME))
        for r in read_overrides(os.path.join(folder, OVR_NAME)):
            self.ov[key(r['invoice_desc'])] = (hs8(r['hs_code']), s(r['product_code']).upper(), s(r.get('note')))

    def qty_is_nw(self, *names):
        """True when one of the importer names (job importer / invoice consignee) is in the (Quantity) = net weight list."""
        keys = [re.sub(r'[^A-Z]', '', s(n).upper()) for n in self.qty_nw]
        keys = [re.sub(r'(PTE|LTD|PRIVATE|LIMITED)+$', '', k) or k for k in keys]
        for n in names:
            nk = re.sub(r'[^A-Z]', '', s(n).upper())
            if nk and any(k and k in nk for k in keys):
                return True
        return False

    def ready(self):
        return bool(self.look)

    def pref_code(self, pref):
        p = re.sub(r'[^A-Z]', '', s(pref).upper())
        return self.pref.get(p, '')

    def lookup(self, name, alt=()):
        """(hs, product code, source, check) for an invoice description. product code '-' = non-controlled."""
        for nm in (name,) + tuple(alt):
            k_ = key(nm)
            if k_ in self.ov:
                h, c, note = self.ov[k_]
                return h, c, 'SFA override' + (' - ' + note if note else ''), False
        for nm in (name,) + tuple(alt):
            r = self.look.get(key(nm))
            if r:
                h, c, src, n_, v = r
                return h, c, ('SFA master data' if src == 'MASTER' else 'Past permits (%d x)' % n_), not v
        nk = name_key(name)                     # same Japanese name, one code in every past permit
        if nk in self.names and len(self.names[nk]) == 1:
            (h, c), (n_, bd) = next(iter(self.names[nk].items()))
            return h, c, "Same name as '%s' (%d x)" % (bd, n_), False
        words = {w for w in re.findall(r'[A-Z]+', base_desc(name)) if w not in STOP and len(w) > 2}
        best = None
        for d, h, c, n_ in self.all:
            dw = set(re.findall(r'[A-Z]+', d))
            hit = len(words & dw)
            if hit:
                sc = (hit / max(len(words), 1), -len(dw - words), n_)
                if not best or sc > best[0]:
                    best = (sc, h, c, d)
        if best and best[0][0] >= 0.5:
            return best[1], best[2], "Nearest description '%s'" % best[3], True
        return '', '', '', False

    def hs_uom(self, hs):
        return self.hs.uom.get(hs, '')

    def product_uom(self, hs, code):
        p = self.hs.pair(hs, code)
        return p[1] if p else ''

    def pair_ok(self, hs, code):
        return bool(self.hs.pair(hs, code)) if self.hs.ok() else True


def read_importers(path):
    try:
        with open(path, encoding='utf-8-sig') as f:
            return [s(l) for l in f if s(l) and not s(l).startswith('#')]
    except Exception:
        return []


def write_importers(path, names):
    names = list(dict.fromkeys(s(n).upper() for n in names if s(n)))
    with open(path, 'w', encoding='utf-8') as f:
        f.write('# Importers whose GLT invoice (Quantity) column is the net weight (KG) - one per line\n')
        f.write(''.join(n + '\n' for n in names))
    return names


def read_overrides(path):
    out = []
    try:
        with open(path, encoding='utf-8-sig', newline='') as f:
            for r in csv.DictReader(f):
                if s(r.get('invoice_desc')) and s(r.get('hs_code')):
                    out.append(r)
    except Exception:
        pass
    return out


def check_overrides_file(path):
    """Raises ValueError when an uploaded overrides CSV does not have the right columns. Returns row count."""
    with open(path, encoding='utf-8-sig', newline='') as f:
        rd = csv.DictReader(f)
        cols = [s(c).lower() for c in rd.fieldnames or []]
        if not {'invoice_desc', 'hs_code', 'product_code'} <= set(cols):
            raise ValueError('The CSV needs the columns invoice_desc, hs_code, product_code (note optional)')
        return sum(1 for r in rd if s(r.get('invoice_desc')))


# ---------------------------------------------------------------- folder (first run: copy what is packed with the server)
def setup_folder(masters_dir, packed_dir):
    """PERMIT_SERVER_DATA/masters/sfa - copied from the server folder masters/sfa the first time (never overwritten)."""
    dst = os.path.join(masters_dir, DIR_NAME)
    os.makedirs(os.path.join(dst, SRC_NAME), exist_ok=True)
    src = os.path.join(packed_dir, DIR_NAME)
    if os.path.isdir(src):
        for fn in (DB_NAME, OVR_NAME, RULES_NAME, QTYNW_NAME):
            if os.path.exists(os.path.join(src, fn)) and not os.path.exists(os.path.join(dst, fn)):
                shutil.copy(os.path.join(src, fn), os.path.join(dst, fn))
        for fn in os.listdir(os.path.join(src, SRC_NAME)) if os.path.isdir(os.path.join(src, SRC_NAME)) else []:
            if not os.path.exists(os.path.join(dst, SRC_NAME, fn)):
                shutil.copy(os.path.join(src, SRC_NAME, fn), os.path.join(dst, SRC_NAME, fn))
    if not os.path.exists(os.path.join(dst, QTYNW_NAME)):
        write_importers(os.path.join(dst, QTYNW_NAME), QTYNW_DEFAULT)
    if not os.path.exists(os.path.join(dst, OVR_NAME)):
        with open(os.path.join(dst, OVR_NAME), 'w', encoding='utf-8', newline='') as f:
            csv.writer(f).writerow(OVR_HEADER)
    return dst


_CACHE = {}


def get(folder, hs_json):
    """Rules of the SFA folder (cached; reloaded when the database / overrides / HS list change)."""
    sig = tuple(os.path.getmtime(p) if os.path.exists(p) else 0
                for p in (os.path.join(folder, DB_NAME), os.path.join(folder, OVR_NAME), os.path.join(folder, QTYNW_NAME),
                          hs_json or ''))
    if _CACHE.get('key') != (folder, hs_json, sig):
        _CACHE.update(key=(folder, hs_json, sig), rules=Rules(folder, hs_json))
    return _CACHE['rules']


def status(folder, hs_json):
    r = get(folder, hs_json)
    dbp = os.path.join(folder, DB_NAME)
    files = []
    try:
        db = sqlite3.connect(dbp)
        files = json.loads(db.execute("SELECT value FROM info WHERE name='files'").fetchone()[0])
        db.close()
    except Exception:
        pass
    import time
    src = sorted(os.listdir(os.path.join(folder, SRC_NAME))) if os.path.isdir(os.path.join(folder, SRC_NAME)) else []
    return dict(ready=r.ready(), counts=r.counts, overrides=len(r.ov), lookup=len(r.look), files=files, source=src,
                qty_nw_importers=r.qty_nw,
                updated=time.strftime('%Y-%m-%d %H:%M', time.localtime(os.path.getmtime(dbp))) if os.path.exists(dbp) else '')


# ---------------------------------------------------------------- apply to one Japan food item
def fmt_qty(q, kg=None):
    """Same format as the Japan food tool: 4 decimals, at least 0.0001 when there is a weight."""
    if q is None:
        return None
    q = float('%.4f' % q)
    if q == 0 and (kg or 0) > 0:
        q = 0.0001
    return q


def apply_codes(it, rules, keep_codes=False, keep_nw=False, qty_as_nw=False):
    """Step 1 (before the Item tab edits): net weight + HS / CASC code by the SFA rules. Starts it['sfa_notes']."""
    notes = it.setdefault('sfa_notes', [])
    it['sfa'] = True
    if not keep_nw:
        kg, n = net_kg(it.get('inv_qty'), it.get('inv_unit'), it.get('std'), it.get('std_unit'), qty_as_nw)
        it['nw_rule'] = 'Invoice quantity (importer rule)' if qty_as_nw else 'Quantity x standard (SFA rule)'
        it['kg'] = kg
        if n:
            notes.append(n)
    if not keep_codes:
        h, c, src, chk = rules.lookup(it.get('name') or it.get('desc') or '')
        it['hs'], it['pc1'], it['pc2'] = h, ('' if c == '-' else c), ''
        it['sfa_src'] = src or 'not found'
        it['status'] = ('SFA: ' + src) if src else 'SFA: NOT FOUND'
        if not h:
            notes.append('HS / product code not found in the SFA data - type it on the Item tab')
        elif chk:
            notes.append('code from %s - CHECK' % src)
    else:
        it['sfa_src'] = 'Saved code (Item tab)'
    pc = rules.pref_code(it.get('pref'))
    if pc:
        it['pref_code'] = pc
    return it


def apply_qty(it, rules, uom_typed=''):
    """Step 2 (after the Item tab edits and permit_core.recalc): HS UOM / HS Qty / CASC qty + UOM by the SFA rules."""
    notes = it.setdefault('sfa_notes', [])
    hs = it.get('hs') or ''
    it['uom'] = (uom_typed or rules.hs_uom(hs) or it.get('uom') or '').upper()
    kg, qty = it.get('kg'), it.get('inv_qty')
    q = conv(it['uom'], qty, kg)
    if q is None and it['uom']:
        if it['uom'] == 'GRM' and kg is not None:
            q = kg * 1000
        else:
            notes.append('HS UOM %s - no SFA conversion, invoice quantity used' % it['uom'])
            q = qty
    it['hs_qty'] = fmt_qty(q, kg)
    if it.get('pc1'):
        pu = rules.product_uom(hs, it['pc1'])
        if not pu:
            notes.append('product code %s is not under HS %s in the UPDATE HS CODE list' % (it['pc1'], hs))
            pu = it['uom']
        pq = conv(pu, qty, kg)
        it['pc1_uom'], it['pc1_qty'] = pu, fmt_qty(pq if pq is not None else it['hs_qty'], kg)
    else:
        it['pc1_uom'] = it['pc1_qty'] = None
    it['group'] = (it.get('pc1') or 'Z')[:1].upper()
    # permit_core warnings that the SFA rules settle
    w = [x for x in it.get('warnings') or [] if not (x == 'PRODUCT CODE NOT UNDER HS')]
    if hs and it.get('pc1') and not rules.pair_ok(hs, it['pc1']):
        w.append('PRODUCT CODE NOT UNDER HS')
    it['warnings'] = w
    return it
