# -*- coding: utf-8 -*-
"""Runs the permit engines (Danfoss / Festo / Agilent / any supplier) for the server.

Same steps as the desktop tools: read PDFs -> shipments (one per HBL / HAWB) -> rules (resolve) ->
edits on items / shipments -> generate dec-upload Excel + CHECK_REPORT + FrtImport XML + RAW_DATA."""
import os, re, sys, json, copy, math, shutil, pickle, importlib, collections, threading

import xml_export as X


def base_dir():
    return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))


TOOLS = collections.OrderedDict([
    ('danfoss', dict(name='Danfoss', engine='danfoss_engine', master='DANFOSS_PRODUCT_MASTER.xlsx', color='#E2000F',
                     charges=True, hint='Danfoss commercial invoice + packing lists + DHL B/L + customs instructions')),
    ('festo', dict(name='Festo', engine='festo_engine', master='FESTO_PRODUCT_MASTER.xlsx', color='#0091DC',
                   charges=True, hint='Festo SPOOL / ICAPO invoices + ZGPL packing list + B/L + customs instructions')),
    ('agilent', dict(name='Agilent', engine='agilent_engine', master='AGILENT_PRODUCT_MASTER.xlsx', color='#0085D5',
                     charges=False, hint='Agilent invoices + packing lists + air waybill / DGF / KWE papers')),
    ('invoice', dict(name='Any supplier', engine='permit_engine', master='INVOICE_PRODUCT_MASTER.xlsx', color='#2E7D32',
                     charges=False, hint='Any supplier + known layouts: Nurmi (scanned), UD Trucks, Kalmar, Getinge, Cummins, L Beauty, BMW, Perkins, Legrand ...')),
    ('msd', dict(name='MSD (export)', engine='msd_server_engine', master='MSD_PRODUCT_MASTER.xlsx', color='#00857C',
                 charges=False, hint='MSD / Merck Sharp & Dohme EXPORT: invoices + packing lists + B/L + DHL instruction e-mail '
                                     '(paste it in Cargo details or upload the .txt / .eml)')),
])
TOOLS['jfood'] = dict(name='Japan food (GLT)', engine='jfood_server_engine', master='JAPAN_FOOD_SFA_MASTER.xlsx', color='#C62828',
                      charges=False, hint='GLT CO.,LTD invoice Excel + air waybill PDF (+ SFA HS CODE & PRODUCT CODE list, '
                                          'previous permit Excel for the header) - Japanese food imports. '
                                          'Team SFA = SFA team rules (SFA master data, overrides, net weight by importer rule, NMB/KGM/TNE/TEN)')
BUILTIN_TOOLS = tuple(TOOLS)
TEMPLATE = 'dec-upload-template-with-description.xlsx'
HEADER_FIELDS = [  # (key, label) - same boxes as tab 3 of the desktop tools; blank = taken from the documents
    ('mawb', 'MAWB / OBL'), ('transport_id', 'Transport identifier (vessel / airline)'),
    ('flight', 'Flight no. / voyage'), ('date', 'Arrival / departure date (dd/mm/yyyy)'),
    ('loading_port_code', 'Loading port code'), ('loading_port_name', 'Loading port name'),
    ('discharge_port_code', 'Discharge port code (export only)'), ('discharge_port_name', 'Discharge port name (export only)'),
    ('release_code', 'Release location code'), ('release_name', 'Release location name'),
    ('receipt_code', 'Receipt location code'), ('receipt_name', 'Receipt location name'),
    ('forwarder', 'Freight forwarder'), ('forwarder_uen', 'Freight forwarder UEN'),
    ('carrier_agent', 'Carrier agent name'), ('carrier_agent_uen', 'Carrier agent UEN'),
    ('outer_pack_uom', 'Total outer pack UOM'),
]
PERMITS = ['', 'IG', 'ME', 'IT', 'IM', 'OD', 'OT', 'ID']
LOCATIONS = collections.OrderedDict([('PPZ', 'PASIR PANJANG FTZ'), ('ME', 'MAJOR EXPORTER SCHEME'), ('KZ', 'KEPPEL FTZ'),
                                     ('JZ', 'JURONG FTZ'), ('TZ', 'TUAS TERMINAL GATEWAY'), ('CZ', 'CHANGI FTZ'),
                                     ('O', 'OTHERS')])
CONSOLIDATE = collections.OrderedDict([('none', 'No - one row per invoice line'),
                                       ('item', 'Same Item No + HS + COO (same invoice)'),
                                       ('hs_coo', 'Same HS code + Origin (COO)'),
                                       ('hawb_hs_coo', 'Same HAWB + HS code + Origin (COO)')])
# HS lists by description (one Excel per customer): used when an item has no HS on the invoice / in the product master
HS_DESC_DEFAULT = [dict(name='PANASONIC FACTORY', match=r'PANASONIC', file='HS_CODE_PANASONIC_FACTORY.xlsx', team='KWE')]

DEFAULT_SETTINGS = {
    'hs_desc_lists': HS_DESC_DEFAULT,
    'header': {'outer_pack_uom': 'PKG'},
    'insurance_pct': '1', 'insurance_cur': 'SGD',
    'remarks_template': '',
    'hs_keywords': [['CAPACITOR', '85322400'], ['INDUCTOR', '85049090']],
    'teams': ['DGF OCEAN', 'KWE', 'SFA', 'CLASQUIN'],
    'cust_teams': {'Qnity_SG_GMID_Lists_DuPont_DSP_EKC_Master_List.xlsx': 'DGF OCEAN', 'DDIS_Material_Master.xlsx': 'DGF OCEAN', 'GIVAUDAN.xlsx': 'DGF OCEAN'},
    'options': {'fill_term': True, 'fill_insurance': True, 'fill_container': True, 'fill_outer_pack': True,
                'auto_cargo': True, 'auto_transport': True, 'fill_forwarder': True, 'item_mawb': True,
                'blank_hs_not_in_list': False, 'control_by_eccn': False, 'hs_auto': True, 'ocr': True,
                'raw': True, 'uen_names': True, 'carrier_by_uen': True, 'xml': True, 'xml_charges': False, 'type_in_desc': True, 'fill_freight': True,
                'ignore_line_nw': False},
    'xml': {'snd': 'ITG_INDIA', 'rcp': 'YCH', 'msg_import': 'IPTDEC', 'msg_export': ''},
    'msg_types': ['IPTDEC'],                 # XML message types offered on the job (Header tab)
    'msg_rules': [],                         # [[team / importer / exporter name contains, message type]] - job choice wins
    'carrier_book': {'197201770G': 'SINGAPORE AIRPORT TERMINAL SERVICES'},   # carrier agent UEN -> name
    'rules': {'desc_max': '512', 'default_uom': '', 'brand': '', 'file_pattern': '{HAWB}', 'consolidate': 'none',
              'max_items': '0', 'cont_dec': '2'},
    'forwarder_map': {'DHL GLOBAL FORWARDING': 'DHL GLOBAL FORWARDING (SINGAPORE) PTE LTD|197000723M'},
    'control': {'casc': 'MISC', 'hs': ''},
    'locations': dict(LOCATIONS),
    'uen_book': {},
    'hs_map': {},
    'uom_rules': None,
}
SNAP = ('overrides', 'gw_override', 'pkg_override', 'file_name', 'include', 'containers_edit', 'charges_mode',
        'hawb', 'hawb_src', 'party', 'party_src', 'party_uen', 'party_uen_src', 'dest', 'dest_src')
DOC_DROP = ('img', 'words', 'chars', 'pages', 'image')

def detect_tool(paths):
    """Which reader fits these documents (Auto on the New job page)."""
    import pdfplumber
    txt = ''
    for p in paths:
        if not p.lower().endswith('.pdf'):
            continue
        try:
            with pdfplumber.open(p) as pdf:
                for pg in pdf.pages[:2]:
                    txt += (pg.extract_text() or '')[:3000] + '\n'
        except Exception:
            continue
    if len(re.sub(r'\s', '', txt)) < 40:              # scanned PDFs: OCR the first page of each (quick, low resolution)
        try:
            import pypdfium2 as pdfium, ocr_runner
            if ocr_runner.STATE.get('cmd') or ocr_runner.setup():
                for p in paths[:4]:
                    if p.lower().endswith('.pdf'):
                        doc = pdfium.PdfDocument(p)
                        for i in range(min(len(doc), 3)):
                            txt += ocr_runner.image_to_string(doc[i].render(scale=2.5).to_pil().convert('L'), 6) + '\n'
                        doc.close()
        except Exception:
            pass
    for p in paths:                                    # Excel documents: the first rows of each sheet
        if p.lower().endswith(('.xlsx', '.xlsm')):
            try:
                from openpyxl import load_workbook
                wb = load_workbook(p, read_only=True, data_only=True)
                for ws in wb.worksheets[:4]:
                    for r in ws.iter_rows(max_row=25, values_only=True):
                        txt += ' '.join(str(v) for v in r if v not in (None, '')) + '\n'
                wb.close()
            except Exception:
                pass
    txt += ' '.join(os.path.basename(p) for p in paths)
    u = txt.upper()
    if re.search(r'GLT\s*CO\.?,?\s*LTD', u) or (re.search(r'PRODUCTION\s+CENTER', u) and 'DESCRIPTION' in u and 'AMOUNT' in u):
        return 'jfood'
    for key, t in TOOLS.items():                      # plugin tools first (their own detect function)
        fn = t.get('detect')
        if fn:
            try:
                if fn(txt):
                    return key
            except Exception:
                pass
    if 'MERCK SHARP' in u or 'MSD INTERNATIONAL' in u or re.search(r'\bMSD\b.*(EXPORT|INVOICE)', u):
        return 'msd'
    if 'DANFOSS' in u:
        return 'danfoss'
    if 'FESTO' in u:
        return 'festo'
    if 'AGILENT' in u:
        return 'agilent'
    return 'invoice'


_ENG = {}
_ENG_LOCK = threading.Lock()


def engine(tool):
    with _ENG_LOCK:
        if tool not in _ENG:
            E = TOOLS[tool].get('module') or importlib.import_module(TOOLS[tool]['engine'])
            try:
                E.set_tesseract('')
            except Exception:
                pass
            _ENG[tool] = E
        return _ENG[tool]


class Paths:
    """data folder next to the exe / server.py: database, jobs, masters, HS list, settings."""
    def __init__(self, data):
        self.data = data
        self.jobs = os.path.join(data, 'jobs')
        self.masters = os.path.join(data, 'masters')
        for p in (self.data, self.jobs, self.masters):
            os.makedirs(p, exist_ok=True)
        for t in TOOLS.values():                     # first run: copy the packed product masters
            dst = os.path.join(self.masters, t['master'])
            src = os.path.join(base_dir(), 'masters', t['master'])
            if not os.path.exists(dst) and os.path.exists(src):
                shutil.copy(src, dst)
        self.hsdesc = os.path.join(self.masters, 'hs_desc')        # HS lists by description (per customer)
        os.makedirs(self.hsdesc, exist_ok=True)
        src_d = os.path.join(base_dir(), 'masters', 'hs_desc')
        for fn in os.listdir(src_d) if os.path.isdir(src_d) else []:
            if not os.path.exists(os.path.join(self.hsdesc, fn)):
                shutil.copy(os.path.join(src_d, fn), os.path.join(self.hsdesc, fn))
        self.cascdir = os.path.join(self.masters, 'casc_by_hs')      # CASC product code by HS (RS Components master)
        os.makedirs(self.cascdir, exist_ok=True)
        src_c = os.path.join(base_dir(), 'masters', 'casc_by_hs')
        for fn in os.listdir(src_c) if os.path.isdir(src_c) else []:
            if not os.path.exists(os.path.join(self.cascdir, fn)):
                shutil.copy(os.path.join(src_c, fn), os.path.join(self.cascdir, fn))
        self.custdir = os.path.join(self.masters, 'customer_products')   # customer product lists (item -> HS + CA/SC code)
        os.makedirs(self.custdir, exist_ok=True)
        src_p = os.path.join(base_dir(), 'masters', 'customer_products')
        for fn in os.listdir(src_p) if os.path.isdir(src_p) else []:
            if not os.path.exists(os.path.join(self.custdir, fn)) and not os.path.exists(os.path.join(self.custdir, fn + '.removed')):
                shutil.copy(os.path.join(src_p, fn), os.path.join(self.custdir, fn))
        self.hs = os.path.join(data, 'hs_data.json')
        if not os.path.exists(self.hs):
            shutil.copy(os.path.join(base_dir(), 'hs_data.json'), self.hs)
        self.profiles = os.path.join(data, 'supplier_profiles.json')
        try:                                                         # SFA team: database + overrides (masters/sfa)
            import sfa_rules
            self.sfa = sfa_rules.setup_folder(self.masters, os.path.join(base_dir(), 'masters'))
        except Exception:
            self.sfa = os.path.join(self.masters, 'sfa')

    def master(self, tool):
        return os.path.join(self.masters, TOOLS[tool]['master'])

    def job(self, jid):
        p = os.path.join(self.jobs, str(jid))
        for s in ('in', 'out'):
            os.makedirs(os.path.join(p, s), exist_ok=True)
        return p


# ======================================================================== job state (pickled per job)
class State:
    def __init__(self, tool):
        self.tool = tool
        self.docs = []
        self.ships = collections.OrderedDict()
        self.gi = []
        self.undo = []
        self.cells = {}          # edits typed on the sheet tabs: {excel file name: {json key: value}}
        self.lock = threading.RLock()


def strip_docs(docs):
    out = []
    for d in docs:
        d = dict(d)
        for k in DOC_DROP:
            d.pop(k, None)
        out.append(d)
    return out


def save_state(path, st):
    tmp = path + '.tmp'
    with open(tmp, 'wb') as f:
        pickle.dump(dict(tool=st.tool, docs=st.docs, ships=st.ships, gi=st.gi, undo=st.undo[-30:], cells=st.cells), f)
    os.replace(tmp, path)


def load_state(path, tool):
    st = State(tool)
    if os.path.exists(path):
        with open(path, 'rb') as f:
            d = pickle.load(f)
        st.docs, st.ships, st.gi, st.undo = d['docs'], d['ships'], d['gi'], d.get('undo', [])
        st.cells = d.get('cells') or {}
    return st


# ======================================================================== settings -> ctx / cfg
def merged_settings(saved):
    s = json.loads(json.dumps(DEFAULT_SETTINGS))
    for k, v in (saved or {}).items():
        if isinstance(v, dict) and isinstance(s.get(k), dict):
            s[k].update(v)
        else:
            s[k] = v
    tl = [str(x).strip().upper() for x in (s.get('teams') or []) if str(x).strip()]
    if 'SFA' not in tl:                  # SFA team (Japanese perishables - SFA rules) is always offered
        tl.append('SFA')
    if 'CLASQUIN' not in tl:             # Clasquin team (wine & spirits imports - liquor rules, ID permit = IPT DNG)
        tl.append('CLASQUIN')
    s['teams'] = tl
    return s


def template_path():
    return os.path.join(base_dir(), TEMPLATE)


_LISTS = {}


_tool_now = ['invoice']


def lists(tool):
    if tool not in _LISTS:
        _LISTS[tool] = engine(tool).template_lists(template_path())
    return _LISTS[tool]


def make_ctx(tool, P, S, form):
    E = engine(tool)
    o = S['options']
    _tool_now[0] = tool
    E.OCR['enabled'] = bool(o.get('ocr', True))
    if S.get('uom_rules'):
        E.set_uom_rules(S['uom_rules'])
    else:
        E.set_uom_rules([list(r) for r in E.DEFAULT_UOM_RULES])
    r = S['rules']
    cfg = dict(master=P.master(tool), hs_list=P.hs, default_uom=(r.get('default_uom') or '').upper(),
               blank_hs_not_in_list=o.get('blank_hs_not_in_list'), desc_template=E.DEFAULT_DESC_TEMPLATE,
               desc_max=int(E.fnum(r.get('desc_max')) or 512), brand=r.get('brand', ''),
               cargo=E.parse_cargo(form.get('cargo_text', '') or ''),
               party_name=form.get('party_name', ''), party_uen=(form.get('party_uen', '') or '').upper(),
               uen_book=S.get('uen_book') or {}, consolidate=form.get('consolidate') or r.get('consolidate', 'none'),
               direction=form.get('direction') or 'Auto', control_casc=S['control'].get('casc') or 'MISC',
               control_eccn=o.get('control_by_eccn', False),
               control_hs=[re.sub(r'\D', '', x) for x in re.split(r'[,;\s]+', S['control'].get('hs') or '') if re.sub(r'\D', '', x)],
               cont_tne_dec=int(E.fnum(r.get('cont_dec')) if E.fnum(r.get('cont_dec')) is not None else 2),
               hs_auto=o.get('hs_auto', True), hs_map=S.get('hs_map') or {}, type_in_desc=o.get('type_in_desc', True),
               charges_as_other=True, charges_mode=form.get('charges_mode') or 'other',
               ignore_line_nw=o.get('ignore_line_nw', False))
    hsl, master = cached_hs(tool, P.hs), cached_master(tool, P.master(tool))
    ctx = E.make_ctx(dict(cfg, hs_list='', master=''), lists(tool))
    ctx['hs_list'], ctx['master'] = hsl, master
    ctx.setdefault('cargo', cfg['cargo'])
    for k in ('hs_map', 'type_in_desc', 'charges_mode', 'charges_as_other', 'ignore_line_nw', 'cont_tne_dec', 'hs_auto'):
        ctx.setdefault(k, cfg[k])
    ctx['uen_names'] = S.get('uen_names') or {}          # registered names by UEN (ACRA open data)
    ctx['uen_names_on'] = (S.get('options') or {}).get('uen_names', True) and form.get('uen_names_on', True) is not False
    ctx['consolidate_explicit'] = bool(form.get('consolidate_explicit'))   # chosen on the Item page (beats supplier default)
    ctx['rules_off'] = list(form.get('rules_off') or [])        # customer rules unticked on the job
    ctx['hs_desc'] = hs_desc_indexes(P, S) if hasattr(E, 'HSDescIndex') else []
    ctx['casc_hs'] = casc_by_hs(P)
    ctx['permit_type'] = (form.get('permit_type') or '').strip().upper()      # OD / OT = export, IG / ME ... = import
    for d_ in (getattr(P, 'masters', ''), os.path.join(base_dir(), 'masters')):     # Clasquin past permits (liquor check)
        if d_ and os.path.exists(os.path.join(d_, 'CLASQUIN_MASTER_LIST.xlsx')):
            ctx['liquor_master'] = os.path.join(d_, 'CLASQUIN_MASTER_LIST.xlsx')
            break
    team = (form.get('team') or '').strip().upper()
    ctx['team'] = team
    ct = S.get('cust_teams') or {}
    lists_ = {f: m for f, m in cust_lists(P).items() if not team or not ct.get(f) or ct[f].upper() == team}
    ctx['cust_lists'] = lists_
    merged = {}
    for f, m in lists_.items():                    # customer product lists of this job's team
        for a, b in m.items():
            merged.setdefault(a, b)
    ctx['cust_master'] = merged
    ctx['cust_list'] = (form.get('cust_list') or '').strip()
    if team:                                       # HS lists by description of this team only (e.g. PANASONIC = KWE)
        tm = {(l.get('name') or l.get('file') or ''): hsd_team(l) for l in S.get('hs_desc_lists') or []}
        ctx['hs_desc'] = [x for x in ctx.get('hs_desc') or [] if tm.get(x[0], '') in ('', team)]
    ctx['blank_hs_not_in_list'] = True        # HS only when it (or its auto-mapped Singapore code) is in the UPDATE HS CODE list
    kw = S.get('hs_keywords')
    if isinstance(kw, list):
        ctx['hs_keywords'] = [(str(a or '').strip(), re.sub(r'\D', '', str(b or ''))) for a, b in
                              (x if isinstance(x, (list, tuple)) and len(x) >= 2 else ('', '') for x in kw) if str(a or '').strip()]
    return ctx


def hsd_team(l):
    """Team of an HS list by description ('' = every team). PANASONIC FACTORY = KWE team."""
    t = str(l.get('team') or '').strip().upper()
    if not t and 'PANASONIC' in str(l.get('name') or l.get('match') or '').upper():
        t = 'KWE'
    return t


def hs_desc_indexes(P, S):
    """[(name, match regex, HSDescIndex)] of the HS lists by description (cached; reloaded when a file changes)."""
    import permit_engine as PE
    out = []
    for l in S.get('hs_desc_lists') or []:
        path = os.path.join(P.hsdesc, l.get('file') or '')
        m = _mtime(path)
        if not l.get('file') or m is None:
            continue
        k = ('hsd', path)
        if k not in _CACHE or _CACHE[k][0] != m:
            try:
                _CACHE[k] = (m, PE.HSDescIndex(PE.load_hs_desc_list(path)))
            except Exception:
                continue
        out.append((l.get('name') or l['file'], l.get('match') or '', _CACHE[k][1]))
    return out


_CACHE = {}


def _mtime(p):
    try:
        return os.path.getmtime(p)
    except OSError:
        return None


def cached_hs(tool, path):
    """HS list loaded once per tool (reloaded when the file changes) - load_hs_list also fills the engine's
    product codes / control info."""
    k, m = ('hs', tool), _mtime(path)
    if k not in _CACHE or _CACHE[k][0] != m:
        _CACHE[k] = (m, engine(tool).load_hs_list(path))
    return _CACHE[k][1]


def cached_master(tool, path):
    k, m = ('master', tool), _mtime(path)
    if k not in _CACHE or _CACHE[k][0] != m:
        _CACHE[k] = (m, engine(tool).load_master(path))
    return _CACHE[k][1]


def hs_options(P, hs):
    """Singapore HS codes for a (supplier) HS - same for every tool."""
    E = engine('danfoss')
    return [[c, d] for c, d in E.hs_candidates(re.sub(r'\D', '', hs or ''), cached_hs('danfoss', P.hs))]


def decl_text(tool, code):
    return next((t for t, c in lists(tool).get('List-DeclarationType', []) if c == code), '')


_cfg_ships = [None]


def build_cfg(tool, P, S, form, out_dir, gi):
    E = engine(tool)
    o = S['options']
    hd = {k: (form.get('header') or {}).get(k, '') or '' for k, _ in HEADER_FIELDS}
    for k, v in (S.get('header') or {}).items():
        if not hd.get(k) and v and k not in ('mawb', 'flight', 'date', 'transport_id', 'loading_port_code', 'loading_port_name'):
            hd[k] = v
    for k in ('declaration_type', 'cargo_pack_type', 'transport', 'vessel_type_list'):
        hd[k] = (form.get('header') or {}).get(k, '') or ''
    pt = (form.get('permit_type') or '').upper()
    if not pt and form.get('decl_auto', True) is not False:          # declaration advisor: permit type it applied
        aps = {(sh.get('declaration') or {}).get('permit_applied') or '' for sh in (_cfg_ships[0] or [])}
        if len(aps) == 1 and list(aps)[0]:
            pt = list(aps)[0]
    if pt == 'ME' and not hd.get('receipt_code'):
        hd['receipt_code'] = 'ME'
    locs = S.get('locations') or LOCATIONS
    for side in ('release', 'receipt'):
        if hd.get(side + '_code') and not hd.get(side + '_name'):
            hd[side + '_name'] = locs.get(hd[side + '_code'].upper(), '')
    if not hd.get('declaration_type') and E.PERMIT_SCHEME.get(pt):
        hd['declaration_type'] = decl_text(tool, E.PERMIT_SCHEME[pt])
    mx = int(E.fnum(form.get('max_items') or S['rules'].get('max_items')) or 0) if form.get('split_excel') else 0
    cfg = dict(template=template_path(), out_dir=out_dir, permit_type=pt, header=hd, job_no=form.get('job_no', ''),
               remarks_template=(form.get('remarks_template') or S.get('remarks_template') or E.DEFAULT_REMARKS),
               file_pattern=S['rules'].get('file_pattern') or '{HAWB}', global_issues=gi,
               cargo=E.parse_cargo(form.get('cargo_text', '') or ''), insurance_pct=S.get('insurance_pct', '1'),
               insurance_cur=(S.get('insurance_cur') or 'SGD').upper(), forwarder_map=S.get('forwarder_map') or {},
               max_items=mx, **{k: o.get(k, True) for k in ('fill_term', 'fill_insurance', 'fill_container',
                                                              'fill_outer_pack', 'auto_cargo', 'auto_transport',
                                                              'fill_forwarder', 'item_mawb')})
    cfg['exchange_rate'], cfg['j_start'] = form.get('exchange_rate', ''), form.get('j_start', '') or '1'   # RS remarks
    cfg['split_hs'] = bool(form.get('split_hs'))
    cfg['uen_names'] = S.get('uen_names') or {}
    cfg['uen_names_on'] = (S.get('options') or {}).get('uen_names', True) and form.get('uen_names_on', True) is not False           # split Excel: one HS code stays in one Excel
    cfg['extra_header'] = dict(S.get('extra_header') or {}, **(form.get('extra_header') or {}))     # defaults + this job
    cfg['extra_invoice'] = dict(S.get('extra_invoice') or {}, **(form.get('extra_invoice') or {}))
    cfg['fill_freight'] = o.get('fill_freight', True)
    cfg['charges_as_other'] = True
    return cfg


# ======================================================================== steps
def extract(tool, P, S, form, paths, st, log, progress=None):
    E = engine(tool)
    progress = progress or (lambda **k: None)
    ctx = make_ctx(tool, P, S, form)
    prof = None
    if tool == 'invoice':
        try:
            import readers as R
            prof = R.Profiles(P.profiles)
        except Exception:
            prof = None
    docs = []
    for n, p in enumerate(paths):
        progress(stage='read', done=n, total=len(paths), file=os.path.basename(p))
        try:
            if tool == 'invoice':
                ds = E.parse_pdf(p, ctx['countries'], log=log, profiles=prof)
            else:
                ds = E.parse_pdf(p, ctx['countries'], log=log)
        except Exception as e:
            log('ERROR %s: %s' % (os.path.basename(p), e))
            continue
        docs.extend(ds)
        for d in ds:
            log(doc_line(E, d))
    progress(stage='match', done=len(paths), total=len(paths), file='')
    old = st.ships
    ships, gi = E.build_shipments(docs)
    ctx['ships_n'] = list(ships)
    for k, sh in ships.items():                      # keep edits of an earlier extract of this job
        if k in old:
            for f in ('file_name', 'gw_override', 'pkg_override', 'overrides', 'include', 'containers_edit', 'charges_mode'):
                if f in old[k]:
                    sh[f] = old[k][f]
            for f in ('hawb', 'party', 'party_uen', 'dest'):
                if old[k].get(f + '_src') == 'Entered':
                    sh[f], sh[f + '_src'] = old[k][f], 'Entered'
        E.resolve(sh, ctx)
        declare(sh, form, ctx)
    st.docs, st.ships, st.gi, st.undo = strip_docs(docs), ships, gi, []
    for fn, lvl, msg in gi:
        log('%s %s: %s' % (lvl, fn, msg))
    cg = [s.get('cargo') or {} for s in ships.values()]
    pt = next((c.get('permit_type') for c in cg if c.get('permit_type') in E.PERMIT_SCHEME), '')
    log('=' * 60)
    log('%d shipment(s) / HBL found, %d item row(s)' % (len(ships), sum(len(s['items']) for s in ships.values())))
    try:                                             # document extraction analysis (Check & files > Document analysis)
        import doc_analysis
        for ln in doc_analysis.log_lines(doc_analysis.analyse(st.docs, ships)):
            log(ln)
    except Exception as e:
        log('Document analysis: %s' % e)
    bm, who = brand_from_book(st, S)
    if bm and not form.get('brand_mode'):
        log('Brand: %s for %s (remembered - change it on the Item tab)' % (bm, who))
    return dict(permit_type=pt, job_no=next((c.get('job_no') for c in cg if c.get('job_no')), ''), brand_mode=bm)


def doc_line(E, d):
    k = d.get('kind')
    if k == 'INV':
        return 'INVOICE  %s  inv %s  %s  %s %s  %d line(s)  GW %s KG%s' % (
            d['file'], d.get('invoice_no') or '?', d.get('invoice_date', ''), d.get('currency', ''),
            '{:,.2f}'.format(d['total']) if d.get('total') is not None else '?', len(d.get('lines', [])),
            E.fmt_thou(d.get('total_gw')) or '?', '   CHECK read by OCR' if d.get('ocr') else '')
    if k == 'PL':
        return 'PACKING LIST  %s  %d row(s)  GW %s KG' % (d['file'], len(d.get('lines', [])), E.fmt_thou(d.get('total_gw')) or '?')
    if k == 'BL':
        return 'BILL OF LADING  %s  B/L %s  vessel %s %s  %s > %s  deliveries %s' % (
            d['file'], d.get('bl_no') or '?', d.get('vessel') or '?', d.get('voyage', ''), d.get('pol') or '?',
            d.get('pod') or '?', ', '.join(d.get('deliveries') or []) or '-')
    if k == 'EMAIL':
        return 'E-MAIL  %s  %s' % (d['file'], '  '.join('%s: %s' % (a.upper(), b) for a, b in (d.get('email') or {}).items()
                                                         if a not in ('staff', 'company_type')))
    if k == 'PL' and d.get('delivery'):
        return 'PACKING LIST  %s  delivery %s  GW %s KG  NW %s KG  pallets %s' % (
            d['file'], d['delivery'], E.fmt_thou(d.get('total_gw')) or '?', E.fmt_thou(d.get('total_nw')) or '?',
            d.get('pallets') or '?')
    if k == 'AWB':
        return '%s  %s  HBL/HAWB %s  OBL/MAWB %s' % ((d.get('doc_type') or 'WAYBILL').upper(), d['file'],
                                                   d.get('hawb') or '?', d.get('mawb') or '?')
    return 'CHECK  %s - %s' % (d.get('file'), '; '.join(d.get('notes') or ['not recognised']))


def declare(sh, form, ctx):
    """Declaration advisor (Module 5 rules): suggested declaration type + Declaration notes."""
    try:
        import declaration_rules
        try:
            allowed = list(lists(ctx.get('tool') or _tool_now[0]).get('List-DeclarationType') or []) or None
            if allowed is not None:
                allowed += [('Duty and GST', 'DNG')]       # ID permit = IPT DNG (liquor / tobacco)
        except Exception:
            allowed = None
        declaration_rules.advise(sh, form, ctx.get('hs_list') or {}, allowed)
    except Exception as e:                       # never stop a job for the advisor
        sh.setdefault('issues_all', []).append(('CHECK', 'Declaration advisor: %s' % e))


def resolve_all(tool, P, S, form, st):
    E = engine(tool)
    ctx = make_ctx(tool, P, S, form)
    ctx['ships_n'] = list(st.ships)
    for sh in st.ships.values():
        E.resolve(sh, ctx)
        declare(sh, form, ctx)
    return ctx


def resolve_one(tool, P, S, form, st, sh):
    E = engine(tool)
    ctx = make_ctx(tool, P, S, form)
    ctx['ships_n'] = list(st.ships)
    E.resolve(sh, ctx)
    declare(sh, form, ctx)


def push_undo(st, key):
    sh = st.ships[key]
    st.undo.append((key, {f: copy.deepcopy(sh.get(f)) for f in SNAP if f in sh}))
    del st.undo[:-30]


def undo(st):
    while st.undo:
        key, snap = st.undo.pop()
        sh = st.ships.get(key)
        if sh is None:
            continue
        if {f: sh.get(f) for f in SNAP if f in sh} == snap:
            continue
        for f in SNAP:
            if f in snap:
                sh[f] = snap[f]
            else:
                sh.pop(f, None)
        return key
    return None


ITEM_FIELDS = ('desc', 'hs', 'uom', 'coo', 'casc', 'nw', 'brand')


def set_item_field(sh, idxs, field, value):
    v = (value or '').strip()
    if field in ('uom', 'coo', 'casc', 'brand'):
        v = v.upper()
    if field == 'hs':
        v = re.sub(r'\D', '', v)                # 8431.49.90 -> 84314990
        if len(v) > 8:
            v = v[:8]                           # 10-digit supplier / import HS -> the 8-digit Singapore HS
        if v and len(v) != 8:
            raise ValueError('HS code %s has %d digits - type the 8-digit Singapore HS code (the HS codes page shows the '
                             'codes of a 6-digit group)' % (v, len(v)))
    for i in idxs:
        it = sh['items'][int(i)]
        for k in it['keys']:
            sh['overrides'].setdefault(k, {})[field] = v


def clear_item_edits(sh, idxs):
    for i in idxs:
        for k in sh['items'][int(i)]['keys']:
            ov = sh['overrides'].get(k, {})
            for f in list(ov):
                if f != 'sep':
                    ov.pop(f)


def set_sep(sh, idxs, on):
    for i in idxs:
        for k in sh['items'][int(i)]['keys']:
            if on:
                sh['overrides'].setdefault(k, {})['sep'] = True
            else:
                sh['overrides'].get(k, {}).pop('sep', None)


SHIP_FIELDS = ('hawb', 'party', 'party_uen', 'dest', 'gw', 'pkg', 'file_name', 'include', 'charges_mode', 'containers')


def set_ship_field(E, sh, field, value, S=None):
    if field == 'gw':
        sh['gw_override'] = E.fnum(value) if str(value or '').strip() else None
    elif field == 'pkg':
        sh['pkg_override'] = int(E.fnum(value)) if str(value or '').strip() and E.fnum(value) else None
    elif field == 'file_name':
        sh['file_name'] = str(value or '').replace('.xlsx', '').strip()
    elif field == 'include':
        sh['include'] = bool(value)
    elif field == 'charges_mode':
        sh['charges_mode'] = value if value in ('last', 'other', 'otc') else 'other'
    elif field == 'containers':
        if value is None:
            sh.pop('containers_edit', None)
        else:
            rows = []
            for c in value:
                no = re.sub(r'\s', '', str(c.get('no', ''))).upper()
                if not no:
                    continue
                size = str(c.get('size', '')).upper()
                m = re.search(r'(20|40|45)', size)
                tne = E.fnum(c.get('tne')) if str(c.get('tne', '')).strip() else None
                if tne is not None and tne > 999:
                    raise ValueError('Container weight must be in tonnes (e.g. 7.83)')
                rows.append(dict(no=no, size=m.group(1) if m else '', size_text=size, tne=tne,
                                 seal=str(c.get('seal', '')).upper()))
            sh['containers_edit'] = rows
    elif field in ('hawb', 'dest', 'party', 'party_uen'):
        sh[field], sh[field + '_src'] = str(value or '').strip().upper(), 'Entered'
        if field == 'party_uen' and value and sh.get('party') and S is not None:
            S.setdefault('uen_book', {})[sh['party'].upper()] = str(value).upper()
    else:
        raise ValueError('unknown field ' + field)


def container_tne(sh, dec=2):
    rows = sh.get('containers_edit') or [dict(c) for c in sh.get('containers', [])]
    if not rows or not sh.get('total_gw'):
        raise ValueError('Needs a container and the gross weight of the shipment')
    each = sh['total_gw'] / 1000.0 / len(rows)
    for r in rows:
        r['tne'] = math.ceil(round(each * 10 ** dec, 6)) / 10 ** dec
    sh['containers_edit'] = rows


def generate(tool, P, S, form, st, out_dir, log):
    E = engine(tool)
    for f in os.listdir(out_dir):
        try:
            os.remove(os.path.join(out_dir, f))
        except OSError:
            pass
    resolve_all(tool, P, S, form, st)
    _cfg_ships[0] = list(st.ships.values())
    cfg = build_cfg(tool, P, S, form, out_dir, st.gi)
    ships = list(st.ships.values())
    res = E.generate(ships, cfg, log=log)
    nb, nd = apply_brand(st, res[0], form)
    apply_declaration(res[0], form, log)
    mode = brand_mode(form)
    if nd:
        log('Brand: from the documents on %d item row(s)' % nd)
    if nb - nd > 0 and mode:
        log('Brand: %s on %d item row(s)%s' % (mode, nb - nd, ' without a brand in the documents' if nd else ''))
    n_cells = apply_all_cells(st, res[0], log)
    if n_cells:
        log('%d value(s) typed on the sheet tabs written into the Excel' % n_cells)
    n_dir = strip_other_direction(res[0])
    if n_dir:
        log('%d value(s) of the other direction removed (export permit = export fields only, import = import fields only)' % n_dir)
    if (S.get('options') or {}).get('carrier_by_uen', True):
        apply_carrier_book(res[0], S, log)
    rezip_excels(res[0], out_dir)
    o = S['options']
    if o.get('xml', True):
        xs = []
        files_ = res[0]
        xo = dict(snd=S['xml'].get('snd') or 'ITG_INDIA', rcp=S['xml'].get('rcp') or 'YCH',
                  msg_import=S['xml'].get('msg_import') or 'IPTDEC', msg_export=S['xml'].get('msg_export') or '',
                  charges=o.get('xml_charges', False))
        for f_ in files_:
            xof = dict(xo)
            mt, mt_how = msg_type_for(f_, form, S)
            if mt:
                xof['msg_export' if file_direction(f_) == 'EXPORT' else 'msg_import'] = mt
                log('%s  -  XML message type %s (%s)' % (f_['name'], mt, mt_how))
            job = (cfg.get('job_no') or '').strip()
            job = (job + ('_' + os.path.splitext(f_['name'])[0] if len(files_) > 1 else '')) if job else \
                os.path.splitext(f_['name'])[0]
            xp = X.excel_to_xml(f_['path'], os.path.join(out_dir, 'FrtImport_%s.xml' % job), job, xof)
            xs.append(xp)
            need = X.check_values(f_['path'])
            log('%s  -  XML of %s%s' % (os.path.basename(xp), f_['name'], ('   CHECK blank: ' + ', '.join(need)) if need else ''))
        if xs:
            X.zip_xmls(xs, os.path.join(out_dir, 'XML_ALL.zip'))
            log('XML_ALL.zip  -  %d XML file(s)' % len(xs))
    if o.get('raw', True) and st.docs:
        E.write_raw(os.path.join(out_dir, 'RAW_DATA.xlsx'), st.docs, [s for s in ships if s.get('include', True)])
        try:
            import doc_analysis
            doc_analysis.write_sheet(os.path.join(out_dir, 'RAW_DATA.xlsx'), doc_analysis.analyse(st.docs, st.ships))
        except Exception as e:
            log('CHECK document analysis sheet: %s' % e)
        log('RAW_DATA.xlsx  -  everything read from the documents + DOCUMENT_ANALYSIS / DOCUMENT_CHECKS')
    log('DONE')
    return res


# ======================================================================== JSON views for the web page
def _f(v, n=None):
    if v is None:
        return None
    try:
        return round(float(v), n) if n is not None else float(v)
    except (TypeError, ValueError):
        return None


def supplier_of(sh):
    """Name the brand choice is remembered under: the seller on the invoice (else the importer / exporter)."""
    i0 = (sh.get('invs') or [{}])[0]
    for k in ('supplier', 'invoice_from', 'ship_from', 'exporter', 'seller'):
        v = i0.get(k) if isinstance(i0, dict) else ''
        if v and isinstance(v, str):
            return re.sub(r'\s+', ' ', v).strip().upper()
    return (sh.get('party') or '').strip().upper()


def brand_from_book(st, S):
    """Brand choice remembered for the supplier of this job (e.g. UNBRANDED for SHELL EASTERN TRADING) - ''."""
    book = S.get('brand_book') or {}
    for sh in st.ships.values():
        v = book.get(supplier_of(sh))
        if v:
            return v, supplier_of(sh)
    return '', ''


def item_json(E, sh, i, it):
    bad = not it['hs'] or not it['coo'] or it.get('hs_qty') is None or not it['hs_uom']
    pc = [c[0] for c in (it.get('pcodes') or [])]
    hr = it.get('hs_row') or ()
    edited = any((sh.get('overrides') or {}).get(k, {}).keys() - {'sep'} for k in it.get('keys', []))
    return dict(i=i, item=it['item'], invoice=it.get('invoice', ''), lines=', '.join(it.get('lines') or []),
                desc=it['description'], inv_desc=it.get('inv_desc') or it.get('desc', ''),
                qty=_f(it.get('qty')), uom=it.get('uom', ''), price=None if it.get('merged') else _f(it.get('price')),
                amount=_f(it.get('amount'), 2), currency=it.get('currency') or sh.get('currency', ''),
                nw=_f(it.get('nw'), 4), nw_src=it.get('nw_src', ''), gw=_f(it.get('gw'), 4),
                doc_hs=it.get('doc_hs8') or it.get('doc_hs') or '', hs=it['hs'], hs_src=it.get('hs_src', ''),
                hs_desc=hr[2] if len(hr) > 2 else '', hs_uom=it['hs_uom'], uom_src=it.get('uom_src', ''),
                hs_qty=_f(it.get('hs_qty'), 4), coo=it['coo'], coo_src=it.get('coo_src', ''), eccn=it.get('eccn', ''),
                casc=it.get('casc', ''), casc_src=it.get('casc_src', ''), controlled=bool(it.get('controlled')),
                ca=it.get('ca', ''), idamisc='IDAMISC' in pc, order=it.get('order', ''), delivery=it.get('delivery', ''),
                src=it.get('src', ''), merged=bool(it.get('merged')), sep=bool(it.get('sep')), edited=bool(edited), bad=bad,
                desc_edited=bool(it.get('desc_edit')), brand=it.get('brand', ''))


def pl_rows(sh, it):
    out = []
    for p in sh.get('pls') or []:
        for l in p.get('lines', []):
            if l.get('item') == it['item'] and (not it.get('delivery') or not l.get('delivery') or l['delivery'] == it['delivery']):
                out.append(dict(ref=l.get('delivery') or p.get('file', ''), qty=_f(l.get('qty')), gw=_f(l.get('gw')),
                                nw=_f(l.get('nw')), boxes=', '.join(l.get('boxes') or [])[:60]))
    return out


def ship_json(E, key, sh, S, full=True):
    cg = sh.get('cargo') or {}
    nw = sum(i['nw'] or 0 for i in sh['items'])
    iss = sh.get('issues_all') or []
    try:
        fname = E.file_name_for([sh], {'file_pattern': S['rules'].get('file_pattern') or '{HAWB}'}) + '.xlsx'
    except Exception:
        fname = (sh.get('hawb') or key) + '.xlsx'
    d = dict(key=key, include=sh.get('include', True), file=fname, hawb=sh.get('hawb') or '', hawb_src=sh.get('hawb_src', ''),
             direction=sh.get('direction', ''), invoices=[i.get('invoice_no') or i.get('file') for i in sh['invs']],
             mawb=cg.get('mawb', ''), flight=cg.get('flight', ''), date=cg.get('date', ''), party=sh.get('party', ''),
             party_uen=sh.get('party_uen', ''), uen_src=sh.get('party_uen_src', ''), dest=sh.get('dest', ''),
             incoterm=sh.get('incoterm', ''), currency=sh.get('currency', ''), total=_f(sh.get('invoice_total'), 2),
             gw=_f(sh.get('total_gw'), 3), gw_src=sh.get('gw_src', ''), nw=round(nw, 3), packages=sh.get('packages'),
             lines=sh.get('lines_n', 0), rows=len(sh['items']),
             problems=sum(1 for l, _ in iss if l in ('MISMATCH', 'MISSING', 'ERROR')) +
             sum(1 for it in sh['items'] if not it['hs'] or not it['coo'] or it.get('hs_qty') is None or not it['hs_uom']),
             checks=sum(1 for l, _ in iss if l == 'CHECK'),
             controlled=sum(1 for it in sh['items'] if it.get('controlled')),
             charges=_f(sh.get('charges')), charges_mode=sh.get('charges_mode', ''), declaration=sh.get('declaration'), consolidate_used=sh.get('consolidate_used', ''),
             has_charges=any(i.get('charges') for i in sh['invs']),
             rules=[dict(r, **{k: v for k, v in (getattr(E, 'RULE_BY_ID', {}).get(r['id']) or {}).items() if k != 'who'})
                    for r in sh.get('rules') or []])
    if full:
        d['items'] = [item_json(E, sh, i, it) for i, it in enumerate(sh['items'])]
        seen, il = set(), []
        for l, m in iss:
            if (l, m) not in seen:
                seen.add((l, m))
                il.append([l, m])
        d['issues'] = il
        d['cargo'] = {k: (', '.join(str(x.get('no', x)) if isinstance(x, dict) else str(x) for x in v)
                          if isinstance(v, list) else v) for k, v in cg.items()
                      if isinstance(v, (str, int, float, list)) and not k.endswith('_qty')}
        d['containers'] = [dict(no=c.get('no', ''), size=c.get('size_text') or c.get('size', ''),
                                tne=_f(c.get('tne')), seal=c.get('seal', ''))
                           for c in (sh.get('containers_edit') if sh.get('containers_edit') is not None else sh.get('containers', []))]
        d['containers_edited'] = sh.get('containers_edit') is not None
        d['pls'] = [p.get('file', '') for p in sh.get('pls') or []]
        d['awbs'] = ['%s %s' % (a.get('doc_type', 'waybill'), a.get('file', '')) for a in sh.get('awbs') or []]
    return d


def summary(E, st, S):
    act = [s for s in st.ships.values() if s.get('include', True)]
    vals = collections.OrderedDict()
    for s in act:
        curs = {it.get('currency') for it in s['items'] if it.get('currency')}
        if len(curs) > 1:                     # several currencies: value per currency
            for it in s['items']:
                c = it.get('currency') or '?'
                vals[c] = vals.get(c, 0) + (it['amount'] or 0)
            continue
        v = s.get('invoice_total')
        if v is None:
            v = sum(i['amount'] or 0 for i in s['items'])
        vals[s.get('currency') or '?'] = vals.get(s.get('currency') or '?', 0) + (v or 0)
    sj = [ship_json(E, k, s, S, full=False) for k, s in st.ships.items()]
    return dict(ships=len(act), items=sum(len(s['items']) for s in act),
                value='  '.join('%s %s' % (k, '{:,.2f}'.format(v)) for k, v in vals.items()),
                gw=round(sum(s.get('total_gw') or 0 for s in act), 3),
                problems=sum(x['problems'] for x in sj if x['include']),
                checks=sum(x['checks'] for x in sj if x['include']),
                controlled=sum(x['controlled'] for x in sj if x['include']),
                hawbs=', '.join(s.get('hawb') or '' for s in act if s.get('hawb')),
                invoices=', '.join(i for x in sj for i in x['invoices'] if i))


# ======================================================================== validation (after extraction)
IMPORT_NEED = (('mawb', 'MAWB / OBL'), ('transport_id', 'Vessel / airline'), ('flight', 'Voyage / flight no.'),
               ('date', 'Arrival date'), ('loading_code', 'Loading port'), ('release', 'Release location'),
               ('receipt', 'Receipt location'))
EXPORT_NEED = (('mawb', 'MAWB / OBL'), ('transport_id', 'Vessel / airline'), ('flight', 'Voyage / flight no.'),
               ('date', 'Departure date'), ('discharge_code', 'Discharge port'))
HEAD_KEY = {'loading_code': 'loading_port_code', 'release': 'release_code', 'receipt': 'receipt_code',
            'discharge_code': 'discharge_port_code'}


def validate(E, st, form, S):
    """Checks per shipment -> [{ship, id, level ok/warn/error, title, detail, fix{tab, item}}] + summary."""
    out = []
    hd = form.get('header') or {}
    pt = (form.get('permit_type') or '').upper()
    for key, sh in st.ships.items():
        if not sh.get('include', True):
            continue
        name = sh.get('hawb') or key
        cg = sh.get('cargo') or {}
        its = sh['items']
        exp = sh.get('direction') == 'EXPORT'

        def add(cid, level, title, detail='', tab='', item=None, group='Shipment'):
            out.append(dict(ship=key, name=name, id=cid, level=level, title=title, detail=detail, group=group,
                            fix=dict(tab=tab, item=item) if tab else None))
        # ---- invoice value
        amt = round(sum(i['amount'] or 0 for i in its), 2)
        tot = sh.get('invoice_total')
        chg = sh.get('charges') or 0
        tol = max(0.05, len(its) * 0.005)          # rounding of the item amounts only
        if tot is None:
            add('value', 'warn', 'Invoice total not found', 'Item total %s %s' % (sh.get('currency', ''), '{:,.2f}'.format(amt)),
                'items', group='Invoice')
        elif abs(amt - tot) <= tol:
            add('value', 'ok', 'Item values match the invoice total', '%s %s' % (sh.get('currency', ''), '{:,.2f}'.format(tot))
                + (' (charges %s added to the last item)' % '{:,.2f}'.format(chg) if sh.get('charges_in_items') else ''), group='Invoice')
        elif chg and abs(amt + chg - tot) <= tol:
            mode = sh.get('charges_mode') or 'other'
            if mode == 'otc':
                add('value', 'ok', 'Items + Other Taxable Charge = invoice total',
                    'items %s + charge %s = %s %s' % ('{:,.2f}'.format(amt), '{:,.2f}'.format(chg), sh.get('currency', ''),
                                                     '{:,.2f}'.format(tot)), group='Invoice')
            else:
                add('value', 'warn', 'Other charges %s are only in the invoice total, not in any item' % '{:,.2f}'.format(chg),
                    'Excel: Invoice total %s, items add up to %s. Choose on the Shipments tab: add to the last item, or '
                    'declare as Other Taxable Charge' % ('{:,.2f}'.format(tot), '{:,.2f}'.format(amt)), 'ships', group='Invoice')
        else:
            add('value', 'warn', 'Item values do not add up to the invoice total',
                'items %s vs invoice %s %s' % ('{:,.2f}'.format(amt), sh.get('currency', ''), '{:,.2f}'.format(tot)),
                'items', group='Invoice')
        if not sh.get('currency'):
            add('cur', 'error', 'Currency missing', '', 'header', group='Invoice')
        # ---- items
        for fid, title, test in (
                ('hs', 'HS code', lambda i: not i['hs']),
                ('uom', 'HS UOM', lambda i: not i['hs_uom']),
                ('hq', 'HS quantity', lambda i: i.get('hs_qty') is None),
                ('coo', 'Country of origin', lambda i: not i['coo']),
                ('nw', 'Net weight', lambda i: not i.get('nw'))):
            bad = [n for n, i in enumerate(its) if test(i)]
            if bad:
                add('items_' + fid, 'error' if fid != 'nw' else 'warn',
                    '%s missing on %d of %d item(s)' % (title, len(bad), len(its)),
                    ', '.join(its[n]['item'] for n in bad[:8]) + (' ...' if len(bad) > 8 else ''), 'detail', bad[0], group='Items')
            else:
                add('items_' + fid, 'ok', '%s on every item' % title, '', group='Items')
        ctl = [n for n, i in enumerate(its) if i.get('controlled')]
        noc = [n for n in ctl if not its[n].get('casc')]
        if ctl:
            add('casc', 'error' if noc else 'ok', ('%d controlled item(s) without CASC code' % len(noc)) if noc else
                '%d controlled item(s) have a CASC code' % len(ctl),
                ', '.join(its[n]['item'] for n in (noc or ctl)[:8]), 'detail' if noc else '', (noc or [None])[0], group='Items')
        auto = sum(1 for i in its if 'auto-mapped' in (i.get('hs_src') or ''))
        if auto:
            add('hs_auto', 'warn', '%d HS code(s) auto-mapped from the supplier code' % auto,
                'Check them, or save a mapping on the HS codes page', 'items', group='Items')
        # ---- weights / packages
        gw, nw = sh.get('total_gw'), sum(i.get('nw') or 0 for i in its)
        if not gw:
            add('gw', 'error', 'Gross weight missing', 'Type it on the Shipments tab', 'ships', group='Cargo')
        elif nw and nw > gw * 1.001:
            add('gw', 'error', 'Net weight is more than gross weight', 'net %.3f > gross %.3f KG' % (nw, gw), 'ships', group='Cargo')
        else:
            add('gw', 'ok', 'Gross weight %s KG' % E.fmt_thou(gw), 'net %s KG' % E.fmt_thou(nw), group='Cargo')
        if sh.get('packages'):
            add('pkg', 'ok', '%s package(s)' % sh['packages'], group='Cargo')
        else:
            add('pkg', 'error', 'Number of packages missing', 'Type it on the Shipments tab', 'ships', group='Cargo')
        conts = sh.get('containers_edit') if sh.get('containers_edit') is not None else sh.get('containers') or []
        if conts and any(c.get('tne') in (None, '') for c in conts):
            add('cont', 'warn', 'Container weight (TNE) missing', ', '.join(c.get('no', '') for c in conts), 'header', group='Cargo')
        # ---- parties / header
        if sh.get('hawb'):
            add('hawb', 'ok', 'HBL / HAWB %s' % sh['hawb'], sh.get('hawb_src', ''), group='Shipment')
        else:
            add('hawb', 'error', 'HBL / HAWB missing', 'Add the B/L / waybill or type it on the Shipments tab', 'ships', group='Shipment')
        if sh.get('party_uen'):
            add('uen', 'ok', '%s UEN %s' % ('Exporter' if exp else 'Importer', sh['party_uen']), sh.get('party', ''), group='Shipment')
        else:
            add('uen', 'error', '%s UEN missing' % ('Exporter' if exp else 'Importer'), sh.get('party', ''), 'ships', group='Shipment')
        if pt or hd.get('declaration_type'):
            add('permit', 'ok', 'Permit type %s' % (pt or hd.get('declaration_type')), group='Shipment')
        else:
            add('permit', 'warn', 'Permit type not chosen', 'Header and cargo tab', 'header', group='Shipment')
        if exp and not sh.get('dest'):
            add('dest', 'error', 'Destination country missing', '', 'ships', group='Shipment')
        need = EXPORT_NEED if exp else IMPORT_NEED
        miss = [lab for k, lab in need if not (hd.get(HEAD_KEY.get(k, k)) or cg.get(k))]
        if pt == 'ME' and 'Receipt location' in miss:
            miss.remove('Receipt location')
        if miss:
            add('cargo', 'warn', '%d cargo detail(s) not found' % len(miss), ', '.join(miss), 'header', group='Cargo')
        else:
            add('cargo', 'ok', 'Cargo details complete', '', group='Cargo')
        # ---- customer rules (Customer rules tab)
        rl = sh.get('rules') or []
        if rl:
            E_ = sys.modules.get('permit_engine')
            rname = lambda r: ((getattr(E_, 'RULE_BY_ID', {}) or {}).get(r['id']) or {}).get('title', r['id'])
            used = [rname(r) for r in rl if r['on']]
            offr = [rname(r) for r in rl if not r['on']]
            add('rules', 'ok' if not offr else 'warn',
                'Customer rules: %d used%s' % (len(used), ', %d unticked' % len(offr) if offr else ''),
                '; '.join(used) + ((' | not used: ' + '; '.join(offr)) if offr else ''), 'rules', group='Customer rules')
        # ---- brand
        bm = brand_mode(form)
        nob = [n for n, i in enumerate(its) if not i.get('brand')]
        withb = len(its) - len(nob)
        doc = ' - %d item(s) keep the brand from the documents' % withb if withb else ''
        if bm == 'UNBRANDED':
            add('brand', 'ok', 'Unbranded on every item without a brand (chosen on the Item tab)' + doc, group='Items')
        elif bm:
            add('brand', 'ok', 'Brand %s on every item without a brand (chosen on the Item tab)%s' % (bm, doc), group='Items')
        elif nob:
            add('brand', 'warn', 'Brand not in the documents for %d of %d item(s)' % (len(nob), len(its)),
                'Item tab: type the brand for all items, or choose "Unbranded for all"', 'sh:Item', nob[0], group='Items')
        else:
            add('brand', 'ok', 'Brand on every item (documents / product master)', group='Items')
        # ---- document notes
        mm = [m for l, m in sh.get('issues_all', []) if l in ('MISMATCH', 'ERROR')]
        if mm:
            add('mismatch', 'error', '%d mismatch(es) between the documents' % len(mm), mm[0][:200], 'notes', group='Documents')
        ck = [m for l, m in sh.get('issues_all', []) if l == 'CHECK']
        if ck:
            add('checks', 'warn', '%d note(s) to read' % len(ck), ck[0][:200], 'notes', group='Documents')
    n = len(out)
    cnt = {k: sum(1 for c in out if c['level'] == k) for k in ('ok', 'warn', 'error')}
    return dict(checks=out, count=cnt, score=round(100.0 * cnt['ok'] / n) if n else 0,
                status='ready' if not cnt['error'] and not cnt['warn'] else 'review' if not cnt['error'] else 'fix')


# ======================================================================== plugins: new suppliers without rebuilding the exe
PLUGINS = []          # [{file, name, kind, status, detail}]


def plugin_dir(data):
    p = os.path.join(data, 'plugins')
    os.makedirs(p, exist_ok=True)
    src = os.path.join(base_dir(), 'plugins')         # plugins shipped with the server: installed once, never overwritten
    if os.path.isdir(src) and os.path.abspath(src) != os.path.abspath(p):
        for fn in os.listdir(src):
            if fn.endswith('.py') and not fn.startswith('_') and not os.path.exists(os.path.join(p, fn)):
                try:
                    shutil.copy(os.path.join(src, fn), os.path.join(p, fn))
                except OSError:
                    pass
    return p


def load_plugins(folder, log=print):
    """Every .py in PERMIT_SERVER_DATA/plugins is loaded (files starting with _ are skipped):
      READER plugin  (NAME, match(text), read(path, pages, C))  -> a new known layout for "Any supplier" / Auto
      TOOL plugin    (TOOL = dict(key, name, ...) + the engine functions)  -> a new tool on the New job page
    Called at start and by Settings > Suppliers > Reload."""
    import importlib.util, export_docs
    global PLUGINS
    builtin = [r for r in export_docs.READERS if not getattr(r[2], '_plugin', False)]
    readers_ = []
    for k in [k for k in TOOLS if k not in BUILTIN_TOOLS]:
        TOOLS.pop(k, None)
        _ENG.pop(k, None)
        _LISTS.pop(k, None)
    out = []
    for fn in sorted(os.listdir(folder)) if os.path.isdir(folder) else []:
        if not fn.endswith('.py') or fn.startswith('_'):
            continue
        path = os.path.join(folder, fn)
        modname = 'plugin_' + re.sub(r'\W', '_', fn[:-3])
        rec = dict(file=fn, name=fn[:-3], kind='?', status='error', detail='')
        try:
            spec = importlib.util.spec_from_file_location(modname, path)
            mod = importlib.util.module_from_spec(spec)
            sys.modules[modname] = mod
            spec.loader.exec_module(mod)
            if hasattr(mod, 'TOOL'):
                t = dict(mod.TOOL)
                key = re.sub(r'\W', '', str(t.get('key') or fn[:-3])).lower()
                if key in BUILTIN_TOOLS:
                    raise ValueError('TOOL key "%s" is a built-in tool - use another key' % key)
                for need in ('parse_pdf', 'build_shipments', 'resolve', 'generate', 'make_ctx', 'template_lists',
                             'load_hs_list', 'load_master', 'save_master'):
                    if not hasattr(mod, need):
                        raise ValueError('TOOL plugin needs a function %s()' % need)
                TOOLS[key] = dict(name=t.get('name') or key, engine=modname, module=mod,
                                  master=t.get('master') or key.upper() + '_PRODUCT_MASTER.xlsx',
                                  color=t.get('color') or '#6A1B9A', charges=bool(t.get('charges')),
                                  hint=t.get('hint') or 'plugin ' + fn, detect=getattr(mod, 'detect', None), plugin=fn)
                rec.update(kind='tool', name=TOOLS[key]['name'], status='ok', detail='New job > ' + TOOLS[key]['name'])
            elif hasattr(mod, 'read') and hasattr(mod, 'match'):
                name = str(getattr(mod, 'NAME', fn[:-3])).upper()
                if name in {r[0] for r in builtin}:
                    rec.update(kind='reader', name=name, status='ok',
                               detail='now built in - this plugin file is not used (you can delete it)')
                    out.append(rec)
                    log('PLUGIN %-28s reader SKIPPED (built in now)' % fn)
                    continue
                rd = mod.read

                def wrap(path_, pages, C, _rd=rd):
                    return _rd(path_, pages, C)
                wrap._plugin = True
                readers_.append((name, mod.match, wrap))
                rec.update(kind='reader', name=name, status='ok',
                           detail='Any supplier / Auto: ' + (getattr(mod, '__doc__', '') or '').strip().split('\n')[0][:160])
            else:
                raise ValueError('not a plugin: needs NAME + match(text) + read(path, pages, C), or TOOL = dict(...)')
        except Exception as e:
            rec['detail'] = '%s: %s' % (type(e).__name__, e)
        out.append(rec)
        log('PLUGIN %-28s %-6s %s %s' % (fn, rec['kind'], rec['status'].upper(), rec['detail']))
    export_docs.READERS = tuple(readers_) + tuple(builtin)        # plugins first: they win over a built-in layout
    PLUGINS = out
    return out


def builtin_layouts():
    import export_docs
    return [r[0] for r in export_docs.READERS if not getattr(r[2], '_plugin', False)]


def test_file(tool, P, S, path, log):
    """Settings > Suppliers > Test: what the reader makes of one PDF (nothing is saved)."""
    res = dict(tool=tool, layout='', docs=[])
    if tool == 'auto':
        tool = detect_tool([path])
    res['tool'] = tool
    E = engine(tool)
    ctx = make_ctx(tool, P, S, {})
    docs = E.parse_pdf(path, ctx['countries'], log=log)
    res['layout'] = next((d.get('reader') for d in docs if d.get('reader')), '') if tool == 'invoice' else TOOLS[tool]['name']
    for d in docs:
        lines = d.get('lines') or []
        res['docs'].append(dict(
            kind=d.get('kind') or '?', reader=d.get('reader', ''), file=d.get('file', ''),
            no=d.get('invoice_no') or d.get('bl_no') or d.get('delivery') or d.get('hawb') or '',
            currency=d.get('currency', ''), total=d.get('total'), lines=len(lines),
            gw=d.get('total_gw') or d.get('gw'), nw=d.get('total_nw'), notes=(d.get('notes') or [])[:12],
            sample=[{k: l.get(k) for k in ('item', 'desc', 'qty', 'uom', 'price', 'amount', 'hs', 'coo', 'nw')}
                    for l in lines[:12]]))
    return res



# ======================================================================== sheet tabs (one tab per sheet of the dec-upload Excel)
SHEET_TABS = ('Header', 'Invoice', 'Item', 'Casc', 'Container', 'Items to be consolidated')
ITEM_FIELD_COLS = {'Country of Origin': 'coo', 'HS Code': 'hs', 'HS UOM': 'uom', 'CASC Product Code 1': 'casc'}


def _item_key(it):
    return (it.get('keys') or ['%s|%s' % (it.get('invoice', ''), it.get('item', ''))])[0]


def file_rowmap(st, f):
    """Item sheet row -> [ship key, item index, item key] for one generated Excel."""
    where = {}
    for k, sh in st.ships.items():
        for i, it in enumerate(sh['items']):
            where[id(it)] = (k, i, _item_key(it))
    its = f.get('items') or [it for s in f.get('ships') or [] for it in s['items']]
    out = {}
    for r, it in enumerate(its, 2):
        w = where.get(id(it))
        if w:
            out[r] = list(w)
    return out


def cell_key(sheet, row, col, rowmap=None):
    if sheet == 'Item' and rowmap and row in rowmap:
        return json.dumps([sheet, ['@', rowmap[row][0], rowmap[row][2]], col])
    return json.dumps([sheet, row, col])


def _write_cells(path, edits):
    """edits: [(sheet, row, column header, value)] -> written into the Excel (XML level, template kept)."""
    import msd_engine as M
    if not edits:
        return 0
    tpl = M.Template(path)
    by = collections.OrderedDict()
    for sh, r, c, v in edits:
        by.setdefault(sh, []).append((r, c, v))
    changed, n = {}, 0
    for name, ed in by.items():
        if name not in tpl.sheets:
            continue
        sh = tpl.sheet(name)
        H = tpl.headers(sh)
        for r, c, v in ed:
            col = H.get(str(c).strip().lower())
            if not col:
                continue
            ref = '%s%d' % (col, r)
            old = sh.get(ref) or ''
            if '<f' in old:
                continue
            v = '' if v is None else str(v).strip()
            if v == '':
                row = sh.rows.get(r)
                if row and col in row[1]:
                    row[1].pop(col)
                    row[2] = True
                n += 1
                continue
            style = None
            numeric_old = bool(old) and '<v>' in old and 't="s"' not in old and 'inlineStr' not in old and 't="str"' not in old
            if 'date' in str(c).lower() and re.fullmatch(r'[\d./-]{6,10}', v):
                v, style = M._date(v, tpl)
            elif re.fullmatch(r'-?\d+(\.\d+)?', v.replace(',', '')) and (numeric_old or re.search(
                    r'amount|weight|qty|price|%|number of|tonnage|seq|item number|copies|hs code', str(c), re.I)):
                f = float(v.replace(',', ''))
                v = int(f) if f.is_integer() and 'amount' not in str(c).lower() and 'price' not in str(c).lower() else f
            sh.set(ref, v, style)
            n += 1
        changed[tpl.sheets[name]] = sh.tostring()
    tpl.save(path, changed)
    return n


def file_edits(st, fname, rowmap):
    out, orphan = [], 0
    rev = {(v[0], v[2]): r for r, v in rowmap.items()}
    for k, v in (st.cells.get(fname) or {}).items():
        sheet, row, col = json.loads(k)
        if isinstance(row, list):
            row = rev.get((row[1], row[2]))
            if row is None:
                orphan += 1
                continue
        out.append((sheet, int(row), col, v))
    return out, orphan


def file_direction(f):
    """'EXPORT' / 'IMPORT' of one output Excel (its first shipment)."""
    return next((s.get('direction') for s in f.get('ships') or [] if s.get('direction')), '') or 'IMPORT'


def same_direction(edits, direction):
    """Typed Header / Item cells of the other direction (e.g. Inward MAWB on an export) are not written."""
    import permit_engine as _E
    exp = direction == 'EXPORT'
    return [e for e in edits if e[0] not in ('Header', 'Item') or not _E.other_direction_field(e[2], exp)]


def strip_other_direction(files):
    """Export permit = export fields only, import permit = import fields only: any filled Header / Item cell of the
    other direction (e.g. Inward MAWB, Importer, In HAWB on an export; Exporter, Consignee, CO ... on an import) is
    cleared, whatever wrote it (settings default, old template, typed value)."""
    import msd_engine as M
    import permit_engine as _E
    n = 0
    for f in files:
        exp = file_direction(f) == 'EXPORT'
        try:
            tpl = M.Template(f['path'])
        except Exception:
            continue
        changed = {}
        for name in ('Header', 'Item'):
            if name not in tpl.sheets:
                continue
            sh = tpl.sheet(name)
            cols = [c for h, c in tpl.headers(sh).items() if _E.other_direction_field(h, exp)]
            hit = False
            for r, row in sh.rows.items():
                if r < 2:
                    continue
                for c in cols:
                    if c in row[1] and '<f' not in (sh.get('%s%d' % (c, r)) or ''):
                        row[1].pop(c)
                        row[2] = hit = True
                        n += 1
            if hit:
                changed[tpl.sheets[name]] = sh.tostring()
        if changed:
            tpl.save(f['path'], changed)
    return n


def msg_type_for(f, form, S):
    """XML message type of one output file: job choice (Header tab) > Settings rule (team / importer / exporter name
    contains ...) > Settings default for import / export. Returns (type, how) - ('', '') = Settings default."""
    v = str(form.get('msg_type') or '').strip().upper()
    if v:
        return v, 'chosen on the job'
    team = str(form.get('team') or '').strip().upper()
    names = [str(s.get('party') or '') for s in f.get('ships') or [] if isinstance(s, dict)]
    for r in S.get('msg_rules') or []:
        if not isinstance(r, (list, tuple)) or len(r) < 2:
            continue
        k, mt = str(r[0] or '').strip().upper(), str(r[1] or '').strip().upper()
        if not k or not mt:
            continue
        if k == team or any(k in n.upper() for n in names):
            return mt, 'Settings rule %s' % k
    if form.get('decl_auto', True) is not False:                  # declaration advisor
        offered = [str(x).strip().upper() for x in (S.get('msg_types') or ['IPTDEC'])]
        ms = {(s.get('declaration') or {}).get('msg_type') or '' for s in f.get('ships') or [] if isinstance(s, dict)}
        ms.discard('')
        if len(ms) == 1:
            m = list(ms)[0]
            if m in offered:
                return m, 'declaration advisor'
    return '', ''


def _cnorm(n):
    n = re.sub(r'[^A-Z0-9]+', ' ', str(n or '').upper())
    n = re.sub(r'\b(PRIVATE|PTE|LIMITED|LTD|CO|COMPANY|SINGAPORE|SG)\b', ' ', n)
    return re.sub(r'\s+', ' ', n).strip()


def apply_declaration(files, form, log=None):
    """Declaration advisor decision -> Header of every output Excel (Declaration Type + code, Release / Receipt code
    when blank). Runs before the values typed on the sheet tabs, so a typed value always wins. Off: job decl_auto False."""
    if form.get('decl_auto', True) is False:
        return 0
    from openpyxl import load_workbook
    n = 0
    for f in files:
        decs = [s.get('declaration') or {} for s in f.get('ships') or [] if isinstance(s, dict)]
        decs = [d for d in decs if d.get('fill')]
        if not decs:
            continue
        d = decs[0]
        try:
            wb = load_workbook(f['path'], read_only=True, data_only=True)
            rows = list(wb['Header'].iter_rows(min_row=1, max_row=2, values_only=True))
            wb.close()
        except Exception:
            continue
        hv = {str(a).strip(): b for a, b in zip(rows[0], rows[1] if len(rows) > 1 else []) if a}
        ed = []
        for col, v in d['fill'].items():
            cur = str(hv.get(col) or '').strip()
            if col not in hv or cur == str(v):
                continue
            if col.startswith(('Release', 'Receipt')) and cur:
                continue                                  # location codes: only when blank
            ed.append(('Header', 2, col, v))
            if col.endswith('Location Code'):
                nm = LOCATIONS.get(str(v).upper(), '')
                if nm and not str(hv.get(col.replace('Code', 'Name')) or '').strip():
                    ed.append(('Header', 2, col.replace('Code', 'Name'), nm))
        if ed:
            n += _write_cells(f['path'], ed)
            if log:
                log('%s: declaration advisor applied %s (%s)%s - %s' % (
                    f['name'], d.get('applied'), d.get('applied_name'), ' changed from the %s' % d.get('source') if d.get('changed') else '',
                    ', '.join('%s %s' % (e[2], e[3]) for e in ed)))
    return n


def apply_carrier_book(files, S, log=None):
    """Carrier agent name copied by UEN from the carrier book (Settings > Carrier agents): UEN filled -> book name;
    only the name filled -> UEN from the book. Inward and outward, every output Excel."""
    book = {str(k).strip().upper(): str(v).strip().upper() for k, v in (S.get('carrier_book') or {}).items() if str(k).strip() and str(v).strip()}
    if not book:
        return 0
    rev = {}
    for u, nm in book.items():
        rev.setdefault(_cnorm(nm), u)
    from openpyxl import load_workbook
    n = 0
    for f in files:
        try:
            wb = load_workbook(f['path'], read_only=True, data_only=True)
            ws = wb['Header']
            rows = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
            wb.close()
        except Exception:
            continue
        if len(rows) < 2:
            continue
        hv = {str(a).strip(): b for a, b in zip(rows[0], rows[1]) if a}
        ed = []
        for side in ('Inward', 'Outward'):
            cn, cu = '%s Carrier Agent Name' % side, '%s Carrier Agent Uen' % side
            name = str(hv.get(cn) or '').strip().upper()
            uen = str(hv.get(cu) or '').strip().upper()
            if uen and uen in book and name != book[uen]:
                ed.append(('Header', 2, cn, book[uen]))
                if log:
                    log('%s: %s carrier agent %s = %s (carrier book%s)' % (f['name'], side.lower(), uen, book[uen],
                                                                         ', was "%s"' % name if name else ''))
            elif name and not uen and rev.get(_cnorm(name)):
                ed.append(('Header', 2, cu, rev[_cnorm(name)]))
                if log:
                    log('%s: %s carrier agent UEN %s for %s (carrier book)' % (f['name'], side.lower(), rev[_cnorm(name)], name))
        if ed:
            n += _write_cells(f['path'], ed)
    return n


def rezip_excels(files, out_dir):
    """EXCEL_ALL.zip again with the final Excel files (brand, typed values, direction clean-up)."""
    import zipfile
    zp = os.path.join(out_dir, 'EXCEL_ALL.zip')
    if not os.path.exists(zp):
        return
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            if os.path.exists(f['path']):
                z.write(f['path'], os.path.basename(f['path']))


PER_FILE_COLS = {"Trader's Remarks", 'Total Gross Weight'}


def apply_all_cells(st, files, log=None):
    n = 0
    for f in files:
        rm = file_rowmap(st, f)
        ed, orphan = file_edits(st, f['name'], rm)
        base = f.get('base') or ''
        if base and base != f['name'] and st.cells.get(base):   # values typed on the page (all items in one view)
            be, _ = file_edits(st, base, rm)                      # -> every split Excel (its own items, Header ...)
            split = len([x for x in files if x.get('base') == base]) > 1
            ed = [e for e in be if not (split and e[0] == 'Header' and e[2].strip() in PER_FILE_COLS)] + ed
        n += _write_cells(f['path'], same_direction(ed, file_direction(f)))
        if orphan and log:
            log('CHECK %s: %d typed item value(s) no longer match an item row (item removed / renamed) - not written'
                % (f['name'], orphan))
    return n


def _val(v):
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime('%d/%m/%Y')
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return int(v)
    return v


def read_sheets(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    out = collections.OrderedDict()
    for name in SHEET_TABS:
        if name not in wb.sheetnames or name == 'Items to be consolidated':
            continue
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        hdr = [str(h).strip() if h is not None else '' for h in rows[0]]
        n = len(hdr)
        while n and not hdr[n - 1]:
            n -= 1
        data = []
        for r in rows[1:]:
            vals = [_val(v) for v in (list(r) + [None] * n)[:n]]
            if any(v not in (None, '') for v in vals):
                data.append(vals)
            elif name in ('Header', 'Invoice'):
                break
        while data and not any(v not in (None, '') for v in data[-1]):
            data.pop()
        out[name] = dict(headers=hdr[:n], rows=data)
    wb.close()
    it = out.get('Item')
    if it:                                     # what the "Items to be consolidated" pivot shows in Excel
        H = {h: i for i, h in enumerate(it['headers'])}
        grp = collections.OrderedDict()
        for r in it['rows']:
            k = (r[H['Country of Origin']], r[H['HS Code']])
            g = grp.setdefault(k, dict(desc=r[H['Description']], qty=0.0, uom=r[H.get('HS UOM', 0)], amt=0.0, n=0))
            g['qty'] += float(r[H['HS Qty']] or 0)
            g['amt'] += float(r[H['Total Line Amount']] or 0)
            g['n'] += 1
        out['Items to be consolidated'] = dict(
            headers=['No.', 'Country of Origin', 'HS Code', 'Description', 'Qty/Unit', 'Line Amount', 'Item rows'],
            rows=[[i, k[0], k[1], (str(g['desc'] or '')[:120] + (' ...' if g['n'] > 1 else '')),
                   '%s %s' % (round(g['qty'], 4), g['uom'] or ''), round(g['amt'], 2), g['n']]
                  for i, (k, g) in enumerate(grp.items(), 1)], readonly=True)
    return out


def preview(tool, P, S, form, st, pdir, statefile):
    """Make the Excel files of this job into a preview folder (only when something changed) and return
    {files: [...], rowmaps: {file: {row: [ship, idx, key]}}}."""
    import hashlib
    key = hashlib.md5(json.dumps([os.path.getmtime(statefile) if os.path.exists(statefile) else 0,
                                  form, S], sort_keys=True, default=str).encode()).hexdigest()
    kf = os.path.join(pdir, 'preview.json')
    try:
        with open(kf, encoding='utf-8') as f:
            d = json.load(f)
        if d.get('key') == key and all(os.path.exists(x['path']) for x in d['files']):
            return d
    except Exception:
        pass
    shutil.rmtree(pdir, ignore_errors=True)
    os.makedirs(pdir)
    E = engine(tool)
    with st.lock:
        resolve_all(tool, P, S, form, st)
        _cfg_ships[0] = list(st.ships.values())
        cfg = build_cfg(tool, P, S, form, pdir, st.gi)
        cfg['report_name'] = '_preview_report.xlsx'
        cfg['no_split'] = True            # the page shows ALL items in one permit; the split (50 items ...) is made at Make Excel
        res = E.generate(list(st.ships.values()), cfg, log=lambda m: None)
        files, rowmaps = [], {}
        apply_brand(st, res[0], form)
        apply_declaration(res[0], form)
        for f in res[0]:
            rm = file_rowmap(st, f)
            ed, orphan = file_edits(st, f['name'], rm)
            _write_cells(f['path'], same_direction(ed, file_direction(f)))
            strip_other_direction([f])
            files.append(dict(name=f['name'], path=f['path'], ships=[s['key'] for s in f.get('ships') or []],
                              part=f.get('part'), orphan=orphan, direction=file_direction(f)))
            rowmaps[f['name']] = {str(k): v for k, v in rm.items()}
    d = dict(key=key, files=files, rowmaps=rowmaps)
    with open(kf, 'w', encoding='utf-8') as f:
        json.dump(d, f)
    return d


def edited_cells(st, fname, rowmap):
    """[[sheet, row, column]] of the typed values (to mark them on the page)."""
    rev = {(v[0], v[2]): int(r) for r, v in rowmap.items()}
    out = []
    for k in st.cells.get(fname) or {}:
        sheet, row, col = json.loads(k)
        if isinstance(row, list):
            row = rev.get((row[1], row[2]))
        if row is not None:
            out.append([sheet, int(row), col])
    return out


def list_cols(E, st, form):
    """Columns of the job list (as the permit system job list)."""
    act = [s for s in st.ships.values() if s.get('include', True)] or list(st.ships.values())
    cg = [s.get('cargo') or {} for s in act]
    hd = form.get('header') or {}
    pt = (form.get('permit_type') or next((c.get('permit_type') for c in cg if c.get('permit_type')), '') or '').upper()
    exp = any(s.get('direction') == 'EXPORT' for s in act)
    first = lambda k: next((c.get(k) for c in cg if c.get(k)), '')
    return dict(decl_type=getattr(E, 'PERMIT_SCHEME', {}).get(pt, '') or pt,
                party=', '.join(dict.fromkeys(s.get('party') or '' for s in act if s.get('party')))[:120],
                mawb=hd.get('mawb') or first('mawb'),
                port=(hd.get('discharge_port_code') or first('discharge_code')) if exp else (hd.get('loading_port_code') or first('loading_code')),
                eta=hd.get('date') or first('date'), direction='EXPORT' if exp else 'IMPORT' if act else '')



# ======================================================================== brand (every tool)
def brand_mode(form):
    """'' = brand from the documents / product master only, 'UNBRANDED' = items without a document brand unbranded,
    other text = that brand on the items without a document brand."""
    return (form.get('brand_mode') or '').strip().upper()


def apply_brand(st, files, form):
    """Brand / Unbranded columns of the Item sheet: a brand read from the documents (e.g. the maker printed under
    a Panasonic item) stays; the job's choice (Item tab: UNBRANDED / a brand) fills every item without one."""
    mode = brand_mode(form)
    n = nd = 0
    for f in files:
        its = f.get('items') or [it for s in f.get('ships') or [] for it in s['items']]
        edits = []
        for r, it in enumerate(its, 2):
            if it.get('brand'):
                edits += [('Item', r, 'Brand', it['brand']), ('Item', r, 'Unbranded', 'N')]
                nd += 1
            elif mode == 'UNBRANDED':
                edits += [('Item', r, 'Unbranded', 'Y'), ('Item', r, 'Brand', '')]
            elif mode:
                edits += [('Item', r, 'Brand', mode), ('Item', r, 'Unbranded', 'N')]
            else:
                continue
            n += 1
        if edits:
            _write_cells(f['path'], edits)
    return n, nd


# ======================================================================== port list (Loading / Discharge / Next / Last Port search)
PORTS_EXTRA = {   # UN/LOCODE -> name (well-known sea + air ports; more come from the uploaded list and from typed ports)
    'SGSIN': 'SINGAPORE', 'MYTPP': 'TANJUNG PELEPAS', 'MYPKG': 'PORT KLANG', 'MYKUL': 'KUALA LUMPUR', 'MYJHB': 'JOHOR BAHRU',
    'MYPGU': 'PASIR GUDANG', 'MYPEN': 'PENANG', 'IDBTM': 'BATAM', 'IDJKT': 'JAKARTA', 'IDSUB': 'SURABAYA', 'THLCH': 'LAEM CHABANG',
    'THBKK': 'BANGKOK', 'VNHPH': 'HAIPHONG', 'VNSGN': 'HO CHI MINH CITY', 'VNCMT': 'CAI MEP', 'VNHAN': 'HANOI', 'KHPNH': 'PHNOM PENH',
    'KHKOS': 'SIHANOUKVILLE', 'MMRGN': 'YANGON', 'PHMNL': 'MANILA', 'CNSHA': 'SHANGHAI', 'CNPVG': 'SHANGHAI PUDONG APT', 'CNNGB': 'NINGBO',
    'CNTAO': 'QINGDAO', 'CNTSN': 'TIANJIN', 'CNTXG': 'TIANJIN XINGANG', 'CNXMN': 'XIAMEN', 'CNDLC': 'DALIAN', 'CNSZX': 'SHENZHEN',
    'CNYTN': 'YANTIAN', 'CNSHK': 'SHEKOU', 'CNCAN': 'GUANGZHOU', 'CNPEK': 'BEIJING', 'HKHKG': 'HONG KONG', 'TWKHH': 'KAOHSIUNG',
    'TWTPE': 'TAIPEI', 'TWKEL': 'KEELUNG', 'TWTXG': 'TAICHUNG', 'KRPUS': 'BUSAN', 'KRINC': 'INCHEON', 'KRICN': 'INCHEON APT',
    'JPTYO': 'TOKYO', 'JPNRT': 'TOKYO NARITA APT', 'JPYOK': 'YOKOHAMA', 'JPUKB': 'KOBE', 'JPOSA': 'OSAKA', 'JPKIX': 'OSAKA KANSAI APT',
    'JPNGO': 'NAGOYA', 'JPHKT': 'HAKATA', 'JPMOJ': 'MOJI', 'INNSA': 'NHAVA SHEVA', 'INMUN': 'MUNDRA', 'INMAA': 'CHENNAI',
    'INBOM': 'MUMBAI', 'INCCU': 'KOLKATA', 'INDEL': 'DELHI', 'INBLR': 'BANGALORE', 'INHYD': 'HYDERABAD', 'INAMD': 'AHMEDABAD',
    'INCOK': 'COCHIN', 'INVTZ': 'VISAKHAPATNAM', 'LKCMB': 'COLOMBO', 'BDCGP': 'CHITTAGONG', 'PKKHI': 'KARACHI',
    'AEJEA': 'JEBEL ALI', 'AEDXB': 'DUBAI', 'AEAUH': 'ABU DHABI', 'OMSLL': 'SALALAH', 'SAJED': 'JEDDAH', 'SADMM': 'DAMMAM',
    'QADOH': 'DOHA', 'KWKWI': 'KUWAIT', 'EGPSD': 'PORT SAID', 'EGALY': 'ALEXANDRIA', 'TRIST': 'ISTANBUL', 'GRPIR': 'PIRAEUS',
    'ITGOA': 'GENOA', 'ITMIL': 'MILAN', 'ITMXP': 'MILAN MALPENSA APT', 'ESVLC': 'VALENCIA', 'ESALG': 'ALGECIRAS', 'ESBCN': 'BARCELONA',
    'ESMAD': 'MADRID', 'FRLEH': 'LE HAVRE', 'FRMRS': 'MARSEILLE', 'FRPAR': 'PARIS', 'FRCDG': 'PARIS CHARLES DE GAULLE APT',
    'BEANR': 'ANTWERP', 'BEBRU': 'BRUSSELS', 'NLRTM': 'ROTTERDAM', 'NLAMS': 'AMSTERDAM', 'DEHAM': 'HAMBURG', 'DEBRV': 'BREMERHAVEN',
    'DEFRA': 'FRANKFURT', 'DEMUC': 'MUNICH', 'GBFXT': 'FELIXSTOWE', 'GBSOU': 'SOUTHAMPTON', 'GBLON': 'LONDON',
    'GBLHR': 'LONDON HEATHROW APT', 'GBMAN': 'MANCHESTER', 'IEDUB': 'DUBLIN', 'SEGOT': 'GOTHENBURG', 'DKAAR': 'AARHUS',
    'DKCPH': 'COPENHAGEN', 'FIHEL': 'HELSINKI', 'NOOSL': 'OSLO', 'PLGDN': 'GDANSK', 'CHZRH': 'ZURICH', 'CHBSL': 'BASEL',
    'ATVIE': 'VIENNA', 'USLAX': 'LOS ANGELES', 'USLGB': 'LONG BEACH', 'USOAK': 'OAKLAND', 'USSEA': 'SEATTLE', 'USTIW': 'TACOMA',
    'USNYC': 'NEW YORK', 'USJFK': 'NEW YORK JFK APT', 'USSAV': 'SAVANNAH', 'USCHS': 'CHARLESTON', 'USHOU': 'HOUSTON',
    'USMIA': 'MIAMI', 'USORF': 'NORFOLK', 'USCHI': 'CHICAGO', 'USORD': 'CHICAGO OHARE APT', 'USATL': 'ATLANTA',
    'USSFO': 'SAN FRANCISCO', 'USDFW': 'DALLAS FORT WORTH', 'CAVAN': 'VANCOUVER', 'CAMTR': 'MONTREAL', 'CATOR': 'TORONTO',
    'MXZLO': 'MANZANILLO', 'BRSSZ': 'SANTOS', 'ARBUE': 'BUENOS AIRES', 'CLSAI': 'SAN ANTONIO', 'CLVAP': 'VALPARAISO',
    'PECLL': 'CALLAO', 'COCTG': 'CARTAGENA', 'ZADUR': 'DURBAN', 'ZACPT': 'CAPE TOWN', 'ZAJNB': 'JOHANNESBURG', 'KEMBA': 'MOMBASA',
    'NGLOS': 'LAGOS', 'AUSYD': 'SYDNEY', 'AUMEL': 'MELBOURNE', 'AUBNE': 'BRISBANE', 'AUFRE': 'FREMANTLE', 'AUADL': 'ADELAIDE',
    'AUPER': 'PERTH', 'NZAKL': 'AUCKLAND', 'NZTRG': 'TAURANGA', 'NZCHC': 'CHRISTCHURCH'}


def _ports_files(P):
    return os.path.join(P.masters, 'PORTS.xlsx'), os.path.join(P.masters, 'ports_learned.json')


def port_list(P):
    """code -> name: built in (engine tables + PORTS_EXTRA) < uploaded PORTS.xlsx < ports typed by the users (learned)."""
    E = engine('invoice')
    out = {}
    for name, code in getattr(E, 'SEA_PORTS', {}).items():
        out.setdefault(code, name)
    for code, name in getattr(E, 'AIRPORTS', {}).values():
        out.setdefault(code, name)
    out.update(PORTS_EXTRA)
    up, lj = _ports_files(P)
    if os.path.exists(up):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(up, read_only=True, data_only=True)
            for r in wb.worksheets[0].iter_rows(values_only=True):
                c, n = (str(r[0] or '').strip().upper().replace(' ', ''), str(r[1] or '').strip().upper()) if len(r) > 1 else ('', '')
                if re.fullmatch(r'[A-Z]{2}[A-Z0-9]{3}', c) and n:
                    out[c] = n
        except Exception:
            pass
    try:
        with open(lj, encoding='utf-8') as f:
            out.update(json.load(f))
    except Exception:
        pass
    return out


def port_learn(P, code, name):
    code, name = str(code or '').strip().upper().replace(' ', ''), str(name or '').strip().upper()
    if not re.fullmatch(r'[A-Z]{2}[A-Z0-9]{3}', code) or not name:
        return False
    _, lj = _ports_files(P)
    try:
        with open(lj, encoding='utf-8') as f:
            d = json.load(f)
    except Exception:
        d = {}
    if d.get(code) == name:
        return False
    d[code] = name
    with open(lj, 'w', encoding='utf-8') as f:
        json.dump(d, f, indent=1, sort_keys=True)
    return True


def port_search(P, q, limit=40):
    L = port_list(P)
    q = re.sub(r'\s+', ' ', str(q or '').strip().upper())
    if not q:
        return sorted(L.items())[:limit]
    qc = q.replace(' ', '')
    score = []
    for c, n in L.items():
        if c == qc or n == q:
            s = 0
        elif c.startswith(qc):
            s = 1
        elif n.startswith(q):
            s = 2
        elif q in n:
            s = 3
        elif len(qc) >= 2 and qc in c:
            s = 4
        else:
            continue
        score.append((s, n, c))
    score.sort()
    return [(c, n) for s, n, c in score[:limit]]


# ======================================================================== CASC product code by HS (RS Components master)
CASC_MASTER = 'RS_COMPONENTS_PRODUCT_CODE.xlsx'
_CASC_CACHE = {}


def read_any_rows(path):
    """Rows of the first sheet of an .xlsx / .xlsm / .xls file."""
    if path.lower().endswith('.xls'):
        import xlrd                                       # pip install xlrd (in requirements.txt)
        sh = xlrd.open_workbook(path).sheet_by_index(0)
        return [[(int(c.value) if c.ctype == 2 and float(c.value).is_integer() else (None if c.ctype == 0 else c.value))
                 for c in sh.row(i)] for i in range(sh.nrows)]
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    wb.close()
    return rows


def load_casc_master(path):
    """{8-digit HS: CASC Product Code 1 ('' = HS listed without a code)} from columns HS Code + CASC Product Code 1."""
    rows = read_any_rows(path)
    hi, ch, cc = None, None, None
    for i, r in enumerate(rows[:30]):
        names = [str(v or '').strip().lower() for v in r]
        if 'hs code' in names and any('casc' in n for n in names):
            hi, ch, cc = i, names.index('hs code'), next(k for k, n in enumerate(names) if 'casc' in n)
            break
    if hi is None:
        raise ValueError('No "HS Code" + "CASC Product Code 1" columns in %s' % os.path.basename(path))
    out = {}
    for r in rows[hi + 1:]:
        if ch >= len(r) or r[ch] in (None, ''):
            continue
        h = re.sub(r'\D', '', str(int(r[ch]) if isinstance(r[ch], float) else r[ch]))
        if not h:
            continue
        v = r[cc] if cc < len(r) else None
        out[h.zfill(8)[:8]] = str(v).strip().upper() if v not in (None, '') else ''
    return out


CUST_ITEM_COLS = ('item no', 'product id', 'gmid', 'material', 'part no', 'part number', 'product code', 'item code',
                  'material no', 'article no', 'sku')
CUST_HS_COLS = ('sg hs code', 'hs code', 'hs', 'tariff code', 'hs no')
CUST_CASC_COLS = ('ca/sc code', 'casc product code 1', 'casc', 'ca/sc', 'product code 1')
CUST_DESC_COLS = ('material description', 'description', 'item description', 'product description', 'short text')
CUST_CTL_COLS = ('control substances', 'controlled substances', 'control substance', 'chemical of concern', 'chemical in concern')
_CUST_CACHE = {}


def _cust_hs(h):
    """HS from a customer list: '040620' (6-digit text) -> 04062000, 4062000 (number lost its 0) -> 04062000."""
    if len(h) == 7 or len(h) == 5:
        h = '0' + h
    return (h + '00')[:8] if len(h) == 6 else h[:8] if len(h) >= 8 else ''


def load_cust_master(path):
    """Customer product list (e.g. Qnity / DuPont GMID master): every sheet with an item column (Product ID / GMID /
    Material / Item No) + HS Code (+ CA/SC Code). -> {squashed item: {hs, casc, desc, sheet, file}}"""
    import openpyxl
    out, sheets = {}, []
    sq = lambda v: re.sub(r'[^A-Z0-9]', '', str(v or '').upper())
    def cell(v):
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return str(v).replace('\xa0', ' ').strip() if v is not None else ''
    with open(path, 'rb') as fh:
        head = fh.read(8)
        body = fh.read(4 * 1024 * 1024) if head.startswith(b'\xd0\xcf\x11\xe0') else b''
    if b'E\x00n\x00c\x00r\x00y\x00p\x00t\x00e\x00d\x00P\x00a\x00c\x00k\x00a\x00g\x00e' in body or b'EncryptedPackage' in body:
        raise ValueError('This Excel file is ENCRYPTED (company sensitivity label / password) - the server cannot open it. '
                         'In Excel: select all cells of each sheet > Copy > a NEW blank workbook > Paste > Save As .xlsx '
                         '(or change the label to General / Public), then upload that copy.')
    if head.startswith(b'\xd0\xcf\x11\xe0') or path.lower().endswith('.xls'):     # old .xls (also when named .xlsx)
        try:
            import xlrd
        except ImportError:
            raise ValueError('Old Excel 97-2003 file - save it as .xlsx (Excel Workbook) and upload again '
                             '(or install xlrd on the server PC).')
        bk = xlrd.open_workbook(path)
        books = [(sh.name, [[c.value if c.ctype else None for c in sh.row(i)] for i in range(sh.nrows)]) for sh in bk.sheets()]
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        books = [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
        wb.close()
    for name, rows in books:
        hi = None
        for i, r in enumerate(rows[:30]):
            names = [re.sub(r'[\s._:#]+', ' ', cell(v).lower()).strip() for v in r]
            def col(exact, rx):            # exact header name first, else a header matching the pattern
                j = next((names.index(c) for c in exact if c in names), None)
                return j if j is not None else next((k for k, n in enumerate(names) if n and re.search(rx, n)), None)
            ci = col(CUST_ITEM_COLS + ('gmids',), r'^(material|gmids?|product id|item|part|article|sku)\b(?!.*(desc|text|group|type))')
            ch = col(CUST_HS_COLS, r'\bhs\b|tariff|hscode')
            if ci is not None and ch is not None:
                cc = col(CUST_CASC_COLS, r'ca ?/ ?sc|casc|product code 1|ca code')
                cd = col(CUST_DESC_COLS, r'desc|short text|material text|product name')
                ck = col(CUST_CTL_COLS, r'control+ed substance|control substance|chemical (?:of|in) concern')
                hi = i
                break
        if hi is None:
            continue
        n = 0
        for r in rows[hi + 1:]:
            g = lambda c: cell(r[c]) if c is not None and c < len(r) else ''
            k, h = sq(g(ci)), re.sub(r'\D', '', g(ch))
            if not k or not h:
                continue
            casc = g(cc).upper().strip()
            e = dict(hs=_cust_hs(h), casc='' if casc in ('', 'NA', 'N/A', '#N/A', 'TBA', 'NIL', 'NONE', '-')
                     else casc, desc=re.sub(r'\s+', ' ', g(cd)).strip(), sheet=name, file=os.path.basename(path), item_raw=g(ci),
                     control=re.sub(r'\s+', ' ', g(ck)).strip())
            out.setdefault(k, e)
            if k.isdigit() and k.lstrip('0') and k.lstrip('0') != k:
                out.setdefault(k.lstrip('0'), e)          # 00004018227 = 4018227
            n += 1
        sheets.append((name, n))
    if not out:
        raise ValueError('No sheet with an item column (Product ID / GMID / Material / Item No) and an HS Code column')
    return out, sheets


def cust_master_files(P):
    d = getattr(P, 'custdir', '')
    return sorted(f for f in os.listdir(d) if f.lower().endswith(('.xlsx', '.xlsm', '.xls')) and not f.startswith(('~', '_upload'))) \
        if d and os.path.isdir(d) else []


def cust_masters(P):
    """All customer product lists merged ({item: {...}}) + [(file, items, sheets)] - cached per file."""
    merged, info, per = {}, [], {}
    for f in cust_master_files(P):
        path = os.path.join(P.custdir, f)
        k = (path, os.path.getmtime(path))
        if k not in _CUST_CACHE:
            try:
                _CUST_CACHE[k] = load_cust_master(path)
            except Exception as e:
                _CUST_CACHE[k] = ({}, [('error: %s' % e, 0)])
        m, sheets = _CUST_CACHE[k]
        for a, b in m.items():
            merged.setdefault(a, b)
        per[f] = m
        info.append((f, len(m), sheets))
    _CUST_PER['last'] = per
    return merged, info


_CUST_PER = {}


def cust_lists(P):
    """{file: {item: row}} of the customer product lists (vendor choice on the job)."""
    cust_masters(P)
    return _CUST_PER.get('last') or {}


def casc_by_hs(P):
    path = next((os.path.join(P.cascdir, f) for f in sorted(os.listdir(P.cascdir))
                 if f.lower().endswith(('.xlsx', '.xls', '.xlsm')) and not f.startswith('~')), '') if os.path.isdir(getattr(P, 'cascdir', '')) else ''
    if not path:
        return {}
    k = (path, os.path.getmtime(path))
    if k not in _CASC_CACHE:
        try:
            _CASC_CACHE.clear()
            _CASC_CACHE[k] = load_casc_master(path)
        except Exception:
            return {}
    return _CASC_CACHE[k]


def analysis(st):
    """Document extraction analysis of a job (doc_analysis.analyse) for the web page."""
    import doc_analysis
    return doc_analysis.analyse(st.docs, st.ships)
