# -*- coding: utf-8 -*-
"""Festo documents: invoice (SPOOL, one position per page), ZGPL packing list, sea waybill / B/L,
DHL "Customs Instructions Import" (CI).  Every reader returns the common document shape of danfoss_engine."""
import os, re, collections
import danfoss_engine as E

NUM = r'\d[\d,]*(?:\.\d+)?'
MON = dict(JAN=1, FEB=2, MAR=3, APR=4, MAY=5, JUN=6, JUL=7, AUG=8, SEP=9, OCT=10, NOV=11, DEC=12)
SEA_PORTS = {  # city -> UN/LOCODE
    'SHANGHAI': 'CNSHA', 'NINGBO': 'CNNGB', 'SHENZHEN': 'CNSZX', 'YANTIAN': 'CNYTN', 'SHEKOU': 'CNSHK',
    'QINGDAO': 'CNTAO', 'XIAMEN': 'CNXMN', 'TIANJIN': 'CNTXG', 'XINGANG': 'CNTXG', 'DALIAN': 'CNDLC', 'HONG KONG': 'HKHKG',
    'BUSAN': 'KRPUS', 'KAOHSIUNG': 'TWKHH', 'KEELUNG': 'TWKEL', 'TOKYO': 'JPTYO', 'YOKOHAMA': 'JPYOK', 'KOBE': 'JPUKB',
    'NAGOYA': 'JPNGO', 'OSAKA': 'JPOSA', 'HAMBURG': 'DEHAM', 'BREMERHAVEN': 'DEBRV', 'ROTTERDAM': 'NLRTM',
    'ANTWERP': 'BEANR', 'LE HAVRE': 'FRLEH', 'FELIXSTOWE': 'GBFXT', 'GENOA': 'ITGOA', 'LA SPEZIA': 'ITSPE',
    'BARCELONA': 'ESBCN', 'VALENCIA': 'ESVLC', 'PORT KLANG': 'MYPKG', 'PASIR GUDANG': 'MYPGU', 'PENANG': 'MYPEN',
    'LAEM CHABANG': 'THLCH', 'BANGKOK': 'THBKK', 'JAKARTA': 'IDJKT', 'SURABAYA': 'IDSUB', 'HO CHI MINH': 'VNSGN',
    'HAIPHONG': 'VNHPH', 'MANILA': 'PHMNL', 'CHENNAI': 'INMAA', 'NHAVA SHEVA': 'INNSA', 'MUNDRA': 'INMUN',
    'NEW YORK': 'USNYC', 'LOS ANGELES': 'USLAX', 'LONG BEACH': 'USLGB', 'SAVANNAH': 'USSAV', 'SYDNEY': 'AUSYD',
    'MELBOURNE': 'AUMEL', 'SINGAPORE': 'SGSIN', 'JEBEL ALI': 'AEJEA', 'DUBAI': 'AEJEA',
}
CODE_NAME = {v: k for k, v in SEA_PORTS.items()}
CODE_NAME.update(CNSHA='SHANGHAI', SGSIN='SINGAPORE', CNTXG='TIANJIN')


def _d(s):
    """24-SEP-26 / 07-Oct-26 / 2026-09-08 / 08/09/2026 -> dd/mm/yyyy"""
    s = (s or '').strip()
    m = re.match(r'(\d{1,2})[-\s]([A-Za-z]{3})[A-Za-z]*[-\s,]*(\d{2,4})$', s)
    if m and m.group(2).upper() in MON:
        y = int(m.group(3)) + (2000 if len(m.group(3)) == 2 else 0)
        return '%02d/%02d/%d' % (int(m.group(1)), MON[m.group(2).upper()], y)
    m = re.match(r'(\d{4})-(\d{2})-(\d{2})$', s)
    if m:
        return '%s/%s/%s' % (m.group(3), m.group(2), m.group(1))
    m = re.match(r'(\d{2})[./](\d{2})[./](\d{4})$', s)
    return '%s/%s/%s' % m.groups() if m else ''


def port_code(text):
    t = (text or '').upper()
    m = re.search(r'\b([A-Z]{2}[A-Z0-9]{3})\b', t)
    if m and m.group(1) in CODE_NAME:
        return m.group(1), CODE_NAME[m.group(1)]
    for city in sorted(SEA_PORTS, key=len, reverse=True):
        if city in t:
            return SEA_PORTS[city], city
    return '', ''


# ======================================================================== classify
def classify(text):
    t = text or ''
    if re.search(r'Kundennummer\s+Belegnummer', t) or re.search(r'HS=\d{6,10},\s*CoO=[A-Z]{2}', t):
        return 'FINV2'                     # Festo invoice, German / EU number layout (e.g. CSC Rohrbach "ICAPO CIV")
    if re.search(r'Customs tariff number\s*=', t) or (re.search(r'Position\s+\d{4}\s*/\s*Your Position', t)) or \
            (re.search(r'^Invoice\s*$', t[:400], re.M) and re.search(r'Customer No\.\s*Document no\.', t)):
        return 'FINV'
    if re.search(r'Packing List', t) and re.search(r'Inpack', t, re.I):
        return 'FPL'
    if re.search(r'Customs Instructions (Import|Export)', t, re.I):
        return 'CI'
    if re.search(r'(SEA\s*WAYBILL|BILL\s+OF\s+LADING|B/L\s*No\.?|CONTINUATION PAGE BILL OF LADING)', t, re.I) and \
            re.search(r'Port of (loading|discharge)|Vessel', t, re.I):
        return 'BL'
    return ''


# ======================================================================== invoice
MAIN = re.compile(r'^(.*?)\s+(\d{3,})\s+(' + NUM + r')\s+(' + NUM + r')\s+(\d[\d,]*\.\d{2})\s*$')
TYPE = re.compile(r'^(\S.*?)\s+([A-Z]{1,3})\s*$')


def parse_invoice(path, pages, countries=None):
    C = countries or E.Countries()
    text = '\n'.join(p['text'] for p in pages)
    d = dict(kind='INV', supplier='FESTO', file=os.path.basename(path), path=path, notes=[], lines=[], title='INVOICE',
             reader='Festo invoice', ocr=any(p.get('ocr') for p in pages))
    m = re.search(r'^Invoice\s*\n\s*(\d{6,10})\b', text, re.M) or re.search(r'Document no\.\s*Currency.*\n\s*\d+\s+(\d{6,10})', text)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Customer No\.\s*Document no\.\s*Currency.*\n\s*(\d+)\s+(\d+)\s+([A-Z]{3})', text)
    d['customer_no'] = m.group(1) if m else ''
    d['currency'] = m.group(3) if m else ''
    m = re.search(r'^Date\s*\n\s*(\d{2}/\d{2}/\d{4})', text, re.M)
    d['invoice_date'] = m.group(1) if m else ''
    m = re.search(r'Incoterms:\s*([A-Z]{3})\s*(?:/\s*([^\n]+?))?\s*(?:\d+\s+of\s+\d+)?\s*$', text, re.M)
    d['incoterm'] = m.group(1) if m else ''
    d['incoterm_place'] = E.clean(m.group(2) or '') if m else ''
    m = re.search(r'Payment terms:\s*(.+?)\s*(?:Page)?\s*$', text, re.M)
    d['payment_terms'] = m.group(1) if m else ''
    m = re.search(r'Transport\s+(\d{6,})\s*/\s*(\d{2}/\d{2}/\d{4})', text)
    d['transport'] = m.group(1) if m else ''
    d['ship_date'] = m.group(2) if m else ''
    m = re.search(r'Gross weight:\s*(' + NUM + r')\s*(KG|G|T)\b', text)
    d['total_gw'] = E.fnum(m.group(1)) if m else None
    m = re.search(r'Net weight:\s*(' + NUM + r')\s*(KG|G|T)\b', text)
    d['total_nw'] = E.fnum(m.group(1)) if m else None
    d['packages_ids'] = list(dict.fromkeys(re.findall(r'^Pack\.\s*-.*?\s(\d{8,})\s*$', text, re.M)))
    d['boxes'] = len(d['packages_ids']) or None
    m = re.search(r'Packing Charges\s+([A-Z]{3})\s+(' + NUM + ')', text)
    d['charges'] = E.fnum(m.group(2)) if m else 0.0
    d['charges_label'] = 'Packing Charges' if m else ''
    m = re.search(r'Freight(?: Charges)?\s+([A-Z]{3})\s+(' + NUM + ')', text)
    d['freight'] = E.fnum(m.group(2)) if m else 0.0
    m = re.search(r'Total Amount\s+([A-Z]{3})\s+(' + NUM + ')', text) or re.search(r'Total Net\s+([A-Z]{3})\s+(' + NUM + ')', text)
    d['total'] = E.fnum(m.group(2)) if m else None
    if m and not d['currency']:
        d['currency'] = m.group(1)
    cars = collections.Counter(E.clean(re.sub(r'\s+DE \d{3} \d{3} \d{3}.*$', '', x)).upper()
                               for x in re.findall(r'^Carrier:[^\n]*\n(?:Umsatz[^\n]*\n)?\s*([^\n]+)', text, re.M))
    cars = [c for c, _ in cars.most_common() if not re.search(r'UMSATZ|NUMMER|HRA|HRB', c)]
    d['carrier'] = cars[0] if cars else ''
    m = re.search(r'shipping type:.*\n?.*?\b(Seafreight|Airfreight|Air freight|Sea freight|Truck|Courier|Road)\b', text, re.I)
    d['ship_mode'] = ('Sea' if 'SEA' in m.group(1).upper() else 'Air' if 'AIR' in m.group(1).upper() else m.group(1)) if m else ''
    # parties
    bt = re.search(r'Bill-to\s*\n(.+?)\n(.+?)\n(.+?)\n', text)
    d['invoice_to_block'] = [E.clean(x) for x in bt.groups()] if bt else []
    d['invoice_to'] = d['invoice_to_block'][0].upper() if bt else ''
    d['party_name'] = re.sub(r'\s*-\s*SG\b.*$', '', d['invoice_to']).strip()     # Festo Pte Ltd-SG -> FESTO PTE LTD
    st = re.search(r'Ship-to\s*\n(?:.*\n){0,3}?\s*(\d{4,8})\s*(?:[^\n]*)\n(?:[^\n]*\n){0,2}?\s*(Festo[^\n]*?|[A-Z][^\n]*?(?:Ltd|LTD|Pte|PTE)[^\n]*?)(?:\s{2,}|\s+Registergericht|\s*$)', text, re.M)
    d['ship_to'] = E.clean(st.group(2)).upper() if st else d['invoice_to']
    d['ship_to_block'] = [d['ship_to']]
    d['invoice_from'] = d['ship_from'] = 'FESTO SE & CO. KG' if 'Festo SE & Co. KG' in text else ''
    if re.search(r'Hub Asia', text, re.I):
        d['ship_from'] = 'FESTO HUB ASIA (SHANGHAI)'
    d['ship_to_country'] = 'SG' if re.search(r'SINGAPORE|SINGAPUR', text[:1500]) else ''
    d['ship_from_country'] = 'CN' if re.search(r'Shanghai|Hub Asia', text, re.I) else ''
    d['direction'] = 'IMPORT' if d['ship_to_country'] == 'SG' else 'EXPORT'
    d['hawb'] = ''
    d['uen'] = ''
    d['tax_ids'] = []
    # ---- positions
    cur = None
    order = delivery = ''
    for p in pages:
        L = p['text'].splitlines()
        i = 0
        while i < len(L):
            ln = L[i].strip()
            mo = re.match(r'^(\d{8,12})\s+from\s+\d{2}/\d{2}/\d{4}', ln)
            if mo:
                order = mo.group(1)
            if re.match(r'Our delivery note no\.:', ln) and i + 1 < len(L):
                md = re.match(r'\s*(\d{6,})', L[i + 1])
                delivery = md.group(1) if md else delivery
            mp = re.match(r'^Position\s+(\d{3,5})\s*/\s*Your Position\s*(\d+)', ln)
            if mp:
                j = i + 1
                while j < len(L) and not MAIN.match(L[j].strip()) and j < i + 4:
                    j += 1
                mm = MAIN.match(L[j].strip()) if j < len(L) else None
                if not mm:
                    d['notes'].append('Position %s: price line not readable' % mp.group(1))
                    i += 1
                    continue
                cur = dict(line=mp.group(1), your_pos=mp.group(2), desc=E.clean(mm.group(1)), item=mm.group(2),
                           price=E.fnum(mm.group(3)), qty=E.fnum(mm.group(4)), amount=E.fnum(mm.group(5)), uom='',
                           type_code='', hs='', coo='', nw=None, gw=None, currency=d['currency'], order=order,
                           delivery=delivery, eccn='', license='', pl='', lots=[], boxes=[], cust_item='')
                mt = TYPE.match(L[j + 1].strip()) if j + 1 < len(L) else None
                if mt and not re.match(r'(Your material|Customs tariff|Country of origin|Weight)', L[j + 1].strip()):
                    cur['type_code'], cur['uom'] = mt.group(1), mt.group(2)
                d['lines'].append(cur)
                i = j + 2
                continue
            if cur is not None:
                m = re.search(r'Customs tariff number\s*=\s*(\d{6,12})', ln)
                if m:
                    cur['hs'] = m.group(1)
                m = re.search(r'Country of origin\s*=\s*([A-Z]{2})\b', ln)
                if m:
                    cur['coo'] = m.group(1)
                m = re.search(r'^Weight:\s*(' + NUM + r')\s*(KG|G|T)\b', ln)
                if m:
                    w = E.fnum(m.group(1))
                    cur['nw'] = w / 1000.0 if m.group(2) == 'G' else w * 1000 if m.group(2) == 'T' else w
                m = re.match(r'^(0{6,}\d+)\s*(?:$|\S)', ln)
                if m and not cur['cust_item']:
                    cur['cust_item'] = m.group(1).lstrip('0')
                m = re.search(r'US-EAR:\s*(EAR99|\d[A-E]\d{3}[A-Za-z0-9.]*)', ln)
                if m:
                    cur['eccn'] = m.group(1).rstrip(',.')
            i += 1
    if not d['lines']:
        d['notes'].append('no invoice positions could be read')
    d['line_nw'] = round(sum(l['nw'] or 0 for l in d['lines']), 3)
    return d


# ======================================================================== invoice - second layout (EU numbers)
EUN = r'\d[\d.]*(?:,\d+)?'
MAIN2 = re.compile(r'^(.*?)\s+(\d{3,})\s+(' + EUN + r')\s+(?:(-?\d+(?:,\d+)?\s*%)\s+)?(' + EUN + r')\s+(\d[\d.]*,\d{2})(?:\s+.*)?$')


def eu(v):
    """1.234,56 -> 1234.56 (German numbers)"""
    v = (v or '').strip().replace(' ', '')
    if not v:
        return None
    try:
        return float(v.replace('.', '').replace(',', '.'))
    except ValueError:
        return None


MONTHS_EN = dict(JANUARY=1, FEBRUARY=2, MARCH=3, APRIL=4, MAY=5, JUNE=6, JULY=7, AUGUST=8, SEPTEMBER=9, OCTOBER=10,
                 NOVEMBER=11, DECEMBER=12)


def parse_invoice2(path, pages, countries=None):
    """Festo invoice, second layout: "Kundennummer Belegnummer ... Invoice", "Position 000010", main line
    "self-tap. screw 727419 0,06 200 12,00", type code line, "HS=73181499, CoO=DE, Weight 160,0 G",
    "Value of goods", "Carriage", "Packing costs", "Invoice Amount (EUR)"."""
    text = '\n'.join(p['text'] for p in pages)
    d = dict(kind='INV', supplier='FESTO', file=os.path.basename(path), path=path, notes=[], lines=[], title='INVOICE',
             reader='Festo invoice (EU layout)', ocr=any(p.get('ocr') for p in pages))
    m = re.search(r'Kundennummer\s+Belegnummer.*\n\s*(\d+)\s+(\d{6,10})', text) or re.search(r'^Invoice\s*\n\s*(\d+)?\s*(\d{6,10})', text, re.M)
    d['customer_no'], d['invoice_no'] = (m.group(1) or '', m.group(2)) if m else ('', '')
    m = re.search(r'^Date\s*\n\s*(\d{1,2})\.?\s+([A-Za-z]+)\s+(\d{4})', text, re.M)
    d['invoice_date'] = ('%02d/%02d/%s' % (int(m.group(1)), MONTHS_EN.get(m.group(2).upper(), 0), m.group(3))) if m and \
        m.group(2).upper() in MONTHS_EN else ''
    m = re.search(r'Price\s+([A-Z]{3})\b', text) or re.search(r'Invoice Amount \(([A-Z]{3})\)', text)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'delivery terms\s+([A-Z]{3})\s+([A-Za-z][\w .-]*?)(?:\s{2,}|\s+Rechtsform|\s+\(BLZ|\s*$)', text, re.M)
    d['incoterm'] = m.group(1) if m else ''
    d['incoterm_place'] = E.clean(m.group(2)) if m else ''
    m = re.search(r'Transport No\.?\s*(\d{6,})\s+dated\s+(\d{2})\.(\d{2})\.(\d{4})(?:\s+from\s+([^\n]+))?', text)
    d['transport'] = m.group(1) if m else ''
    d['ship_date'] = '%s/%s/%s' % m.group(2, 3, 4) if m else ''
    d['ship_from'] = ('FESTO ' + E.clean(m.group(5)).upper()) if m and m.group(5) else 'FESTO SE & CO. KG'
    m = re.search(r'grossweight\s+(' + EUN + r')\s*KG\s*/\s*netweight\s+(' + EUN + r')\s*KG', text, re.I)
    d['total_gw'], d['total_nw'] = (eu(m.group(1)), eu(m.group(2))) if m else (None, None)
    d['packages_ids'] = list(dict.fromkeys(re.findall(r'^(?:VERP\.?-?\s*KARTON|Pack\.?\s*-|KARTON|PALETTE|Pallet)[^\n]*?\s(\d{8,})\s*$',
                                                      text, re.M | re.I)))
    d['boxes'] = len(d['packages_ids']) or None
    m = re.search(r'Packing costs:\s*(' + EUN + ')', text, re.I)
    d['charges'] = eu(m.group(1)) or 0.0 if m else 0.0
    d['charges_label'] = 'Packing costs' if d['charges'] else ''
    m = re.search(r'Carriage:\s*(' + EUN + ')', text, re.I)
    d['freight'] = eu(m.group(1)) or 0.0 if m else 0.0
    m = re.search(r'Invoice Amount \(([A-Z]{3})\)\s+(' + EUN + ')', text) or re.search(r'Invoice Amount net\s+(' + EUN + ')', text)
    d['total'] = eu(m.group(2) if m and m.lastindex == 2 else m.group(1)) if m else None
    m = re.search(r'Value of goods\s+(' + EUN + ')', text)
    d['goods_value'] = eu(m.group(1)) if m else None
    m = re.search(r'Delivery by\s+(Seafreight|Airfreight|Sea freight|Air freight|Truck|Courier|Road)', text, re.I)
    d['ship_mode'] = ('Sea' if 'SEA' in m.group(1).upper() else 'Air' if 'AIR' in m.group(1).upper() else m.group(1)) if m else ''
    d['carrier'] = ''
    top = pages[0]['text'].splitlines()
    d['invoice_to_block'] = [E.clean(x) for x in top[:4]]
    d['invoice_to'] = d['invoice_to_block'][0].upper() if top else ''
    d['party_name'] = re.sub(r'\s*-\s*SG\b.*$', '', d['invoice_to']).strip()
    d['ship_to'] = d['invoice_to']
    d['ship_to_block'] = [d['ship_to']]
    d['invoice_from'] = 'FESTO SE & CO. KG'
    d['ship_to_country'] = 'SG' if re.search(r'SINGAPORE|SINGAPUR', '\n'.join(top[:5])) else ''
    d['ship_from_country'] = 'DE' if re.search(r'Rohrbach|Esslingen|St\. Ingbert', text) else ''
    d['direction'] = 'IMPORT' if d['ship_to_country'] == 'SG' else 'EXPORT'
    d['hawb'] = d['uen'] = ''
    d['tax_ids'] = []
    cur = None
    order = delivery = ''
    for p in pages:
        L = [x.strip() for x in p['text'].splitlines()]
        for i, ln in enumerate(L):
            mo = re.match(r'^Your order\s+(\d{6,})', ln)
            if mo:
                order = mo.group(1)
            if re.match(r'^Order\s+\d+/\d+\s+delivery note', ln):
                md = next((re.match(r'(\d{8,})\s*$', x) for x in L[i + 1:i + 4] if re.match(r'(\d{8,})\s*$', x)), None)
                delivery = md.group(1) if md else delivery
            mp = re.match(r'^Position\s+(\d{4,6})\b', ln)
            if mp:
                mm = None
                for j in range(i + 1, min(i + 4, len(L))):
                    mm = MAIN2.match(L[j])
                    if mm:
                        break
                if not mm:
                    d['notes'].append('Position %s: price line not readable' % mp.group(1))
                    continue
                qty = eu(mm.group(5))
                cur = dict(line=mp.group(1).lstrip('0').zfill(4), desc=E.clean(mm.group(1)), item=mm.group(2),
                           price=eu(mm.group(3)), discount=(mm.group(4) or '').replace(' ', ''), qty=qty,
                           amount=eu(mm.group(6)),
                           uom='M' if re.search(r'tubing|hose|cable|tube', mm.group(1), re.I) else 'PC', type_code='', hs='', coo='', nw=None, gw=None,
                           currency=d['currency'], order=order, delivery=delivery, eccn='', license='', pl='', lots=[],
                           boxes=[], cust_item='')
                if j + 1 < len(L) and not re.match(r'(Your Part|HS=|Persönlich|Orderposition)', L[j + 1]):
                    cur['type_code'] = re.split(r'\s{2,}|\s+(?:Persönlich|Gesellschafterin|Rechtsform)', L[j + 1])[0].strip()
                d['lines'].append(cur)
                continue
            if cur is None:
                continue
            m = re.search(r'HS=(\d{6,10}),\s*CoO=([A-Z]{2}),\s*Weight\s+(' + EUN + r')\s*(KG|G|T)\b', ln)
            if m:
                cur['hs'], cur['coo'] = m.group(1), m.group(2)
                w = eu(m.group(3))
                cur['nw'] = w / 1000.0 if m.group(4) == 'G' else w * 1000 if m.group(4) == 'T' else w
            m = re.match(r'^(0{6,}\d+)\b', ln)
            if m and not cur['cust_item']:
                cur['cust_item'] = m.group(1).lstrip('0')
            m = re.search(r'US-EAR:\s*(EAR99|\d[A-E]\d{3}[A-Za-z0-9.]*)', ln)
            if m:
                cur['eccn'] = m.group(1)
    if not d['lines']:
        d['notes'].append('no invoice positions could be read')
    if d.get('goods_value') is not None:
        ls = round(sum(l['amount'] or 0 for l in d['lines']), 2)
        if abs(ls - d['goods_value']) > 0.01:
            d['notes'].append('Value of goods %s but positions add up to %s' % (d['goods_value'], ls))
    d['line_nw'] = round(sum(l['nw'] or 0 for l in d['lines']), 3)
    return d


# ======================================================================== packing list (ZGPL)
PL_LINE = re.compile(r'^(\d{8,12})\s+(\d{3,5})/(\d+)\s+(\S+)\s+([A-Z]{2})\s+(.+?)\s+(\d[\d,]*\.\d{3})\s+(' + NUM + r')\s+([A-Z]{1,3})\s*$')
PL_PKG = re.compile(r'^(\d{6,})\s+(\d{8,})\s+(' + NUM + r')\s*KG\s+(' + NUM + r')\s*KG\s*$')


def parse_pl(path, pages):
    text = '\n'.join(p['text'] for p in pages)
    d = dict(kind='PL', supplier='FESTO', file=os.path.basename(path), path=path, reader='Festo packing list', notes=[],
             invoice_refs=[], hawb='', mawb='', total_gw=None, total_nw=None, packages=None, package_unit='PKG',
             containers=[], lines=[], ocr=any(p.get('ocr') for p in pages), deliveries=[], transport='', pkgs={})
    cur_pkg = ''
    for raw in text.splitlines():
        ln = raw.strip()
        m = PL_PKG.match(ln)
        if m:
            d['transport'] = d['transport'] or m.group(1)
            cur_pkg = m.group(2)
            d['pkgs'][cur_pkg] = (E.fnum(m.group(3)), E.fnum(m.group(4)))
            continue
        m = PL_LINE.match(ln)
        if m:
            d['lines'].append(dict(line=m.group(2), delivery=m.group(1), item=m.group(4), coo=m.group(5),
                                   desc=E.clean(m.group(6)), nw=E.fnum(m.group(7)), gw=None, qty=E.fnum(m.group(8)),
                                   uom=m.group(9), hs='', boxes=[cur_pkg], lots=[], amount=None, price=None,
                                   currency='', eccn='', order='', license='', pl=''))
            continue
        m = re.match(r'^HS-Code:\s*(\d{6,12})', ln)
        if m and d['lines']:
            d['lines'][-1]['hs'] = m.group(1)
    if not d['transport']:
        m = re.search(r'Shipment\s*No\.?.*\n\s*(\d{6,})', text)
        d['transport'] = m.group(1) if m else ''
    d['packages'] = len(d['pkgs']) or None
    d['total_gw'] = round(sum(g for g, n in d['pkgs'].values()), 3) or None
    d['total_nw'] = round(sum(n for g, n in d['pkgs'].values()), 3) or None
    d['deliveries'] = list(dict.fromkeys(l['delivery'] for l in d['lines']))
    if not d['lines']:
        d['notes'].append('packing list rows could not be read')
    return d


# ======================================================================== sea waybill / B/L and customs instructions
def _cargo_doc(path, pages, kind_label):
    return dict(kind='AWB', doc_type=kind_label, file=os.path.basename(path), path=path, notes=[], hawb='', mawb='',
                pieces=None, gw=None, departure='', legs=[], flights=[], consignee='', shipper='', cargo={},
                invoice_refs=[], marks=[], issued_by='', ocr=any(p.get('ocr') for p in pages))


def parse_bl(path, pages):
    text = '\n'.join(p['text'] for p in pages)
    U = text.upper()
    d = _cargo_doc(path, pages, 'B/L')
    cg = dict(mode='Sea')
    m = re.search(r'B/?L\.?\s*No\.?\s*\n.*?\b([A-Z]{4}\d{5,})\s*$', text, re.M) or \
        re.search(r'BILL OF LADING\s*-\s*([A-Z0-9]{6,})', text) or re.search(r'B/?L\s*No\.?\s*:?\s*([A-Z]{3,4}[A-Z0-9]{5,})', text)
    d['hawb'] = m.group(1) if m else ''
    m = re.search(r'Vessel\s+Voyage No\.?.*\n(?:.*\n){0,4}?\s*([A-Z][A-Z0-9 .\-]+?)\s+([0-9]{2,5}[A-Z]?)\s+(?:upon|Such|$)', text, re.M)
    if m:
        cg['transport_id'], cg['flight'] = E.clean(m.group(1)), m.group(2)
    m = re.search(r'Place of receipt\s+Port of loading.*\n(?:.*\n){0,2}?\s*[A-Z][A-Z ]+,\s*[A-Z ]+?\s+([A-Z][A-Z ]+,\s*[A-Z ]+?)\s*$', text, re.M) \
        or re.search(r'^Port of loading\s*:?\s*\n?\s*(.+)$', text, re.M | re.I)
    if m:
        c, n = port_code(m.group(1))
        if c:
            cg['loading_code'], cg['loading_name'] = c, n
    m = re.search(r'^Port of discharge.*\n\s*(.+)', text, re.M)
    if m:
        c, n = port_code(m.group(1))
        if c:
            cg['discharge_code'], cg['discharge_name'] = c, n
    m = re.search(r'SHIPPED ON BOARD\s+(\d{1,2}-[A-Z]{3}-\d{2,4})', U)
    if m:
        cg['etd'] = _d(m.group(1))
    m = re.search(r'Total No\.? of containers/packages:\s*(\d+)', text, re.I)
    conts = []
    for mc in re.finditer(r'\b([A-Z]{4}\d{7})\s+([A-Z0-9]{5,})\s+(\d{2}[A-Z0-9]{2})\s+(' + NUM + r')\s*KG\s+(' + NUM + r')\s*M3\s+(\d+)\s*([A-Z]{2,4})', text):
        conts.append(dict(no=mc.group(1), seal=mc.group(2), size=mc.group(3)[:2], size_type=mc.group(3),
                          tne=E.fnum(mc.group(4)) / 1000.0, packages=int(mc.group(6))))
    if not conts:
        conts = [dict(no=c, seal='', size='', tne=None) for c in dict.fromkeys(re.findall(r'\b([A-Z]{4}\d{7})\b', U))]
    cg['containers'] = [c['no'] for c in conts]
    cg['containers_info'] = conts
    m = re.search(r'(' + NUM + r')\s*KG\s+(' + NUM + r')\s*M3', text)
    if m:
        cg['gw_kg'] = E.fnum(m.group(1))
    m = re.search(r'PACKAGE QUANTITY\s*\n\s*(\d+)\s*([A-Z]{2,4})', text) or re.search(r'\b(\d+)\s+Package\(s\)', text)
    if m:
        cg['packages_qty'] = int(m.group(1))
        cg['packages_unit'] = 'PKG'
    m = re.search(r'INVOICE\s*(?:STC.*?)?\n?\s*NO\.?:?\s*([\d, ]{6,})', text)
    d['invoice_refs'] = re.findall(r'\d{6,}', m.group(1)) if m else []
    m = re.search(r'Marks and Nos\..*\n(?:.*\n)?\s*(\d{6,})\s', text)
    d['marks'] = [m.group(1)] if m else []
    m = re.search(r'Shipper\s+Document No\..*\n\s*(.+?)\s+S\d{6,}', text)
    d['shipper'] = E.clean(m.group(1)).upper() if m else ''
    m = re.search(r'Consignee .*\n(?:.*\n)?\s*(?:[A-Z][^\n]*\n)?.*?(FESTO PTE LTD[^\n]*?|[A-Z][A-Z &().-]+PTE\.? LTD\.?[^\n]*?)(?:\s{2,}|\s+\d|$)', text, re.M)
    d['consignee'] = E.clean(m.group(1)) if m else ''
    m = re.search(r'For the release of goods apply to:.*\n(?:.*\n)?\s*([A-Z][A-Z &().,-]+(?:LTD|LIMITED)\.?)', text)
    d['release_agent'] = E.clean(m.group(1)) if m else ''
    m = re.search(r'DHL GLOBAL FORWARDING \([A-Z]+\) CO\., LTD\.|DHL GLOBAL FORWARDING[^\n]*?LTD', text)
    d['issued_by'] = m.group(0) if m else ''
    if d['release_agent']:
        cg['carrier_agent'] = d['release_agent']
    d['mawb'] = ''
    cg['hawb'] = d['hawb']
    d['pieces'] = cg.get('packages_qty')
    d['gw'] = cg.get('gw_kg')
    d['departure'] = cg.get('loading_name', '')
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    if not d['hawb']:
        d['notes'].append('B/L number not readable - type the HBL on tab 3')
    return d


def parse_ci(path, pages):
    """DHL 'Sea Freight FCL Arrival Customs Instructions Import'."""
    text = '\n'.join(p['text'] for p in pages)
    d = _cargo_doc(path, pages, 'Customs instructions')
    cg = {}
    m = re.search(r'Sea Freight|SEA\b', text)
    cg['mode'] = 'Sea' if m else 'Air'
    m = re.search(r'ORDER NUMBERS / REFERENCE\s+OCEAN BILL OF LADING\s+HOUSE BILL OF LADING\s*\n\s*(?:SRN:\s*([\d,\s]+?)\s+)?([A-Z0-9]{8,})\s+([A-Z0-9]{6,})', text)
    if m:
        d['invoice_refs'] = re.findall(r'\d{6,}', m.group(1) or '')
        d['mawb'], d['hawb'] = m.group(2), m.group(3)
    else:
        m = re.search(r'HOUSE (?:BILL OF LADING|AIR ?WAYBILL)[^\n]*\n[^\n]*?([A-Z]{4}\d{5,})', text)
        d['hawb'] = m.group(1) if m else ''
    cg['mawb'], cg['hawb'] = d['mawb'], d['hawb']
    m = re.search(r'^(SEA|AIR)\s+(.+?)\s*/\s*(\S+)\s*/\s*.*?\b([A-Z]{5})\s+([A-Z]{5})\s+(\d{1,2}-[A-Za-z]{3}-\d{2,4})\s+(\d{1,2}-[A-Za-z]{3}-\d{2,4})', text, re.M)
    if m:
        cg['transport_id'], cg['flight'] = E.clean(m.group(2)).upper(), m.group(3)
        cg['loading_code'], cg['discharge_code'] = m.group(4), m.group(5)
        cg['loading_name'] = CODE_NAME.get(m.group(4), '')
        cg['discharge_name'] = CODE_NAME.get(m.group(5), '')
        cg['etd'], cg['date'] = _d(m.group(6)), _d(m.group(7))
    m = re.search(r'PACKAGES\s+WEIGHT\s+VOLUME.*\n\s*(\d+)\s*([A-Z]{2,4})\b.*?(' + NUM + r')\s*KG', text)
    if m:
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = int(m.group(1)), m.group(2), E.fnum(m.group(3))
    conts = []
    for mc in re.finditer(r'CONTAINER:\s*([A-Z]{4}\d{7})\s+TYPE:\s*(\S+)\s+PACKS:\s*(\d+).*?\n\s*SEAL:\s*(\S+).*?GROSS:\s*(' + NUM + r')\s*KG', text):
        conts.append(dict(no=mc.group(1), size=re.sub(r'\D', '', mc.group(2))[:2], size_type=mc.group(2),
                          packages=int(mc.group(3)), seal=mc.group(4), tne=E.fnum(mc.group(5)) / 1000.0))
    if conts:
        cg['containers'] = [c['no'] for c in conts]
        cg['containers_info'] = conts
    m = re.search(r'INCO TERM\s*\n\s*([A-Z]{3})\b', text)
    d['incoterm'] = m.group(1) if m else ''
    for key, lab in (('release', 'Place of Release'), ('receipt', 'Place of Receipt'), ('permit_type', 'Type of Permit'),
                     ('permit_remarks', 'Permit Remarks'), ('special', 'Special Instructions')):
        m = re.search(lab + r'[ \t]*:[ \t]*([^\n]*)', text, re.I)
        v = E.clean(m.group(1)) if m else ''
        if v:
            cg[key] = v
    if cg.get('permit_type'):
        cg['permit_type'] = cg['permit_type'].upper()[:2]
    m = re.search(r'(?:FRT|FREIGHT)\s*[:=]?\s*(' + NUM + r')\s*([A-Z]{3})|(?:FRT|FREIGHT)\s*[:=]?\s*([A-Z]{3})\s*(' + NUM + ')',
                  cg.get('permit_remarks', ''), re.I)
    if m:
        cg['freight_amount'] = E.fnum(m.group(1) or m.group(4))
        cg['freight_currency'] = (m.group(2) or m.group(3)).upper()
    m = re.search(r'SHIPPER\s+CONSIGNEE\s*\n\s*(.+?\bLTD\.?)\s+(.+?)\s*$', text, re.M)
    if m:
        d['shipper'], d['consignee'] = E.clean(m.group(1)), E.clean(m.group(2))
    m = re.search(r'EXPORT AGENT\s+IMPORT AGENT\s*\n(?:.*@.*\n)?\s*(.+?(?:LTD\.?))\s+(.+?(?:LTD\.?))\s*$', text, re.M)
    if m:
        d['issued_by'] = E.clean(m.group(1))
        cg['import_agent'] = E.clean(m.group(2))
    m = re.search(r'Co\.?\s*Reg\.?\s*No:?\s*([0-9]{8,9}[A-Z])', text)
    if m and cg.get('import_agent'):
        cg['carrier_agent'] = cg['forwarder'] = cg['import_agent']
        cg['carrier_agent_uen'] = cg['forwarder_uen'] = m.group(1)
    m = re.search(r'SHIPMENT\s*(S\d{6,})', text)
    d['job_ref'] = m.group(1) if m else ''
    d['pieces'], d['gw'] = cg.get('packages_qty'), cg.get('gw_kg')
    d['departure'] = cg.get('loading_name', '')
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}
    if not d['hawb']:
        d['notes'].append('house B/L number not readable')
    return d


# ======================================================================== linking + checks
def link_docs(docs, gi):
    """Invoices carry no HBL: link them to the B/L / customs instructions by invoice no. (SRN / B/L marks) or by the
    Festo transport number; invoices of one transport share the transport gross / net weight and packages."""
    invs = [d for d in docs if d.get('kind') == 'INV' and d.get('supplier') in ('FESTO', 'DANFOSS')]
    cargo = [d for d in docs if d.get('kind') == 'AWB' and d.get('hawb')]
    for inv in invs:
        if inv.get('hawb'):
            continue
        hit = [a for a in cargo if inv['invoice_no'] in a.get('invoice_refs', [])] or \
              [a for a in cargo if inv.get('transport') and inv['transport'] in a.get('marks', [])]
        if not hit and inv.get('supplier') == 'DANFOSS':
            dns = {l.get('delivery') for l in inv['lines'] if l.get('delivery')}
            hit = [a for a in cargo if dns & set(a.get('deliveries') or [])]
            if not hit and len({a['hawb'] for a in cargo}) == 1 and not any(
                    (a.get('deliveries') and dns) or (a.get('marks') and inv.get('transport')) for a in cargo):
                hit = cargo                    # nothing to compare: the only B/L
        elif not hit and len({a['hawb'] for a in cargo}) == 1:
            hit = cargo
        if hit:
            inv['hawb'], inv['hawb_from'] = hit[0]['hawb'], '%s %s' % (hit[0].get('doc_type', 'B/L'), hit[0]['file'])
    # transport totals once per transport
    seen = {}
    for inv in invs:
        t = inv.get('transport')
        if not t:
            continue
        if t in seen:
            first = seen[t]
            first['packages_ids'] = list(dict.fromkeys(first['packages_ids'] + inv['packages_ids']))
            first['boxes'] = len(first['packages_ids']) or first['boxes']
            first.setdefault('shared_with', []).append(inv['invoice_no'])
            inv['total_gw'] = inv['total_nw'] = None
            inv['boxes'] = None
            inv['transport_shared'] = first['invoice_no']
        else:
            seen[t] = inv
    for t, first in seen.items():
        if first.get('shared_with'):
            first['notes_shared'] = ('Transport %s: gross %s KG / net %s KG and %d packages are for invoices %s together '
                                     '(counted once)' % (t, E.fmt_thou(first['total_gw']), E.fmt_thou(first['total_nw']),
                                                         first['boxes'] or 0,
                                                         ', '.join([first['invoice_no']] + first['shared_with'])))
    # packing lists -> HBL by transport number
    for p in docs:
        if p.get('kind') == 'PL' and p.get('supplier') in ('FESTO', 'DANFOSS') and not p.get('hawb'):
            inv = next((i for i in invs if i.get('transport') and i['transport'] == p.get('transport')), None)
            if inv is not None:
                p['hawb'] = inv['hawb']
                p['invoice_refs'] = [i['invoice_no'] for i in invs if i.get('transport') == p['transport']]


def check_pl(sh):
    """Festo packing list vs invoices: qty / COO / HS per delivery + part no., packages and weights."""
    pls = [p for p in sh.get('pls', []) if p.get('supplier') == 'FESTO']
    if not pls:
        return
    iss = sh['issues']
    pq = collections.OrderedDict()
    for p in pls:
        for l in p['lines']:
            k = (l['delivery'], l['item'])
            m = pq.setdefault(k, dict(qty=0.0, coo=set(), hs=set(), nw=0.0, desc=l['desc']))
            m['qty'] += l['qty'] or 0
            m['coo'].add(l['coo'])
            m['hs'].add(l['hs'])
            m['nw'] += l['nw'] or 0
    iq = collections.OrderedDict()
    for inv in sh['invs']:
        for l in inv['lines']:
            m = iq.setdefault((l['delivery'], l['item']), dict(qty=0.0, coo=l['coo'], hs=l['hs'], nw=0.0, inv=inv['invoice_no'],
                                                             pos=l['line']))
            m['qty'] += l['qty'] or 0
            m['nw'] += l['nw'] or 0
    bad = 0
    for k, m in iq.items():
        p = pq.get(k)
        if not p:
            bad += 1
            if bad <= 25:
                iss.append(('CHECK', 'Invoice %s pos %s part %s (delivery %s) not on the packing list' % (m['inv'], m['pos'], k[1], k[0])))
            continue
        if abs(p['qty'] - m['qty']) > 0.0005:
            iss.append(('MISMATCH', 'Part %s (delivery %s) qty: invoice %s, packing list %s' % (k[1], k[0], E.fmt_plain(m['qty']), E.fmt_plain(p['qty']))))
        if m['coo'] and p['coo'] and m['coo'] not in p['coo']:
            iss.append(('MISMATCH', 'Part %s COO: invoice %s, packing list %s' % (k[1], m['coo'], '/'.join(sorted(p['coo'])))))
        if m['hs'] and p['hs'] and m['hs'] not in p['hs'] - {''}:
            iss.append(('MISMATCH', 'Part %s HS: invoice %s, packing list %s' % (k[1], m['hs'], '/'.join(sorted(p['hs'] - {''})))))
    extra = [k for k in pq if k not in iq]
    if extra:
        iss.append(('CHECK', '%d packing list row(s) not on these invoices: %s' % (len(extra), ', '.join('%s/%s' % k for k in extra[:12]))))
    pk = sum(p['packages'] or 0 for p in pls)
    ib = sum(i['boxes'] or 0 for i in sh['invs'])
    if pk and ib and pk != ib:
        iss.append(('MISMATCH', 'Packages: invoices %s, packing list %s' % (ib, pk)))
    pg = sum(p['total_gw'] or 0 for p in pls)
    ig = sum(i['total_gw'] or 0 for i in sh['invs'])
    if pg and ig and abs(pg - ig) > 0.5:
        iss.append(('MISMATCH', 'Gross weight: invoice %s KG, packing list %s KG' % (E.fmt_thou(ig), E.fmt_thou(pg))))
    pn = sum(p['total_nw'] or 0 for p in pls)
    inn = sum(i['total_nw'] or 0 for i in sh['invs'])
    if pn and inn and abs(pn - inn) > 0.5:
        iss.append(('CHECK', 'Net weight: invoice %s KG, packing list %s KG' % (E.fmt_thou(inn), E.fmt_thou(pn))))


def read_pdf(path, pages, countries=None):
    """Festo documents in one PDF (or None when nothing Festo / B/L / CI is in it)."""
    kinds = [classify(p['text']) for p in pages]
    if not any(kinds):
        return None
    k = next(x for x in kinds if x)
    if k == 'FINV2':
        docs, cur = [], []
        for p in pages:
            if cur and re.search(r'\b1 of \d+\b', p['text'][:1500]):
                docs.append(cur)
                cur = []
            cur.append(p)
        if cur:
            docs.append(cur)
        return [parse_invoice2(path, ps, countries) for ps in docs]
    if k == 'FINV':
        # one SPOOL can hold several invoices: split at "Page 1 of"
        docs, cur = [], []
        for p in pages:
            if cur and re.search(r'\b1 of \d+\b', p['text'][:900]) and re.search(r'^Invoice\s*$', p['text'][:300], re.M):
                docs.append(cur)
                cur = []
            cur.append(p)
        if cur:
            docs.append(cur)
        return [parse_invoice(path, ps, countries) for ps in docs]
    if k == 'FPL':
        return [parse_pl(path, pages)]
    if k == 'CI':
        return [parse_ci(path, pages)]
    if k == 'BL':
        return [parse_bl(path, pages)]
    return None
