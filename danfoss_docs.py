# -*- coding: utf-8 -*-
"""Danfoss documents: Commercial Invoice ("Internal number"), packing list per delivery note, the US shipment pack
(short-form B/L, instruction to forwarder, shipper's letter of instructions - skipped), plus the DHL customs
instructions (LCL / FCL, multi-leg routing) and the Danmar / Danzas express sea waybill.
Returns the common document shape of danfoss_engine; linking via the Danfoss shipping number (e.g. 6366296)."""
import os, re, collections
import danfoss_engine as E

NUM = r'\d[\d,]*(?:\.\d+)?'
COO_NAMES = {'USA': 'US', 'U.S.A.': 'US', 'UNITED STATES': 'US', 'UNITED STATES OF AMERICA': 'US', 'GERMANY': 'DE',
             'FRANCE': 'FR', 'DENMARK': 'DK', 'CHINA': 'CN', 'ITALY': 'IT', 'POLAND': 'PL', 'SLOVAKIA': 'SK',
             'SLOVENIA': 'SI', 'MEXICO': 'MX', 'INDIA': 'IN', 'JAPAN': 'JP', 'BRAZIL': 'BR', 'UNITED KINGDOM': 'GB',
             'GREAT BRITAIN': 'GB', 'CZECH REPUBLIC': 'CZ', 'CZECHIA': 'CZ', 'HUNGARY': 'HU', 'FINLAND': 'FI',
             'SWEDEN': 'SE', 'NETHERLANDS': 'NL', 'BELGIUM': 'BE', 'SPAIN': 'ES', 'SWITZERLAND': 'CH', 'AUSTRIA': 'AT',
             'KOREA': 'KR', 'SOUTH KOREA': 'KR', 'TAIWAN': 'TW', 'MALAYSIA': 'MY', 'THAILAND': 'TH', 'VIETNAM': 'VN',
             'BULGARIA': 'BG', 'ROMANIA': 'RO', 'LITHUANIA': 'LT', 'CANADA': 'CA', 'SINGAPORE': 'SG', 'ESTONIA': 'EE',
             'INDONESIA': 'ID', 'PHILIPPINES': 'PH', 'TURKEY': 'TR', 'ISRAEL': 'IL', 'IRELAND': 'IE', 'PORTUGAL': 'PT',
             'NORWAY': 'NO', 'AUSTRALIA': 'AU', 'SOUTH AFRICA': 'ZA', 'HONG KONG': 'HK', 'CROATIA': 'HR'}
MON = dict(JAN=1, FEB=2, MAR=3, APR=4, MAY=5, JUN=6, JUL=7, AUG=8, SEP=9, OCT=10, NOV=11, DEC=12)


def coo_code(name, C=None):
    n = (name or '').strip().upper().strip(' .')
    if len(n) == 2:
        return n
    return COO_NAMES.get(n) or (C.code(n) if C else '')


def _mdy(s):
    """Jun 19, 2026 -> 19/06/2026"""
    m = re.match(r'([A-Za-z]{3})[a-z]*\.?\s+(\d{1,2}),?\s+(\d{4})', (s or '').strip())
    if m and m.group(1).upper() in MON:
        return '%02d/%02d/%s' % (int(m.group(2)), MON[m.group(1).upper()], m.group(3))
    return ''


def classify(text):
    t = text or ''
    if re.search(r'^Commercial Invoice', t, re.M) and re.search(r'Internal number\s+\d+', t):
        return 'DINV'
    if re.search(r'Packing\s*list\s*for\s*delivery\s*note\s*no[.,]?\s*\d+', t, re.I):
        return 'DPL'
    if re.search(r'BILL OF LADING - SHORT FORM|Instruction to forwarder|SHIPPER.S LETTER OF INSTRUCTIONS|'
                 r'CONTINUATION SHEET TO SHIPPER', t[:400], re.I):
        return 'DSKIP'                          # US export papers in the Danfoss pack - not needed for the permit
    return ''


# ======================================================================== commercial invoice
POS = re.compile(r'^(\d{1,4})\s+(' + NUM + r')\s+([A-Z]{1,3})\s+(\S+)\s+(' + NUM + r')\s+(' + NUM + r')\s*$')


def parse_invoice(path, pages, countries=None):
    C = countries or E.Countries()
    text = '\n'.join(p['text'] for p in pages)
    hdr = '\n'.join(p.get('hdr_text', '') for p in pages)      # scanned: header table read again without grid lines
    htext = text + '\n' + hdr
    d = dict(kind='INV', supplier='DANFOSS', file=os.path.basename(path), path=path, notes=[], lines=[],
             title='COMMERCIAL INVOICE', reader='Danfoss commercial invoice', ocr=any(p.get('ocr') for p in pages))
    m = re.search(r'Internal number\s+(\d{6,})', text)
    d['invoice_no'] = m.group(1) if m else ''
    m = re.search(r'Original Date\s+([A-Za-z]{3}[a-z]*\.?\s+\d{1,2},?\s+\d{4})', text)
    d['invoice_date'] = _mdy(m.group(1)) if m else ''
    m = re.search(r'Currency\s+([A-Z]{3})\b', htext) or re.search(r'^Total\s+([A-Z]{3})\s+\d', text, re.M) or \
        re.search(r'Currency:\s*([A-Z]{3})\b', text)
    d['currency'] = m.group(1) if m else ''
    m = re.search(r'Incoterms\s+([A-Z]{3})\s*([A-Za-z][A-Za-z .-]*?)?\s+(?:Shipping number|$)', htext, re.M)
    d['incoterm'] = m.group(1) if m else ''
    d['incoterm_place'] = E.clean(m.group(2) or '') if m else ''
    m = re.search(r'Shipping number\s+(\d{5,})', htext)
    d['transport'] = m.group(1) if m else ''
    m = re.search(r'Customer\s*VAT\s*no\.?\s*[-:\W]?\s*([0-9]{8,9}[A-Z]|[A-Z]\d{2}[A-Z]{2}\d{4}[A-Z])', htext)
    d['gst_no'] = m.group(1) if m else ''
    d['uen'] = d['gst_no']                    # SG companies: GST reg. no. = UEN (reported CHECK)
    m = re.search(r'Danfoss order no[.,]?\s*\W?\s*(\d{8,})', htext)
    d['order_no'] = m.group(1) if m else ''
    d['dn_header'] = sorted(set(re.findall(r'Delivery\s*note\s*no[.,]?\s*(\d{10})', htext)))
    m = re.search(r'Total gross\s*\n?\s*(' + NUM + r')\s*(KG|G|LB)', text)
    d['total_gw'] = E.fnum(m.group(1)) if m else None
    m = re.search(r'Total net\s*\n?\s*(' + NUM + r')\s*(KG|G|LB)', text)
    d['total_nw'] = E.fnum(m.group(1)) if m else None
    if d['total_gw'] is None or d['total_nw'] is None:      # scanned: "Total gross / net" labels garbled by OCR
        ws = [E.fnum(x) for x in re.findall(r'(' + NUM + r')\s*[-~]?\s*K?G\b[^\n]*\n\s*weight', text)]
        if len(ws) >= 2:
            d['total_gw'] = d['total_gw'] if d['total_gw'] is not None else ws[0]
            d['total_nw'] = d['total_nw'] if d['total_nw'] is not None else ws[1]
        elif len(ws) == 1 and d['total_gw'] is None:
            d['total_gw'] = ws[0]
    m = re.search(r'Number of\s*\n?\s*(\d+)\s*\n?\s*packages', text)
    d['boxes'] = int(m.group(1)) if m else None
    d['packages_ids'] = []
    m = re.search(r'^Total\s+([A-Z]{3})\s+(' + NUM + ')', text, re.M)
    d['total'] = E.fnum(m.group(2)) if m else None
    m = re.search(r'^Net value\s+(' + NUM + ')', text, re.M)
    d['net_value'] = E.fnum(m.group(1)) if m else None
    d['charges'] = 0.0
    if d['total'] is not None and d['net_value'] is not None and d['total'] - d['net_value'] > 0.005:
        d['charges'] = round(d['total'] - d['net_value'], 2)
        d['charges_label'] = 'Tax / other value'
    d['freight'] = 0.0
    m = re.search(r'Forwarding agent[^\n]*\n?[^\n]*?(DHL GLOBAL FORWARDING|KUEHNE|SCHENKER|EXPEDITORS|DSV|CEVA|KINTETSU)', text, re.I) or \
        re.search(r'Your reference\s+Orig\. doc\. ref\.\s*\n\s*Contact person:\s*([A-Z][A-Z &]+)', text)
    d['carrier'] = E.clean(m.group(1)).upper() if m else ''
    m = re.search(r'Customer address[^\n]*\n\s*(.+)', text)
    d['invoice_to'] = d['ship_to'] = E.clean(m.group(1)).upper() if m else ''
    d['party_name'] = d['invoice_to']
    d['invoice_to_block'] = d['ship_to_block'] = [d['invoice_to']]
    m = re.search(r'\n(Danfoss [^\n]*?(?:LLC|GmbH|A/S|Inc\.?|S\.A\.|Ltd\.?))\s+\d+ of \d+', text)
    d['invoice_from'] = d['ship_from'] = _roman(E.clean(m.group(1)).upper()) if m else 'DANFOSS'
    d['ship_to_country'] = 'SG' if re.search(r'SINGAPORE', text[:600], re.I) else ''
    d['ship_from_country'] = ''
    d['direction'] = 'IMPORT' if d['ship_to_country'] == 'SG' else 'EXPORT'
    d['hawb'] = ''
    d['tax_ids'] = []
    d['ship_mode'] = ''
    cur = None
    for p in pages:
        for raw in p['text'].splitlines():
            ln = raw.strip()
            m = POS.match(ln)
            if m and int(m.group(1)) % 1 == 0 and not re.match(r'^\d{1,4}\s+of\s+\d', ln):
                cur = dict(line=m.group(1).zfill(4), qty=E.fnum(m.group(2)), uom=m.group(3), item=m.group(4),
                           price=E.fnum(m.group(5)), amount=E.fnum(m.group(6)), desc='', type_code='', hs='', coo='',
                           nw=None, gw=None, currency=d['currency'], order='', delivery='', eccn='', license='', pl='',
                           lots=[], boxes=[], cust_item='', need_desc=True)
                d['lines'].append(cur)
                continue
            if cur is None:
                continue
            if cur.get('need_desc'):
                if re.match(r'^[0-9A-Z]{10,}\s*/\s*1\s*[A-Z]{1,3}\s*$', ln) and re.search(r'0{5,}', ln):
                    cur['cust_item'] = ln.split('/')[0].strip().lstrip('0')   # customer item no. line - description follows
                    continue
                if not re.match(r'(Country of origin|Commodity code|US ECCN|Net weight|Danfoss order|Delivery note|Subtotal)', ln):
                    cur['desc'] = E.clean(re.sub(r'\s*/\s*1\s*[A-Z]{1,3}\s*$', '', ln))
                cur['need_desc'] = False
            m = re.match(r'Country of origin\s+(.+)$', ln)
            if m:
                cur['coo'] = coo_code(m.group(1), C)
                if not cur['coo']:
                    d['notes'].append('Position %s: country of origin "%s" not known' % (cur['line'], m.group(1)))
            m = re.match(r'Commodity code\s+(\d{6,10})', ln)
            if m:
                cur['hs'] = m.group(1)
            m = re.match(r'US ECCN\s+(\S+)', ln)
            if m:
                cur['eccn'] = m.group(1)
            m = re.match(r'Net weight\s+(' + NUM + r')\s*(KG|G|LB)', ln)
            if m:
                w = E.fnum(m.group(1))
                cur['nw'] = w / 1000.0 if m.group(2) == 'G' else w * 0.45359237 if m.group(2) == 'LB' else w
            m = re.match(r'Danfoss order no\.?\s+(\d+)', ln)
            if m:
                cur['order'] = m.group(1)
            m = re.match(r'Delivery note no\.?\s+(\d+)', ln)
            if m:
                cur['delivery'] = m.group(1)
    for l in d['lines']:
        l.pop('need_desc', None)
        if not l['delivery'] and len(d['dn_header']) == 1:
            l['delivery'] = d['dn_header'][0]          # scanned single-delivery invoice: delivery note from the header
        if not l['order'] and d['order_no']:
            l['order'] = d['order_no']
    if not d['lines']:
        d['notes'].append('no invoice positions could be read')
    d['line_nw'] = round(sum(l['nw'] or 0 for l in d['lines']), 3)
    if d['total_nw'] is None and d['line_nw']:
        d['total_nw'] = d['line_nw']
        if d['ocr']:
            d['notes'].append('total net weight not readable (scan) - sum of the position net weights used')
    if d['ocr'] and not d['incoterm']:
        d['notes'].append('Incoterms not readable on the scanned invoice - check Term Type on tab 3')
    if d['ocr']:
        tot = round(sum(l['amount'] or 0 for l in d['lines']), 2)
        if d['total'] is not None and abs(tot - (d['net_value'] or d['total'])) > 0.01:
            d['notes'].append('scanned invoice: positions add up to %.2f but the invoice says %.2f - check' %
                              (tot, d['net_value'] or d['total']))
        for l in d['lines']:
            if l['qty'] and l['price'] and l['amount'] and abs(l['qty'] * l['price'] - l['amount']) > 0.02 * l['amount'] + 0.01:
                q = round(l['amount'] / l['price'], 3)
                if abs(q - round(q)) < 0.01:
                    d['notes'].append('Position %s: OCR qty %s does not match price x qty - %s used' % (l['line'], l['qty'], round(q)))
                    l['qty'] = float(round(q))
    return d


def _roman(s):
    """OCR: 'SOLUTIONS I] LLC' / '{{' / 'Il' / 'II]' -> 'II'."""
    return re.sub(r'(?<=\s)[Il1\]\[\{\}|!]{2,3}(?=[\s,])', 'II', s)


# ======================================================================== packing list per delivery note
PL_ROW = re.compile(r'^(\d+(?:\.\d+)?)\s+(\d{17,20})\s+(?:(' + NUM + r')\s+([A-Z]{1,3})\s+(\S+)\s+)?(' + NUM + r')\s*$')
PL_EXTRA = re.compile(r'^(' + NUM + r')\s+([A-Z]{1,3})\s+(\S+)\s*$')


def parse_pl(path, pages):
    """All packing lists in the PDF (one per delivery note) -> ONE packing list document."""
    text = '\n'.join(p['text'] for p in pages)
    d = dict(kind='PL', supplier='DANFOSS', file=os.path.basename(path), path=path, reader='Danfoss packing list',
             notes=[], invoice_refs=[], hawb='', mawb='', total_gw=None, total_nw=None, packages=None,
             package_unit='PKG', containers=[], lines=[], ocr=any(p.get('ocr') for p in pages), deliveries=[],
             transport='', pkgs={})
    gws, nws, pks = {}, {}, {}
    dn = ''
    lines = text.splitlines()
    for i, raw in enumerate(lines):
        ln = raw.strip()
        if d['ocr']:
            ln = re.sub(r'^(\d{1,3}\s+)(\d{4,})\s+(\d{4,})(?=\s)',
                        lambda k: k.group(1) + k.group(2) + k.group(3) if 17 <= len(k.group(2) + k.group(3)) <= 20
                        else k.group(0), ln)                       # SSCC split by OCR
            ln = re.sub(r'^\d+\s+(?:Pallets?|Boxes|Box|Cartons?)\s+(?=\d)', '', ln)   # "1 Pallets 1 PC 119-..."
            ln = re.sub(r'\s*[;:,|.]+\s*$', '', ln)
            lines[i] = ln
        m = re.search(r'Packing\s*list\s*for\s*delivery\s*note\s*no[.,]?\s*(\d+)', ln, re.I)
        if m:
            dn = m.group(1)
            if dn not in d['deliveries']:
                d['deliveries'].append(dn)
            continue
        m = re.search(r'Shipping number\s+(\d{5,})', ln)
        if m:
            d['transport'] = m.group(1)
        if ln.startswith('Total gross') and i + 1 < len(lines):
            mm = re.match(r'(' + NUM + r')\s*KG', lines[i + 1].strip())
            if mm:
                gws[dn] = E.fnum(mm.group(1))
        if ln.startswith('Total net') and i + 1 < len(lines):
            mm = re.match(r'(' + NUM + r')\s*KG', lines[i + 1].strip())
            if mm:
                nws[dn] = E.fnum(mm.group(1))
        if ln == 'Number of' and i + 1 < len(lines) and re.match(r'\d+$', lines[i + 1].strip()):
            pks[dn] = int(lines[i + 1].strip())
        m = PL_ROW.match(ln)
        if m:
            if m.group(5):
                nxt = lines[i + 1].strip() if i + 1 < len(lines) else ''
                mn = re.search(r'(' + NUM + r')\s*$', nxt)
                d['lines'].append(dict(line=m.group(1), delivery=dn, item=m.group(5), qty=E.fnum(m.group(3)),
                                       uom=m.group(4), gw=E.fnum(m.group(6)), nw=E.fnum(mn.group(1)) if mn else None,
                                       coo='', desc='', hs='', boxes=[m.group(2)], lots=[], amount=None, price=None,
                                       currency='', eccn='', order='', license='', pl=''))
            d['pkgs'][m.group(2)] = E.fnum(m.group(6))
            continue
        m = PL_EXTRA.match(ln)
        if m and dn and d['lines'] and not re.match(r'\d{2}\.\d', ln):
            d['lines'].append(dict(line='', delivery=dn, item=m.group(3), qty=E.fnum(m.group(1)), uom=m.group(2),
                                   gw=None, nw=None, coo='', desc='', hs='', boxes=[], lots=[], amount=None, price=None,
                                   currency='', eccn='', order='', license='', pl=''))
    if not gws and d['deliveries']:                          # scanned: "Total gross" label garbled by OCR
        text = '\n'.join(lines)
        ws = [E.fnum(x) for x in re.findall(r'(' + NUM + r')\s*[-~]?\s*K?G\b[^\n]*\n\s*weight', text)]
        if len(ws) < 2:
            ws = [E.fnum(x) for x in re.findall(r'(?:gross|net)\s*\n\s*weight\s+(' + NUM + r')\s*K?G', text)]
        if len(ws) >= 2 and len(d['deliveries']) == 1:
            gws[d['deliveries'][0]], nws[d['deliveries'][0]] = ws[0], ws[1]
    for x in d['deliveries']:
        if x not in pks:
            n = len({b for l in d['lines'] if l['delivery'] == x for b in l['boxes']})
            if n:
                pks[x] = n                               # packages = pallets / boxes listed
    nototal = [x for x in d['deliveries'] if x not in gws]
    if nototal:
        d['notes'].append('packing list %s: totals page missing or blank in the PDF - rows / weights may be incomplete'
                          % ', '.join(nototal))
    d['total_gw'] = round(sum(gws.values()), 3) or None
    d['total_nw'] = round(sum(nws.values()), 3) or None
    d['packages'] = sum(pks.values()) or None
    d['pl_by_delivery'] = dict(gw=gws, nw=nws, pk=pks)
    if not d['lines']:
        d['notes'].append('packing list rows could not be read')
    return d


# ======================================================================== waybill / customs instructions (Danfoss layouts)
def improve_ci(d, text):
    """DHL customs instructions with multi-leg routing (truck / sea) and ORDER NUMBERS block 'ASL ... DNN ... SRN'."""
    cg = d['cargo']
    m = re.search(r'ORDER NUMBERS / REFERENCE\s+OCEAN BILL OF LADING\s+HOUSE BILL OF LADING\s*\n(.*?)\s+([A-Z0-9]{6,})\s+([A-Z]{3,4}[A-Z0-9]{4,})\s*$',
                  text, re.M)
    if m and (not d.get('hawb') or not d.get('mawb') or len(d.get('mawb', '')) < 6):
        d['mawb'], d['hawb'] = m.group(2), m.group(3)
        cg['mawb'], cg['hawb'] = m.group(2), m.group(3)
    refs = re.findall(r'SRN:\s*([\d, ]+)', text)
    d['marks'] = list(dict.fromkeys(re.findall(r'\d{5,}', ' '.join(refs))))
    d['deliveries'] = list(dict.fromkeys(re.findall(r'\b\d{10}\b', ' '.join(re.findall(r'DNN:\s*([\d,\s]+?)(?:\s+[A-Z]|$)', text, re.M)))))
    ms = re.search(r'^SEA\s+(.+?)\s*/\s*(\S*)\s*/?\s*.*?\b([A-Z]{5})\s+([A-Z]{5})\s+(\d{1,2}-[A-Za-z]{3}-\d{2,4})\S*\s+(?:\d{2}:\d{2}\s+)?(\d{1,2}-[A-Za-z]{3}-\d{2,4})', text, re.M)
    if ms:
        vsl, voy = E.clean(ms.group(1)).upper(), ms.group(2)
        if not voy:
            mv = re.search(r'^SEA\s+.*\n\s*([A-Z0-9]{2,8})\s*$', text, re.M)
            voy = mv.group(1) if mv else ''
        cg.update(mode='Sea', transport_id=vsl, flight=voy, loading_code=ms.group(3), discharge_code=ms.group(4))
        import danfoss_festo_docs as F
        cg['etd'], cg['date'] = F._d(ms.group(5)), F._d(ms.group(6))
        cg['loading_name'] = F.CODE_NAME.get(ms.group(3), cg.get('loading_name', ''))
        cg['discharge_name'] = F.CODE_NAME.get(ms.group(4), cg.get('discharge_name', ''))
    m = re.search(r'PACKAGES\s+WEIGHT\s+VOLUME.*\n\s*(\d+)\s*([A-Z]{2,4})\b.*?(' + NUM + r')\s*KG', text)
    if m:
        cg['packages_qty'], cg['packages_unit'], cg['gw_kg'] = int(m.group(1)), m.group(2), E.fnum(m.group(3))
    if re.search(r'Sea Freight LCL', text, re.I):
        cg['lcl'] = True
        cg.pop('containers', None)
        cg.pop('containers_info', None)
    mc = re.search(r'SHIPPER\s+CONSIGNEE\s*\n\s*(.+?\b(?:LLC|LTD\.?|GMBH|INC\.?|A/S|AG|S\.A\.|CO\.))\s+(.+?)\s*$', text, re.M)
    if mc:
        d['shipper'], d['consignee'] = E.clean(mc.group(1)), E.clean(mc.group(2))
    m = re.search(r'IMPORT AGENT\s*\n(?:.*\n){0,3}?.*?([A-Z][A-Z &().,-]*\(SINGAPORE\)\s*PTE\.?\s*LTD\.?)', text)
    mu = re.search(r'Co\.?\s*Reg\.?\s*No:?\s*([0-9]{8,9}[A-Z])', text)
    if m:
        nm = E.clean(m.group(1))
        nm = nm[nm.rfind(nm.split()[0] + ' '):] if nm.split() and nm.rfind(nm.split()[0] + ' ') > 0 else nm  # export agent on the same line
        cg['import_agent'] = cg['carrier_agent'] = cg['forwarder'] = nm
        if mu:
            cg['carrier_agent_uen'] = cg['forwarder_uen'] = mu.group(1)
    m = re.search(r'^Carrier:[ \t]*(\S.*)$', text, re.M)
    if m:
        cg['shipping_line'] = E.clean(m.group(1))
    d['pieces'], d['gw'] = cg.get('packages_qty'), cg.get('gw_kg')
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}


def improve_bl(d, text):
    """Danmar / Danzas express sea waybill: SHIPPER REF, weight without 'KG', 'Total Number of Containers/Packages'."""
    cg = d['cargo']
    if not d.get('hawb'):
        m = re.search(r'B\.?\s*/?L\.?\s*No\.?\s*:?\s*([A-Z]{3,4}\d{5,})\b', text) or re.search(r'\b([A-Z]{3,4}\d{5,}),S\d{8,},HBL', text) or \
            re.search(r'B\.?L\.?\s*No\.?\s*\n(?:.*\n){0,3}?.*?\b([A-Z]{4}\d{5,})\s*$', text, re.M)
        if m:
            d['hawb'] = cg['hawb'] = m.group(1)
            d['notes'] = [n for n in d['notes'] if 'B/L number not readable' not in n]
    m = re.search(r'SHIPPER REF:\s*([\d, ]+)', text)
    if m:
        d['marks'] = list(dict.fromkeys(d.get('marks', []) + re.findall(r'\d{5,}', m.group(1))))
    if not cg.get('gw_kg'):
        m = re.search(r'Gross Weight\s+Measurement\s*\n\s*in kilos.*\n\s*(' + NUM + r')\s+(' + NUM + r')\s*$', text, re.M)
        if m:
            cg['gw_kg'] = E.fnum(m.group(1))
    if not cg.get('packages_qty'):
        m = re.search(r'Total Number of Containers/Packages:\s*(\d+)', text, re.I)
        if m:
            cg['packages_qty'], cg['packages_unit'] = int(m.group(1)), 'PKG'
    if not cg.get('loading_code'):
        import danfoss_festo_docs as F
        m = re.search(r'^PORT OF LOADING\s+(.+)$', text, re.M) or re.search(r'Port of Loading.*\n(?:.*\n)?.*?,\s*[A-Z]{2}\s+([A-Z ]+),\s*[A-Z]{2}', text)
        if m:
            c, n = F.port_code(m.group(1))
            if c:
                cg['loading_code'], cg['loading_name'] = c, n
    if not cg.get('discharge_code'):
        cg['discharge_code'], cg['discharge_name'] = ('SGSIN', 'SINGAPORE') if re.search(r'PORT OF DISCHARGE\s+SINGAPORE', text) else ('', '')
    m = re.search(r'Vessel\s+Voyage No\..*\n(?:.*\n){0,8}?\s*([A-Z][A-Z0-9 .\-]{3,}?)\s+(\d{2,4}[A-Z])\s*$', text, re.M) or \
        re.search(r'VESSEL / VOYAGE\s+(.+?)\s*/\s*(\S+)', text)
    if m:
        cg['transport_id'], cg['flight'] = E.clean(m.group(1)).upper(), m.group(2)
    d['pieces'], d['gw'] = cg.get('packages_qty'), cg.get('gw_kg')
    d['cargo'] = {k: v for k, v in cg.items() if v not in (None, '', [])}


# ======================================================================== linking / duplicates / checks
def dedupe(docs, gi):
    """The Danfoss pack comes as one PDF AND as split PDFs (invoice, packing lists): keep each document once."""
    if any(d.get('kind') == 'AWB' and (d.get('cargo') or {}).get('lcl') for d in docs):   # LCL: consol container, not declared
        for d in docs:
            if d.get('kind') == 'AWB' and (d.get('cargo') or {}).get('containers'):
                gi.append((d['file'], 'CHECK', 'LCL shipment - consol container %s not declared (add it on tab 3 if needed)'
                           % ', '.join(d['cargo']['containers'])))
                d['cargo'].pop('containers', None)
                d['cargo'].pop('containers_info', None)
    seen, out = set(), []
    for d in docs:
        if d.get('supplier') == 'DANFOSS' and d.get('kind') == 'INV':
            k = ('INV', d['invoice_no'], round(sum(l['amount'] or 0 for l in d['lines']), 2), len(d['lines']))
        elif d.get('supplier') == 'DANFOSS' and d.get('kind') == 'PL':
            k = None
            new = [x for x in d['deliveries'] if ('PL', x) not in seen]
            if d['deliveries'] and not new:
                gi.append((d['file'], 'CHECK', 'packing list(s) %s already read from another file - skipped'
                           % ', '.join(d['deliveries'])))
                continue
            if len(new) < len(d['deliveries']):       # keep only the new delivery notes of this file
                d['lines'] = [l for l in d['lines'] if l['delivery'] in new]
                for kk in ('gw', 'nw', 'pk'):
                    d['pl_by_delivery'][kk] = {x: v for x, v in d['pl_by_delivery'][kk].items() if x in new}
                d['total_gw'] = round(sum(d['pl_by_delivery']['gw'].values()), 3) or None
                d['total_nw'] = round(sum(d['pl_by_delivery']['nw'].values()), 3) or None
                d['packages'] = sum(d['pl_by_delivery']['pk'].values()) or None
                d['deliveries'] = new
            for x in d['deliveries']:
                seen.add(('PL', x))
            out.append(d)
            continue
        else:
            out.append(d)
            continue
        if k in seen:
            gi.append((d['file'], 'CHECK', 'invoice %s already read from another file - skipped (same invoice)' % d['invoice_no']))
            continue
        seen.add(k)
        out.append(d)
    docs[:] = out


def check_pl(sh):
    pls = [p for p in sh.get('pls', []) if p.get('supplier') == 'DANFOSS']
    if not pls:
        return
    iss = sh['issues']
    for p in pls:
        for n in p.get('notes', []):
            iss.append(('CHECK', '%s: %s' % (p['file'], n)))
    pq = collections.defaultdict(float)
    for p in pls:
        for l in p['lines']:
            pq[(l['delivery'], l['item'])] += l['qty'] or 0
    iq = collections.defaultdict(float)
    for inv in sh['invs']:
        for l in inv['lines']:
            iq[(l['delivery'], l['item'])] += l['qty'] or 0
    for k, q in iq.items():
        if k not in pq:
            if any(dd == k[0] for dd, _ in pq):
                iss.append(('CHECK', 'Part %s (delivery %s) not found on the packing list' % (k[1], k[0])))
            continue
        if abs(pq[k] - q) > 0.0005:
            iss.append(('MISMATCH', 'Part %s (delivery %s) qty: invoice %s, packing list %s' % (k[1], k[0], E.fmt_plain(q), E.fmt_plain(pq[k]))))
    pk = sum(p['packages'] or 0 for p in pls)
    ib = sum(i['boxes'] or 0 for i in sh['invs'])
    if pk and ib and pk != ib:
        iss.append(('CHECK', 'Packages: invoice %s, packing lists %s (packing lists read: %s)'
                    % (ib, pk, ', '.join(x for p in pls for x in p['deliveries']))))
    for inv in sh['invs']:
        if inv.get('gst_no'):
            iss.append(('CHECK', 'Importer UEN %s taken from the invoice "Customer VAT no." (GST no.) - check' % inv['gst_no']))


def _nogrid(img):
    """Remove table lines + grey cell shading so Tesseract reads the header table."""
    import numpy as np
    from PIL import Image
    a = np.array(img)
    dark = a < 160
    out = a.copy()
    H, W = a.shape
    def runs(v):
        d = np.diff(np.concatenate([[0], v.astype(np.int8), [0]]))
        return zip(np.where(d == 1)[0], np.where(d == -1)[0])
    for y in np.where(dark.sum(1) > W * 0.15)[0]:
        for a0, b0 in runs(dark[y]):
            if b0 - a0 > 120:
                out[y, a0:b0] = 255
    for x in np.where(dark.sum(0) > 60)[0]:
        for a0, b0 in runs(dark[:, x]):
            if b0 - a0 > 60:
                out[a0:b0, x] = 255
    out[out > 150] = 255
    return Image.fromarray(out)


def ocr_nogrid(path, idx):
    try:
        import pypdfium2 as pdfium
        import ocr_runner
        doc = pdfium.PdfDocument(path)
        img = doc[idx].render(scale=E.OCR['dpi'] / 72.0).to_pil().convert('L')
        doc.close()
        h = img.size[1]
        img = img.crop((0, 0, img.size[0], int(h * 0.62)))       # address + header table only
        return E.ocr_fix(ocr_runner.image_to_string(_nogrid(img), 6))
    except Exception:
        return ''


def read_pdf(path, pages, countries=None):
    kinds = [classify(p['text']) for p in pages]
    if not any(k in ('DINV', 'DPL') for k in kinds) and not all(k == 'DSKIP' or not p['text'].strip()
                                                              for k, p in zip(kinds, pages)):
        return None
    out = []
    inv = [p for p, k in zip(pages, kinds) if k == 'DINV']
    for i, (p, k) in enumerate(zip(pages, kinds)):
        if k == 'DINV' and p.get('ocr') and not re.search(r'Incoterms\s+[A-Z]{3}', p['text']) and \
                re.search(r'Internal number', p['text']) and 'hdr_text' not in p:
            p['hdr_text'] = ocr_nogrid(path, i)
            break
    # packing list pages: from a "Packing list" page until the next other document (blank / scanned pages included)
    pl, last = [], ''
    for p, k in zip(pages, kinds):
        if k:
            last = k
        if k == 'DPL' or (not k and last == 'DPL'):
            pl.append(p)
    if inv:
        out.append(parse_invoice(path, inv, countries))
    if pl:
        out.append(parse_pl(path, pl))
    if not out:
        return [dict(kind='', file=os.path.basename(path), path=path,
                     notes=['US export papers (short-form B/L / instructions / SLI) - not needed for the permit'])]
    return out
