# -*- coding: utf-8 -*-
"""Declaration advisor - Singapore Customs 'Module 5: Types of Permit Applications' (TradeNet, 25 Jun 2021) as rules.

After a shipment is worked out, advise() looks at the direction, the permit type / declaration type, the Place of
Release / Receipt codes and the items (dutiable, liquor / tobacco, controlled - from the UPDATE HS CODE list) and
  - works out the declaration type the documents point to (message type INP / IPT / OUT / TNP + declaration type),
  - lists what that declaration type needs (fields, approvals, time limits),
  - adds CHECK / MISMATCH / ERROR notes when the permit does not follow the rules.
Every tool engine uses it (toolrun.resolve_all / resolve_one call it after the engine's resolve)."""
import re

# permit type on the job / customs instructions -> (message type, declaration type)
PERMIT_DECL = {'IG': ('IPT', 'GST'), 'ME': ('INP', 'APS'), 'IT': ('INP', 'SFZ'), 'IM': ('INP', 'SFZ'),
               'OD': ('OUT', 'DRT'), 'OT': ('OUT', 'DRT'), 'ID': ('IPT', 'DNG')}
DECL_PERMIT = {('IPT', 'GST'): 'IG', ('INP', 'APS'): 'ME', ('INP', 'SFZ'): 'IT', ('OUT', 'DRT'): 'OD', ('IPT', 'DNG'): 'ID'}
MSG_NAME = {'INP': 'In-Non-Payment (import, no duty / GST now)', 'IPT': 'In-Payment (import, duty and / or GST paid)',
            'OUT': 'Out (export)', 'TNP': 'Transhipment / movement'}
DECL_NAME = {
    'APS': 'Approved Premises / Schemes', 'BKT': 'Blanket', 'DES': 'Destruction', 'GTR': 'GST Relief / Duty Exemption',
    'REX': 'Re-export', 'SFZ': 'Storage in FTZ', 'SHO': 'Shut Out', 'TCE': 'Temporary - exhibition',
    'TCS': 'Temporary - exhibition with sale', 'TCR': 'Temporary - repair / modification', 'TCO': 'Temporary - other approved purpose',
    'TCI': 'Temporary export / re-import', 'DNG': 'Duty & GST', 'DUT': 'Duty only', 'GST': 'GST (including duty exemption)',
    'DRT': 'Direct', 'BRE': 'Blanket Removal', 'REM': 'Removal', 'TTF': 'Through transhipment, same FTZ',
    'TTI': 'Through transhipment, inter-gateway', 'IGM': 'Inter-gateway movement'}
# what each declaration type needs (Module 5)
NEEDS = {
    ('INP', 'APS'): ['Licensed warehouse (dutiable): lot no. + product codes; bonded / ZG warehouse: non-dutiable only - never both in one permit',
                     'IRAS schemes: Place of Receipt ME / IGDS / AISS / ARCS; importer and declaring agent registered with IRAS',
                     'LCL with Class II / dangerous goods: DG Indicator = Y, HBL and lot no. on every item, one permit per container'],
    ('INP', 'SFZ'): ['Goods stay in the FTZ pending local release or re-export (no dutiable liquor / tobacco)',
                     'Same FTZ: Release = Receipt = the FTZ (e.g. KZ / KZ); road via Tuas into Keppel: Release THQ, Receipt KZ'],
    ('INP', 'REX'): ['Inward AND outward transport in the same permit; no Place of Storage',
                     'Outward details unknown or shipment split: INP SFZ first, then OUT DRT'],
    ('INP', 'DES'): ['Place of Receipt = DUMP; CA approval first for controlled goods',
                     "NEA date / time / venue + approval letter ref in Trader's Remarks; disposal certificate to Customs within 4 working days"],
    ('INP', 'GTR'): ['Place of Receipt: EM (diplomatic), INTORG, VEHSG, TRADESP (samples <= S$400), RELIEF (air <= S$400)',
                     'Personal effects: Declaration of Facts on ntp.gov.sg -> PE Ref No. in CA Licence No., permit within 5 working days'],
    ('INP', 'SHO'): ['Outward permit no. in Previous Permit No.; no inward transport; lot no. + product codes for dutiable goods'],
    ('INP', 'TCE'): ['End date of exhibition required; re-export within 21 days after it; banker\'s guarantee'],
    ('INP', 'TCS'): ['End date of exhibition required; re-export within 21 days after it; banker\'s guarantee'],
    ('INP', 'TCR'): ['End of temporary import period (max 6 months); banker\'s guarantee'],
    ('INP', 'TCO'): ['End of temporary import period (max 6 months); banker\'s guarantee'],
    ('INP', 'TCI'): ['Re-import of goods exported under OUT (TCI); samples / repair within 3 months'],
    ('IPT', 'GST'): ['Non-dutiable goods (GST only); dutiable goods only with a duty exemption (Receipt ETHYL / IEFxxx / VEHRAC) and claimant details'],
    ('IPT', 'DNG'): ['At least one dutiable item; short payment: Receipt SPNOSTK / SPSTK, recovery: RCNOSTK'],
    ('IPT', 'DUT'): ['Dutiable goods only; duty-exempt vehicle transfer: Release EXEMPT, Receipt O'],
    ('OUT', 'DRT'): ['Local / GST-paid goods: Release = O, Receipt = an FTZ or approved exit point; or re-export of goods stored in an FTZ'],
    ('OUT', 'APS'): ['Release = licence no. of the licensed / ZG warehouse, Receipt = FTZ or exit point',
                     'Sea stores: CPC SEASTORE, vessel type, NRT, crew (PC 1), voyage duration (PC 2), next / final port; shipyard: Receipt SY'],
    ('OUT', 'TCE'): ['Inward TIS permit no. in Previous Permit No.; same declaration type as the inward permit'],
    ('OUT', 'TCI'): ['Temporary export for repair / exhibition / rental - to be re-imported (INP TCI)'],
    ('TNP', 'REM'): ['Release / Receipt = licence nos. of both warehouses; previous and current lot no.'],
    ('TNP', 'IGM'): ['One permit per container; loading port ZZZZZ + COO ZZ allowed; non-controlled non-dutiable: HS 99999999 + HBL per line'],
}
for _t in ('TCS', 'TCR', 'TCO'):
    NEEDS[('OUT', _t)] = NEEDS[('OUT', 'TCE')]

FTZ = {'KZ': 'KEPPEL FTZ', 'PPZ': 'PASIR PANJANG FTZ', 'JZ': 'JURONG FTZ', 'CZ': 'CHANGI FTZ', 'TZ': 'TUAS TERMINAL GATEWAY',
       'SZ': 'SEMBAWANG FTZ', 'BZ': 'BRANI FTZ'}
# Place of Receipt codes that only belong to one declaration type
RECEIPT_IMPLIES = {
    'DUMP': ('INP', 'DES'), 'EM': ('INP', 'GTR'), 'INTORG': ('INP', 'GTR'), 'VEHSG': ('INP', 'GTR'),
    'TRADESP': ('INP', 'GTR'), 'RELIEF': ('INP', 'GTR'), 'ME': ('INP', 'APS'), 'IGDS': ('INP', 'APS'),
    'AISS': ('INP', 'APS'), 'ARCS': ('INP', 'APS'), 'ETHYL': ('IPT', 'GST'), 'VEHRAC': ('IPT', 'GST'),
    'SY': ('OUT', 'APS'), 'RCNOSTK': ('IPT', None), 'SPNOSTK': ('IPT', None), 'SPSTK': ('IPT', None)}
LIQUOR_TOBACCO = ('2203', '2204', '2205', '2206', '2207', '2208', '2402', '2403', '2404')


def _hs_info(hsl, hs):
    r = hsl.get(hs) or hsl.get((hs or '')[:8]) if hsl else None
    return (r[3] if r and len(r) > 3 and isinstance(r[3], dict) else {}) if r else {}


def _decl_from_text(t):
    t = (t or '').upper().strip()
    if 'RELIEF' in t:
        return 'GTR'
    if t.startswith('DIRECT'):
        return 'DRT'
    for code, name in DECL_NAME.items():
        if t == code or name.upper() in t or t.startswith(code + ' '):
            return code
    if 'DIRECT' in t:
        return 'DRT'
    if 'APPROVED PREMISES' in t:
        return 'APS'
    if 'STORAGE IN FTZ' in t:
        return 'SFZ'
    if 'GST RELIEF' in t:
        return 'GTR'
    if 'INTER-GATEWAY' in t or 'INTER GATEWAY' in t:
        return 'IGM'
    if 'SHUT' in t:
        return 'SHO'
    if t.startswith('GST'):
        return 'GST'
    return ''


def advise(sh, form=None, hsl=None, allowed=None):
    """-> sh['declaration'] (suggested type, reasons, needs, the decision APPLIED to the permit) and Declaration notes
    added to sh['issues_all']. allowed = [(text, code)] declaration types of the upload template (List-DeclarationType).

    The advisor DECIDES: who chose the declaration type -
      you   (job: Declaration type on the Header details / permit type picked on the job)  -> kept, only checked;
      customs instructions / documents (permit type read from them)                       -> kept, unless the rules say
            it cannot be right (dutiable goods under IPT GST, Place of Receipt of another type) -> changed automatically;
      nobody                                                                               -> the advisor's type.
    The applied type fills the Header (Declaration Type + code, Release / Receipt code when blank) and the XML message
    type; a value typed on the sheet tabs is written over it (always editable)."""
    form = form or {}
    hd = form.get('header') or {}
    cg = sh.get('cargo') or {}
    exp = sh.get('direction') == 'EXPORT'
    fpermit = (form.get('permit_type') or '').upper()[:2]
    permit = (fpermit or cg.get('permit_type') or '').upper()[:2]
    chosen = PERMIT_DECL.get(permit)
    typed = _decl_from_text(hd.get('declaration_type'))
    src = ('you' if fpermit and form.get('permit_src') != 'documents' else 'customs instructions') if chosen else ''
    if typed:
        src = 'you'
        msg = 'OUT' if exp else ('IPT' if typed in ('GST', 'DNG', 'DUT') else 'INP' if typed not in ('REM', 'BRE', 'TTF', 'TTI', 'IGM') else 'TNP')
        chosen = (msg, typed)
    release = (hd.get('release_code') or cg.get('release') or '').strip().upper()
    receipt = (hd.get('receipt_code') or cg.get('receipt') or ('ME' if permit == 'ME' else '')).strip().upper()
    items = sh.get('items') or []
    dut, liq, ctl, nondut = [], [], [], []
    for it in items:
        hs = re.sub(r'\D', '', str(it.get('hs') or ''))[:8]
        if not hs:
            continue
        inf = _hs_info(hsl, hs)
        label = '%s (%s)' % (it.get('item') or (it.get('desc') or '')[:25], hs)
        if str(inf.get('duty') or '').strip():
            dut.append(label)
        else:
            nondut.append(label)
        if hs[:4] in LIQUOR_TOBACCO:
            liq.append(label)
        if str(inf.get('ctl_exp' if exp else 'ctl_imp') or '').strip().upper() == 'CONTROL':
            ctl.append(label)
    why, notes = [], []
    # ---- the declaration type the documents point to
    if exp:
        sug = ('OUT', 'DRT')
        why.append('export of local / GST-paid goods or goods stored in an FTZ')
        if release and release not in ('O',) and release not in FTZ and re.match(r'[A-Z0-9]{5,}', release):
            sug = ('OUT', 'APS')
            why.append('Place of Release %s looks like a warehouse licence no.' % release)
        if receipt == 'SY':
            sug = ('OUT', 'APS')
            why.append('Place of Receipt SY = supply to a vessel at a shipyard')
    else:
        if receipt in RECEIPT_IMPLIES and RECEIPT_IMPLIES[receipt][1]:
            sug = RECEIPT_IMPLIES[receipt]
            why.append('Place of Receipt %s is used for %s %s' % (receipt, sug[0], sug[1]))
        elif receipt in FTZ:
            sug = ('INP', 'SFZ')
            why.append('goods received into %s (FTZ storage, no GST now)' % FTZ[receipt])
        elif dut:
            sug = ('IPT', 'DNG')
            why.append('%d dutiable item(s) - duty and GST payable' % len(dut))
        else:
            sug = ('IPT', 'GST')
            why.append('non-dutiable goods for local use - GST payable')
    if chosen and chosen != sug and not (chosen[1] == sug[1]):
        notes.append(('CHECK', 'Declaration: the documents point to %s %s (%s) - the permit is %s %s. %s' % (
            sug[0], sug[1], DECL_NAME.get(sug[1], ''), chosen[0], chosen[1], '; '.join(why))))
    use = chosen or sug
    # ---- the decision applied to the permit
    applied, auto_why = use, ''
    if not chosen:
        auto_why = 'no permit type chosen - the advisor\'s type'
    elif src != 'you':
        hard = None
        if receipt in RECEIPT_IMPLIES and RECEIPT_IMPLIES[receipt][1] and RECEIPT_IMPLIES[receipt] != chosen:
            hard = (RECEIPT_IMPLIES[receipt], 'Place of Receipt %s belongs to it' % receipt)
        elif chosen == ('IPT', 'GST') and dut and receipt not in ('ETHYL', 'VEHRAC') and not receipt.startswith('IEF'):
            hard = (('IPT', 'DNG'), '%d dutiable item(s) cannot go under IPT GST' % len(dut))
        if hard:
            applied, auto_why = hard[0], hard[1]
    codes = {c: t for t, c in (allowed or [])}
    fill = {}
    if allowed is not None and applied[1] not in codes:
        notes.append(('ERROR', 'Declaration: %s %s (%s) is needed but the upload template has no such declaration type '
                               '(%s) - declare it in TradeNet directly or choose another type on the Header tab'
                      % (applied[0], applied[1], DECL_NAME.get(applied[1], ''), ', '.join(sorted(codes)))))
        applied_ok = False
    else:
        applied_ok = True
        if src != 'you' or not typed:
            fill['Declaration Type'] = codes.get(applied[1], DECL_NAME.get(applied[1], ''))
            fill['Declaration Type Code'] = applied[1]
    if not release and applied[0] == 'OUT' and applied[1] in ('DRT', 'TCE', 'TCS', 'TCR', 'TCO', 'TCI'):
        fill['Release Location Code'] = 'O'
    if not receipt:
        r_ = {('INP', 'DES'): 'DUMP', ('INP', 'APS'): 'ME' if permit == 'ME' else ''}.get(applied, '')
        if r_:
            fill['Receipt Location Code'] = r_
    if applied != use and applied_ok:
        notes.append(('CHECK', 'Declaration (auto): %s %s from the %s changed to %s %s - %s. Change it on the Header tab if needed'
                      % (use[0], use[1], src, applied[0], applied[1], auto_why)))
    asked, use = use, (applied if applied_ok else use)
    # ---- rules of the declaration type in use
    if receipt in RECEIPT_IMPLIES and RECEIPT_IMPLIES[receipt][1] and chosen and RECEIPT_IMPLIES[receipt] != use:
        notes.append(('MISMATCH', 'Declaration: Place of Receipt %s belongs to %s %s, the permit is %s %s'
                      % (receipt, RECEIPT_IMPLIES[receipt][0], RECEIPT_IMPLIES[receipt][1], use[0], use[1])))
    if use == ('IPT', 'GST') and dut and receipt not in ('ETHYL', 'VEHRAC') and not receipt.startswith('IEF'):
        notes.append(('MISMATCH', 'Declaration: IPT GST is for non-dutiable goods (or duty-exempt with ETHYL / IEFxxx / VEHRAC) - '
                                  'dutiable item(s): %s -> IPT DNG (Duty & GST) or bond them (INP APS licensed warehouse)'
                      % ', '.join(dut[:8])))
    if use == ('INP', 'APS'):
        if dut and nondut:
            notes.append(('ERROR', 'Declaration: INP APS - dutiable goods (licensed warehouse) and non-dutiable goods (bonded / '
                                   'ZG warehouse) must not be in the same permit - split it. Dutiable: %s' % ', '.join(dut[:8])))
        if dut:
            notes.append(('CHECK', 'Declaration: INP APS into a licensed warehouse - lot no. and product codes needed for the dutiable items'))
    if use == ('INP', 'SFZ'):
        if liq:
            notes.append(('ERROR', 'Declaration: dutiable liquor / tobacco cannot be stored in an FTZ (INP SFZ): %s' % ', '.join(liq[:8])))
        if receipt and receipt not in FTZ:
            notes.append(('CHECK', 'Declaration: INP SFZ - Place of Receipt should be the FTZ (KZ / PPZ / JZ / CZ ...), not %s' % receipt))
        if not receipt:
            notes.append(('MISSING', 'Declaration: INP SFZ - Place of Receipt (the FTZ, e.g. KZ) is blank'))
        if cg.get('mode') == 'Road' and release and release != 'THQ':
            notes.append(('CHECK', 'Declaration: by road via Tuas checkpoint the Place of Release is usually THQ (now %s)' % release))
    if use[1] == 'GTR':
        if receipt not in ('EM', 'INTORG', 'VEHSG', 'TRADESP', 'RELIEF'):
            notes.append(('MISSING', 'Declaration: INP GTR needs Place of Receipt EM / INTORG / VEHSG / TRADESP / RELIEF'))
        if receipt in ('TRADESP', 'RELIEF'):
            tot = sh.get('invoice_total')
            if (sh.get('currency') or '').upper() == 'SGD' and tot and tot > 400:
                notes.append(('ERROR', 'Declaration: %s relief only for a total value up to S$400 (invoice SGD %.2f)' % (receipt, tot)))
            if receipt == 'RELIEF' and cg.get('mode') and cg.get('mode') != 'Air':
                notes.append(('ERROR', 'Declaration: RELIEF (<= S$400) is for imports by air only'))
            notes.append(('CHECK', 'Declaration: controlled samples / imports - approval of the Competent Authority before import'))
    if use[1] == 'DES' and receipt != 'DUMP':
        notes.append(('MISSING', 'Declaration: INP DES needs Place of Receipt DUMP'))
    if use[1] in ('TCE', 'TCS', 'TCR', 'TCO', 'SHO') or (use[0] == 'OUT' and use[1].startswith('TC')):
        notes.append(('CHECK', 'Declaration: %s %s needs the Previous Permit No. / end date - type it on the Header tab' % use))
    if use == ('OUT', 'DRT'):
        if not release and 'Release Location Code' not in fill:
            notes.append(('CHECK', 'Declaration: OUT DRT - local goods: Place of Release = O'))
        if release == 'O' and receipt and receipt not in FTZ and not re.match(r'^(TQ|WQ|THQ|WHQ|CAZ|PSA)', receipt):
            notes.append(('CHECK', 'Declaration: OUT DRT - Place of Receipt should be an FTZ or approved exit point (now %s)' % receipt))
        if not receipt:
            notes.append(('MISSING', 'Declaration: OUT DRT - Place of Receipt (FTZ / exit point) is blank'))
    if liq and use[1] == 'REX':
        notes.append(('CHECK', 'Declaration: liquor / tobacco re-export - break bulk in the first FTZ, one INP REX per re-export shipment'))
    if ctl and not any('controlled line' in str(x[1]) for x in (sh.get('issues_all') or [])):
        notes.append(('CHECK', 'Declaration: %d controlled item(s) for %s - CA licence / product code needed: %s'
                      % (len(ctl), 'export' if exp else 'import', ', '.join(ctl[:6]))))
    sh['declaration'] = dict(applied='%s %s' % applied, applied_name=DECL_NAME.get(applied[1], ''), source=src or 'advisor',
                             auto=src != 'you', auto_why=auto_why, changed=applied != use and applied_ok,
                             allowed=[list(x) for x in (allowed or [])], fill=fill, msg_type=applied[0] + 'DEC' if applied_ok else '',
                             permit_applied=DECL_PERMIT.get(applied, ''), applied_ok=applied_ok,
                             message=applied[0], type=applied[1], name=DECL_NAME.get(applied[1], ''), message_name=MSG_NAME.get(applied[0], ''),
                             suggested='%s %s' % sug, suggested_name=DECL_NAME.get(sug[1], ''), chosen='%s %s' % chosen if chosen else '',
                             permit_code=DECL_PERMIT.get(sug, ''), why=why, needs=NEEDS.get(applied, []),
                             release=release, receipt=receipt, dutiable=len(dut), liquor_tobacco=len(liq), controlled=len(ctl))
    iss = [x for x in (sh.get('issues_all') or []) if not str(x[1]).startswith('Declaration:')]
    sh['issues_all'] = iss + notes
    return sh['declaration']
