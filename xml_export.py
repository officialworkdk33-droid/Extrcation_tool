# -*- coding: utf-8 -*-
"""dec-upload Excel  ->  FrtImport CUSTOMS PERMIT XML (same layout as the user's sample FrtImport_SD2026071600468.xml).

The XML is built FROM the generated (or any chosen) dec-upload Excel, so XML and Excel always carry the same values.
Element order / nesting comes from frtimport_skeleton.xml (made from the sample, packed into the exe).
Output: UTF-8, CRLF, TAB indentation, empty elements written as <Tag></Tag> - exactly like the sample."""
import os, sys, re, datetime, zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

DEFAULT_OPTS = dict(snd='ITG_INDIA', rcp='YCH', version='4.1', msg_import='IPTDEC', msg_export='',
                    charges=False, code_frc='FRC', code_ins='INS', code_otc='OTC', item_inv_dash=True)


def _res(name):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def skeleton():
    """XML layout: frtimport_skeleton.xml next to the exe (lets you change the layout) > inside the exe > built in."""
    for p in (os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)),
                           'frtimport_skeleton.xml'), _res('frtimport_skeleton.xml')):
        if os.path.exists(p):
            try:
                return ET.parse(p).getroot()
            except ET.ParseError:
                pass
    return ET.fromstring(SKELETON)


# layout of the user's sample FrtImport_SD2026071600468.xml (one Item, values cleared) - built in, no file needed
SKELETON = r'''<FrtImport>
	<Comm>
		<DocType>CUSTOMS PERMIT</DocType>
		<Version>4.1</Version>
		<SndOrgID>ITG_INDIA</SndOrgID>
		<RcpOrgID>YCH</RcpOrgID>
		<DateTime />
	</Comm>
	<ShipmentHeader>
		<ShptJobNo />
		<ShptJobDate />
		<ShptXRef />
		<ShptRef />
		<ShptMsgType />
		<ShptDeclType />
		<ShptPackingType />
		<ShptBGInd />
		<ShptMAWB />
		<ShptHAWB />
		<ShptPlaceOfRelease />
		<ShptLocForRelease />
		<ShptPlaceOfReceipt />
		<ShptLocForReceipt />
		<ShptImpExpCRNo />
		<ShptImpExpName1 />
		<ShptImpExpName2 />
		<ShptInwardCarrCRNo />
		<ShptInwardCarrName />
		<ShptExpCRNo />
		<ShptExpName1 />
		<ShptExpName2 />
		<ShptExpStreetName />
		<ShptExpPostBoxNo />
		<ShptExpCityName />
		<ShptExpSubDivCode />
		<ShptExpSubDivName />
		<ShptExpPostalCode />
		<ShptExpCtryCode />
		<ShptOutwardCarrCRNo />
		<ShptOutwardCarrName />
		<ShptCsgnCRNo />
		<ShptCsgnName1 />
		<ShptCsgnName2 />
		<ShptCsgnStreetName />
		<ShptCsgnPostBoxNo />
		<ShptCsgnCityName />
		<ShptCsgnSubDivCode />
		<ShptCsgnSubDivName />
		<ShptCsgnPostalCode />
		<ShptCsgnCtryCode />
		<ShptCgoAgtCRNo />
		<ShptCgoAgtName />
		<ShptEndUserCRNo />
		<ShptEndUserName />
		<ShptEndUserStreetName />
		<ShptEndUserPostBoxNo />
		<ShptEndUserCityName />
		<ShptEndUserSubDivCode />
		<ShptEndUserSubDivName />
		<ShptEndUserPostalCode />
		<ShptEndUserCtryCode />
		<ShptEndUserCtryName />
		<ShptHandlingAgtCRNo />
		<ShptHandlingAgtName />
		<ShptClaimCRNo />
		<ShptClaimCmp />
		<ShptClaimName />
		<ShptClaimNRIC />
		<ShptClaimNation />
		<ShptClaimTel />
		<ShptMfgName />
		<ShptMfgStreetName />
		<ShptMfgPostBoxNo />
		<ShptMfgCityName />
		<ShptMfgSubDivCode />
		<ShptMfgSubDivName />
		<ShptMfgPostalCode />
		<ShptMfgCtryCode />
		<ShptTptMode />
		<ShptArrDeptDate />
		<ShptStartDate />
		<ShptFltVsl />
		<ShptVoy />
		<ShptPOL />
		<ShptOutTptMode />
		<ShptPOD />
		<ShptCtryFinalDest />
		<ShptVslType />
		<ShptStartRemoveDate />
		<ShptStartImpDate />
		<ShptEndImpDate />
		<ShptOutFltVsl />
		<ShptOutVoy />
		<ShptStartDeptDate />
		<ShptOutMAWB />
		<ShptOutHAWB />
		<ShptTtlDuty1 />
		<ShptTtlDuty2 />
		<ShptTtlOtherTax />
		<ShptTtlGSTPayable />
		<ShptTtItem />
		<ShptDeclValue />
		<ShptTtlGSTPaid />
		<ShptTtlWgt />
		<ShptTtlWgtUnit />
		<ShptTtlPcs />
		<ShptTtlPcsUnit />
		<ShptTtlDutyQty />
		<ShptTtlDutyQtyUnit />
		<ShptRemark />
		<ShptReadySubmit />
		<ShptPrevPmtNo />
		<ShptOUTCOLink />
		<ShptApplPrdCode />
		<ShptCertType1 />
		<ShptCertType2 />
		<ShptEntryYear />
		<ShptGSPDonorCountry />
		<ShptPreferenceContentPercent />
		<ShptCurrencyCode />
		<ShptAddCertificateDet1 />
		<ShptAddCertificateDet2 />
		<ShptAddCertificateDet3 />
		<ShptAddCertificateDet4 />
		<ShptAddCertificateDet5 />
		<ShptTransportDet1 />
		<ShptTransportDet2 />
		<ShptTransportDet3 />
		<ShptTransportDet4 />
		<ShptTransportDet5 />
		<ShptPlaceOfStorage />
		<VccSlac />
		<CPCCode>
			<CPCCode />
			<ProcessingCode>
				<ProcessCode1 />
				<ProcessCode2 />
				<ProcessCode3 />
			</ProcessingCode>
		</CPCCode>
		<Invoice>
			<InvoiceNo />
			<InvoiceDate />
			<TermType />
			<SupplierName />
			<ManuName />
			<InvoiceFreightAndOtherCharges>
				<ChargeCode />
				<Amount />
				<PercentTtlQty />
				<Rate />
				<FromCurrCode />
			</InvoiceFreightAndOtherCharges>
		</Invoice>
		<Item>
			<ItemHsCode />
			<ItemDesc />
			<ItemCtryOrigin />
			<ItemQty />
			<ItemUnit />
			<ItemWgt />
			<ItemWgtUnit />
			<ItemFOB />
			<ItemMark />
			<ItemBrand />
			<ItemModel />
			<ItemInvoiceTerm />
			<ItemValue />
			<ItemGST />
			<ItemFee />
			<ItemInvoiceNo />
			<ItemInvoiceDate />
			<ItemRateType />
			<ItemDutyInd />
			<ItemSDNPInd />
			<ItemDutyRate1 />
			<ItemDutyUnit1 />
			<ItemDutyAmt1 />
			<ItemDutyRate2 />
			<ItemDutyUnit2 />
			<ItemDutyAmt2 />
			<ItemOtherTaxRate />
			<ItemOtherTaxUnit />
			<ItemOtherTaxAmt />
			<ItemDutyQty />
			<ItemDutyQtyUnit />
			<ItemTtlDutyQty />
			<ItemTtlDutyQtyUnit />
			<ItemPrefRate />
			<ItemCatCode />
			<ItemEndUseCode1 />
			<ItemEndUseCode2 />
			<ItemEndUseCode3 />
			<ItemEndUseDesc />
			<ItemUnitPrice />
			<ItemUnitPriceCurrencyCode />
			<ItemInHAWB />
			<ItemOutHAWB />
			<ItemProdCode1 />
			<ItemProdQty1 />
			<ItemProdUnit1 />
			<ItemProdCode2 />
			<ItemProdQty2 />
			<ItemProdUnit2 />
			<ItemProdCode3 />
			<ItemProdQty3 />
			<ItemProdUnit3 />
			<ItemProdCode4 />
			<ItemProdQty4 />
			<ItemProdUnit4 />
			<ItemProdCode5 />
			<ItemProdQty5 />
			<ItemProdUnit5 />
			<ItemCostLUDate />
			<ItemCooCertQty />
			<ItemCooCertQtyUnit />
			<ItemCooCertValue />
			<ItemCooMfrCostDate />
			<ItemCooTxtCode />
			<ItemCooTxtQuotQty />
			<ItemCooQtyUnit />
			<ItemCooInvoiceNo />
			<ItemCooInvoiceDate />
			<ItemCooOrigCritCond1 />
			<ItemCooOrigCritCond2 />
			<ItemCooOrigCritCond3 />
			<ItemCooCertHsCode />
			<ItemCooPercentContent />
			<ItemCooCertDesc />
			<ItemAdditionalInfo>
				<ItemInHseCode />
				<ItemFOBForeignAmt />
				<ItemFOBForeignCurr />
			</ItemAdditionalInfo>
			<ItemPackingAndVehicleInfo>
				<CurrLotNo />
				<PrevLotNo />
				<VolPercent />
				<VolUnit />
				<OuterPack />
				<OuterUnit />
				<InPack />
				<InUnit />
				<InnerPack />
				<InnerUnit />
				<InMostPack />
				<InMostUnit />
				<CalDutyInd />
				<CalQtyInd />
				<NoOfStick />
				<TestWt />
				<ProdCode />
				<EngineCapacity />
				<EnginePower />
				<OrigRegDate />
				<SingaporeRegDate />
			</ItemPackingAndVehicleInfo>
			<ItemVehicleEngineChasis>
				<VehicleEngineNo />
				<VehicleChassisNo />
				<VehicleType />
			</ItemVehicleEngineChasis>
			<ItemCascInfo>
				<CascCode1 />
				<CascCode2 />
				<CascCode3 />
				<CascEndUseDesc />
			</ItemCascInfo>
		</Item>
		</ShipmentHeader>
</FrtImport>'''



def _num(v, dec=4):
    """226.0 -> 226, 71024.370 -> 71024.37 (no thousands separator, no trailing zeros)"""
    if v in (None, ''):
        return ''
    try:
        f = float(str(v).replace(',', ''))
    except ValueError:
        return str(v)
    s = ('%.' + str(dec) + 'f') % f
    return s.rstrip('0').rstrip('.') if '.' in s else s


def _date(v, fmt):
    """Excel date / 'dd/mm/yyyy' -> fmt ('%d%m%Y' or '%d/%m/%Y')"""
    if v in (None, ''):
        return ''
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime(fmt)
    if isinstance(v, (int, float)) and 20000 < float(v) < 80000:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime(fmt)
    m = re.match(r'(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})', str(v))
    if m:
        return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1))).strftime(fmt)
    return str(v)


def read_dec_excel(path):
    """dec-upload Excel -> (header dict, invoice dict, [item dicts]) using the column titles of row 1."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    def rows(name):
        if name not in wb.sheetnames:
            return []
        ws = wb[name]
        it = ws.iter_rows(values_only=True)
        head = [str(h).strip() if h is not None else '' for h in next(it, [])]
        out = []
        for r in it:
            if not any(v not in (None, '') for v in r):
                continue
            out.append({h: v for h, v in zip(head, r) if h})
        return out
    h = (rows('Header') or [{}])[0]
    invs = rows('Invoice') or [{}]
    inv = dict(invs[0])
    inv['_all'] = invs                     # several currencies = one invoice row each
    items = rows('Item')
    conts = rows('Container')
    return h, inv, items, conts


def _split_name(name, n=35):
    name = (name or '').strip()
    if len(name) <= n:
        return name, ''
    cut = name.rfind(' ', 0, n)
    cut = cut if cut > 10 else n
    return name[:cut].strip(), name[cut:].strip()[:n]


WEIGHT_UOMS = ('KGM', 'TNE')


def build_values(h, inv, items, job_no='', opts=None):
    """Header / invoice / item values of one permit, keyed by XML element name."""
    o = dict(DEFAULT_OPTS, **(opts or {}))
    g = lambda k: h.get(k) if h.get(k) not in (None, '') else ''
    exp = bool(g('Exporter') or g('Exporter UEN')) and not (g('Importer') or g('Importer Uen'))
    today = datetime.date.today()
    party = g('Exporter') if exp else g('Importer')
    n1, n2 = _split_name(str(party))
    hawb = next((str(it.get('Out HAWB' if exp else 'In HAWB') or '') for it in items if it.get('Out HAWB' if exp else 'In HAWB')), '')
    tmode = str(g('Outward Transport Mode' if exp else 'Inward Transport Mode'))
    sea = tmode == '1'
    flight = str(g('Out Flight Number' if exp else 'In Flight Number'))
    vessel = str(g('Out Transport Identifier' if exp else 'In Transport Identifier'))
    rem = re.sub(r'\s*\n\s*', ' ', str(g("Trader's Remarks"))).strip()
    v = {
        'SndOrgID': o['snd'], 'RcpOrgID': o['rcp'], 'Version': o['version'], 'DocType': 'CUSTOMS PERMIT',
        'ShptJobNo': '', 'ShptJobDate': '',            # left blank (user rule) - the file name still uses the job no
        'ShptMsgType': o['msg_export'] if exp else o['msg_import'],
        'ShptDeclType': g('Declaration Type Code'), 'ShptPackingType': g('Cargo Pack Type Code'),
        'ShptBGInd': g('BG Indicator'),
        'ShptPlaceOfRelease': g('Release Location Code'), 'ShptLocForRelease': g('Release Location Name'),
        'ShptPlaceOfReceipt': g('Receipt Location Code'), 'ShptLocForReceipt': g('Receipt Location Name'),
        'ShptImpExpCRNo': g('Exporter UEN') if exp else g('Importer Uen'), 'ShptImpExpName1': n1, 'ShptImpExpName2': n2,
        'ShptCsgnName1': g('Consignee Name'), 'ShptCsgnStreetName': g('Consignee Address 1'),
        'ShptCsgnCtryCode': g('Consignee Address Country Code'),
        'ShptHandlingAgtCRNo': g('Handling Agent Uen'), 'ShptHandlingAgtName': g('Handling Agent Name'),
        'ShptPOL': g('Loading Port Code'), 'ShptPOD': g('Discharge Port Code'),
        'ShptVslType': g('Vessel Type') if sea else '',
        'ShptTtlWgt': _num(g('Total Gross Weight'), 3), 'ShptTtlWgtUnit': g('Total Gross Weight Uom'),
        'ShptTtlPcs': _num(g('Total Outer Pack Qty')), 'ShptTtlPcsUnit': g('Total Outer Pack Uom'),
        'ShptRemark': rem, 'ShptPrevPmtNo': g('Previous Permit No.'), 'ShptCurrencyCode': g('Currency Code'),
        'ShptPlaceOfStorage': g('Storage Location Code'),
    }
    if exp:
        v.update({'ShptOutwardCarrCRNo': g('Outward Carrier Agent Uen'), 'ShptOutwardCarrName': g('Outward Carrier Agent Name'),
                  'ShptOutTptMode': tmode, 'ShptStartDeptDate': _date(g('Departure Date'), '%d/%m/%Y'),
                  'ShptOutFltVsl': vessel if sea else flight, 'ShptOutVoy': flight if sea else '',
                  'ShptOutMAWB': g('Outward MAWB'), 'ShptOutHAWB': hawb,
                  'ShptCtryFinalDest': g('Final Destination Country Code')})
        e1, e2 = n1, n2
        v.update({'ShptExpCRNo': g('Exporter UEN'), 'ShptExpName1': e1, 'ShptExpName2': e2})
    else:
        v.update({'ShptInwardCarrCRNo': g('Inward Carrier Agent Uen'), 'ShptInwardCarrName': g('Inward Carrier Agent Name'),
                  'ShptTptMode': tmode, 'ShptArrDeptDate': _date(g('Arrival Date'), '%d/%m/%Y'),
                  'ShptFltVsl': vessel if sea else flight, 'ShptVoy': flight if sea else '',
                  'ShptMAWB': g('Inward MAWB'), 'ShptHAWB': hawb})
    # invoice
    gi = lambda k: inv.get(k) if inv.get(k) not in (None, '') else ''
    v['InvoiceNo'] = str(gi('Invoice Number') or '-')
    v['InvoiceDate'] = _date(gi('Invoice Date'), '%d%m%Y') or today.strftime('%d%m%Y')
    v['TermType'] = gi('Term Type')
    v['SupplierName'] = gi('Supplier/Manufacturer Name')
    charges = [('INV', gi('Total Invoice Amount'), '', gi('Total Invoice Currency'))]
    if o.get('charges'):
        for code, amt, pct, cur in ((o['code_frc'], gi('Freight Charge Amount'), gi('Freight Charge %'), gi('Freight Charge Currency')),
                                    (o['code_ins'], gi('Insurance Charge Amount'), gi('Insurance Charge %'), gi('Insurance Charge Currency')),
                                    (o['code_otc'], gi('Other Taxable Charge Amount'), gi('Other Taxable Charge %'),
                                     gi('Other Taxable Charge Currency'))):
            if amt not in ('', None) or pct not in ('', None):
                charges.append((code, amt, pct, cur))
    elif gi('Other Taxable Charge Amount') not in ('', None):     # charges kept apart from the items: always declared
        charges.append((o['code_otc'], gi('Other Taxable Charge Amount'), gi('Other Taxable Charge %'),
                        gi('Other Taxable Charge Currency') or gi('Total Invoice Currency')))
    v['_charges'] = charges
    # several invoices (one per currency): each invoice block followed by its items (linked by the invoice number)
    groups = []
    all_inv = inv.get('_all') or []
    for iv in (all_inv if len(all_inv) > 1 else []):
        g2 = lambda k: iv.get(k) if iv.get(k) not in (None, '') else ''
        ch = [('INV', g2('Total Invoice Amount'), '', g2('Total Invoice Currency'))]
        for code, amt, pct, cur in ((o['code_frc'], g2('Freight Charge Amount'), g2('Freight Charge %'), g2('Freight Charge Currency')),
                                    (o['code_ins'], g2('Insurance Charge Amount'), g2('Insurance Charge %'), g2('Insurance Charge Currency')),
                                    (o['code_otc'], g2('Other Taxable Charge Amount'), g2('Other Taxable Charge %'),
                                     g2('Other Taxable Charge Currency'))):
            if (o.get('charges') or code == o['code_otc']) and (amt not in ('', None) or pct not in ('', None)):
                ch.append((code, amt, pct, cur))
        groups.append(dict(InvoiceNo=str(g2('Invoice Number') or '-'),
                           InvoiceDate=_date(g2('Invoice Date'), '%d%m%Y') or today.strftime('%d%m%Y'),
                           TermType=g2('Term Type'), SupplierName=g2('Supplier/Manufacturer Name'), _charges=ch))
    # items
    its = []
    for it in items:
        gt = lambda k: it.get(k) if it.get(k) not in (None, '') else ''
        hs = gt('HS Code')
        hs = str(int(hs)).zfill(8) if isinstance(hs, (int, float)) else str(hs)
        x = {'ItemHsCode': hs, 'ItemDesc': gt('Description'), 'ItemCtryOrigin': gt('Country of Origin'),
             'ItemQty': _num(gt('HS Qty')), 'ItemUnit': gt('HS UOM'), 'ItemBrand': gt('Brand'), 'ItemModel': gt('Model'),
             'ItemValue': _num(gt('Total Line Amount'), 2),
             'ItemInvoiceNo': '-' if o.get('item_inv_dash', True) else str(gt('Invoice Number')),
             'ItemInvoiceDate': v['InvoiceDate'],
             'ItemUnitPrice': _num(gt('Unit Price')), 'ItemUnitPriceCurrencyCode': gt('Item Currency'),
             'ItemInHAWB': '' if exp else gt('In HAWB'), 'ItemOutHAWB': gt('Out HAWB') if exp else '',
             'ItemMark': gt('Shipping Marks 1'), 'CurrLotNo': gt('Current Lot'), 'PrevLotNo': gt('Previous Lot'),
             'OuterPack': _num(gt('Outer Pack Qty')), 'OuterUnit': gt('Outer Pack UOM'),
             # dutiable goods (liquor): alcohol %, dutiable qty (bottle size) / total dutiable qty
             'VolPercent': _num(gt('Alcohol Percentage')), 'ItemDutyQty': _num(gt('Dutiable Qty')),
             'ItemDutyQtyUnit': gt('Dutiable UOM'), 'ItemTtlDutyQty': _num(gt('Total Dutiable Qty')),
             'ItemTtlDutyQtyUnit': gt('Total Dutiable UOM')}
        if str(x['ItemUnit']).upper() in WEIGHT_UOMS:    # net weight (KGM) / net weight based (TNE): ItemWgt, not ItemQty
            x['ItemWgt'], x['ItemWgtUnit'] = x['ItemQty'], str(x['ItemUnit']).upper()
            x['ItemQty'], x['ItemUnit'] = '', ''
        for n in (1, 2):
            if gt('CASC Product Code %d' % n):
                x['ItemProdCode%d' % n] = gt('CASC Product Code %d' % n)
                x['ItemProdQty%d' % n] = _num(gt('CASC Product Code %d Qty' % n))
                x['ItemProdUnit%d' % n] = gt('CASC Product Code %d UOM' % n)
        x['_inv'] = str(gt('Invoice Number'))
        its.append(x)
    v['_items'] = its
    if groups:
        for g in groups:
            g['_items'] = [dict(x, ItemInvoiceNo=g['InvoiceNo'], ItemInvoiceDate=g['InvoiceDate']) for x in its if x['_inv'] == g['InvoiceNo']]
        left = [x for x in its if x['_inv'] not in {g['InvoiceNo'] for g in groups}]
        if left:
            groups[0]['_items'] += [dict(x, ItemInvoiceNo=groups[0]['InvoiceNo']) for x in left]
        v['_inv_groups'] = groups
    return v


def _write(e, vals, level, out):
    tab = '\t' * level
    if len(e) == 0:
        val = vals.get(e.tag, e.text if e.tag in ('DocType',) else '')
        out.append('%s<%s>%s</%s>' % (tab, e.tag, escape(str(val)) if val not in (None, '') else '', e.tag))
        return
    out.append('%s<%s>' % (tab, e.tag))
    for c in e:
        if c.tag == 'Item':
            continue                                   # items written after the invoice (below)
        if c.tag == 'Invoice':
            for g in vals.get('_inv_groups') or [dict(_items=vals['_items'])]:
                _write_invoice(c, dict(vals, **g), level + 1, out)
                for x in g['_items']:
                    _write(c_item(e), x, level + 1, out)
            continue
        _write(c, vals, level + 1, out)
    out.append('%s</%s>' % (tab, e.tag))


def c_item(parent):
    return parent.find('Item')


def _write_invoice(e, vals, level, out):
    tab = '\t' * level
    out.append('%s<%s>' % (tab, e.tag))
    for c in e:
        if c.tag == 'InvoiceFreightAndOtherCharges':
            for code, amt, pct, cur in vals['_charges']:
                cv = {'ChargeCode': code, 'Amount': _num(amt, 2), 'PercentTtlQty': _num(pct), 'FromCurrCode': cur}
                _write(c, cv, level + 1, out)
        else:
            _write(c, vals, level + 1, out)
    out.append('%s</%s>' % (tab, e.tag))


def excel_to_xml(xlsx, xml_path=None, job_no='', opts=None):
    """One dec-upload Excel -> FrtImport XML. Returns the XML path."""
    h, inv, items, conts = read_dec_excel(xlsx)
    stem = os.path.splitext(os.path.basename(xlsx))[0]
    job = job_no or stem
    vals = build_values(h, inv, items, job, opts)
    out = ['<?xml version="1.0" encoding="UTF-8"?>']
    _write(skeleton(), vals, 0, out)
    xml_path = xml_path or os.path.join(os.path.dirname(xlsx), 'FrtImport_%s.xml' % re.sub(r'[^\w.-]+', '_', job))
    with open(xml_path, 'w', encoding='utf-8', newline='') as f:
        f.write('\r\n'.join(out))
    ET.parse(xml_path)                                 # must be well-formed
    return xml_path


def check_values(xlsx):
    """Things the XML needs that are blank in the Excel (shown to the user)."""
    h, inv, items, _ = read_dec_excel(xlsx)
    need = []
    exp = bool(h.get('Exporter') or h.get('Exporter UEN')) and not (h.get('Importer') or h.get('Importer Uen'))
    for k in (('Exporter UEN', 'Exporter') if exp else ('Importer Uen', 'Importer')) + (
            'Declaration Type Code', 'Cargo Pack Type Code', 'Release Location Code', 'Receipt Location Code',
            'Total Gross Weight', 'Total Outer Pack Qty'):
        if h.get(k) in (None, ''):
            need.append(k)
    bad = [i + 2 for i, it in enumerate(items) if not it.get('HS Code') or it.get('HS Qty') in (None, '') or not it.get('HS UOM')]
    if bad:
        need.append('HS Code / HS Qty / HS UOM on item row(s) %s' % ', '.join(map(str, bad[:15])))
    return need


def zip_xmls(paths, zpath):
    with zipfile.ZipFile(zpath, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, os.path.basename(p))
    return zpath
