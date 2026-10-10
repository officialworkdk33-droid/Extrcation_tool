# -*- coding: utf-8 -*-
"""UEN (CR UEI) of a Singapore company from ACRA's open data on data.gov.sg ("ACRA Information on Corporate Entities").

bizfile.gov.sg itself has a reCAPTCHA (no automatic look-ups); the same register is published as open data, one
dataset per first letter of the entity name. Look-ups go to https://data.gov.sg (needs internet on the server PC),
answers are cached in PERMIT_SERVER_DATA/acra_cache.json. Nothing found / no internet -> '' (the user types it).
"""
import os, re, json, time, threading, urllib.parse, urllib.request

API = 'https://data.gov.sg/api/action/datastore_search'
DATASETS = {   # first letter of the entity name -> dataset id (collection 2 on data.gov.sg)
    'A': 'd_8575e84912df3c28995b8e6e0e05205a', 'B': 'd_3a3807c023c61ddfba947dc069eb53f2',
    'C': 'd_c0650f23e94c42e7a20921f4c5b75c24', 'D': 'd_acbc938ec77af18f94cecc4a7c9ec720',
    'E': 'd_124a9bd407c7a25f8335b93b86e50fdd', 'F': 'd_4526d47d6714d3b052eed4a30b8b1ed6',
    'G': 'd_b58303c68e9cf0d2ae93b73ffdbfbfa1', 'H': 'd_fa2ed456cf2b8597bb7e064b08fc3c7c',
    'I': 'd_85518d970b8178975850457f60f1e738', 'J': 'd_478f45a9c541cbe679ca55d1cd2b970b',
    'K': 'd_5573b0db0575db32190a2ad27919a7aa', 'L': 'd_a2141adf93ec2a3c2ec2837b78d6d46e',
    'M': 'd_9af9317c646a1c881bb5591c91817cc6', 'N': 'd_67e99e6eabc4aad9b5d48663b579746a',
    'O': 'd_5c4ef48b025fdfbc80056401f06e3df9', 'P': 'd_181005ca270b45408b4cdfc954980ca2',
    'Q': 'd_4130f1d9d365d9f1633536e959f62bb7', 'R': 'd_2b8c54b2a490d2fa36b925289e5d9572',
    'S': 'd_df7d2d661c0c11a7c367c9ee4bf896c1', 'T': 'd_72f37e5c5d192951ddc5513c2b134482',
    'U': 'd_0cc5f52a1f298b916f317800251057f3', 'V': 'd_e97e8e7fc55b85a38babf66b0fa46b73',
    'W': 'd_af2042c77ffaf0db5d75561ce9ef5688', 'X': 'd_1cd970d8351b42be4a308d628a6dd9d3',
    'Y': 'd_31af23fdb79119ed185c256f03cb5773', 'Z': 'd_4e3db8955fdcda6f9944097bef3d2724',
    'OTHERS': 'd_300ddc8da4e8f7bdc1bfc62d0d99e2e7'}
LIVE = re.compile(r'LIVE|REGISTERED|EXISTING', re.I)
_lock = threading.Lock()
STATE = dict(cache_path='', cache=None, last_error='')


def norm(name):
    """'RS Components Pte. Ltd.' -> 'RS COMPONENTS PTE LTD' (for comparing names)."""
    n = re.sub(r'[^A-Z0-9&()]+', ' ', str(name or '').upper())
    n = re.sub(r'\bPRIVATE\b', 'PTE', n)
    n = re.sub(r'\bLIMITED\b', 'LTD', n)
    n = re.sub(r'\bCOMPANY\b', 'CO', n)
    n = re.sub(r'\(\s*', '(', re.sub(r'\s*\)', ')', n))
    return re.sub(r'\s+', ' ', n).strip()


def _cache():
    if STATE['cache'] is None:
        try:
            with open(STATE['cache_path'], encoding='utf-8') as f:
                STATE['cache'] = json.load(f)
        except Exception:
            STATE['cache'] = {}
    return STATE['cache']


def _save():
    if STATE['cache_path']:
        try:
            with open(STATE['cache_path'], 'w', encoding='utf-8') as f:
                json.dump(STATE['cache'], f, indent=1, sort_keys=True)
        except OSError:
            pass


def _get(params, timeout=12):
    url = API + '?' + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={'User-Agent': 'PermitServer/1.0'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        d = json.loads(r.read().decode('utf-8'))
    if not d.get('success', True):
        raise RuntimeError(str(d.get('error') or d)[:200])
    return (d.get('result') or {}).get('records') or []


def search(name, limit=20, timeout=12):
    """Companies whose name contains the words of `name`: [(uen, entity name, status)] - live ones first."""
    n = norm(name)
    if not n:
        return []
    first = n[0] if n[0].isalpha() else 'OTHERS'
    ds = DATASETS.get(first, DATASETS['OTHERS'])
    words = [w for w in re.findall(r'[A-Z0-9&]+', n) if w not in ('PTE', 'LTD', 'CO', 'THE', 'AND', '&')]
    recs = []
    for q in (json.dumps({'entity_name': name.upper()}), ' '.join(words) or n):
        try:
            recs = _get(dict(resource_id=ds, q=q, limit=limit, fields='uen,entity_name,entity_status_description'), timeout)
        except Exception as e:
            STATE['last_error'] = str(e)[:200]
            raise
        if recs:
            break
    out = []
    for r in recs:
        if r.get('uen') and r.get('entity_name'):
            out.append((str(r['uen']).strip().upper(), str(r['entity_name']).strip(), str(r.get('entity_status_description') or '')))
    out.sort(key=lambda x: (norm(x[1]) != n, not LIVE.search(x[2]), len(x[1])))
    return out


def lookup(name, timeout=12):
    """UEN of exactly this company (same name, live) - '' when not sure. Cached (also 'not found', for 7 days)."""
    n = norm(name)
    if not n or len(n) < 4:
        return ''
    with _lock:
        c = _cache().get(n)
    if c and (c.get('uen') or time.time() - c.get('t', 0) < 7 * 86400):
        return c.get('uen', '')
    try:
        hits = search(name, timeout=timeout)
    except Exception:
        return ''
    exact = [h for h in hits if norm(h[1]) == n and LIVE.search(h[2])] or [h for h in hits if norm(h[1]) == n]
    uen = exact[0][0] if len({h[0] for h in exact}) == 1 or (exact and LIVE.search(exact[0][2])) else ''
    with _lock:
        _cache()[n] = dict(uen=uen, name=exact[0][1] if exact else '', t=time.time())
        _save()
    return uen


def _address(r):
    """Registered address of an ACRA record (block, street, #level-unit, building, SINGAPORE postal code)."""
    unit = ''
    if r.get('level_no') or r.get('unit_no'):
        unit = '#%s-%s' % (str(r.get('level_no') or '').strip(), str(r.get('unit_no') or '').strip())
    a1 = ' '.join(x for x in (str(r.get('block') or '').strip(), str(r.get('street_name') or '').strip()) if x and x.lower() != 'na')
    a2 = ' '.join(x for x in (unit, str(r.get('building_name') or '').strip()) if x and x.lower() != 'na')
    pc = str(r.get('postal_code') or '').strip()
    lines = [x for x in (a1, a2, ('SINGAPORE ' + pc) if pc else '') if x]
    return [x.upper() for x in lines]


def by_uen(uen, hint_name='', timeout=12):
    """Registered name (+ address) of a UEN from the ACRA open data -> dict(name, status, address[]) or {}.
    Cached for 30 days (also 'not found' for 7 days). The dataset of the hint name's first letter is tried first."""
    uen = re.sub(r'\s', '', str(uen or '')).upper()
    if not re.fullmatch(r'\d{8,9}[A-Z]|[STR]\d{2}[A-Z]{2}\d{4}[A-Z]', uen):
        return {}
    key = 'UEN:' + uen
    with _lock:
        c = _cache().get(key)
    if c and time.time() - c.get('t', 0) < (30 if c.get('name') else 7) * 86400:
        return {k: v for k, v in c.items() if k != 't'} if c.get('name') else {}
    n = norm(hint_name)
    first = (n[0] if n and n[0].isalpha() else 'OTHERS') if n else ''
    order = ([DATASETS[first]] if first in DATASETS else []) + [v for k, v in DATASETS.items() if DATASETS.get(first) != v]
    hit = {}
    for ds in order:
        try:
            recs = _get(dict(resource_id=ds, filters=json.dumps({'uen': uen}), limit=3), timeout)
        except Exception as e:
            STATE['last_error'] = str(e)[:200]
            return {}                      # no internet / site down: try again next time (nothing cached)
        if recs:
            recs.sort(key=lambda r: not LIVE.search(str(r.get('entity_status_description') or '')))
            r = recs[0]
            hit = dict(name=str(r.get('entity_name') or '').strip().upper(), status=str(r.get('entity_status_description') or ''),
                       address=_address(r))
            break
    with _lock:
        _cache()[key] = dict(hit, t=time.time())
        _save()
    return hit
