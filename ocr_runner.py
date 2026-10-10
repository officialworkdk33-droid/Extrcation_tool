# -*- coding: utf-8 -*-
"""Runs tesseract.exe directly (no pytesseract).

Why: in a --windowed exe on a PC WITHOUT Tesseract installed, the copy packed inside the exe must be started with
its own tessdata folder, a hidden window and no console handles. Doing it ourselves makes that reliable and lets
the tool show the exact reason when OCR cannot start.

Search order for tesseract.exe:
  1. path typed on tab 1
  2. folder "Tesseract-OCR" next to the exe            (copy it there - recommended for other PCs)
  3. folder "Tesseract-OCR" packed inside the exe      (BUILD_EXE.bat packs it) - first copied once to
     %LOCALAPPDATA%\\PermitToolOCR\\Tesseract-OCR and run from there, then straight from the exe's temp folder
  4. C:\\Program Files\\Tesseract-OCR, Program Files (x86), %LOCALAPPDATA%\\Programs\\Tesseract-OCR
  5. tesseract on PATH
"""
import os, sys, subprocess, tempfile, shutil, csv, io

STATE = dict(cmd='', tessdata='', version='', error='', tried=[])


def _app_dir():
    return os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))


def _local_copy():
    """The Tesseract packed inside the exe is unpacked to a temp folder that antivirus / company rules often block.
    Copy it once to %LOCALAPPDATA%\\PermitToolOCR\\Tesseract-OCR and run it from there (nothing to install)."""
    src = os.path.join(getattr(sys, '_MEIPASS', ''), 'Tesseract-OCR')
    if not os.path.isdir(src):
        return ''
    base = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
    dst = os.path.join(base, 'PermitToolOCR', 'Tesseract-OCR')
    files = []
    for root, _, fs in os.walk(src):
        for f in fs:
            fp = os.path.join(root, f)
            files.append((os.path.relpath(fp, src), os.path.getsize(fp)))
    sig = '%d %d' % (len(files), sum(sz for _, sz in files))
    mark = os.path.join(dst, '_copied.txt')
    try:
        if os.path.exists(mark) and open(mark).read().strip() == sig:
            return dst
        for rel, sz in files:
            t = os.path.join(dst, rel)
            if os.path.exists(t) and os.path.getsize(t) == sz:
                continue
            os.makedirs(os.path.dirname(t), exist_ok=True)
            shutil.copy2(os.path.join(src, rel), t)
        with open(mark, 'w') as f:
            f.write(sig)
        return dst
    except Exception as e:
        STATE['tried'].append('%s (copy failed: %s)' % (dst, e))
        return ''


def candidates(typed=''):
    exe = 'tesseract.exe' if os.name == 'nt' else 'tesseract'
    out = []
    if typed:
        out.append(typed if typed.lower().endswith(('tesseract.exe', 'tesseract')) else os.path.join(typed, exe))
    out.append(os.path.join(_app_dir(), 'Tesseract-OCR', exe))
    if getattr(sys, '_MEIPASS', None):
        lc = _local_copy()
        if lc:
            out.append(os.path.join(lc, exe))
        out.append(os.path.join(sys._MEIPASS, 'Tesseract-OCR', exe))
    out += [r'C:\Program Files\Tesseract-OCR\tesseract.exe', r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Tesseract-OCR', 'tesseract.exe')]
    w = shutil.which('tesseract')
    if w:
        out.append(w)
    seen, res = set(), []
    for p in out:
        if p and p not in seen:
            seen.add(p)
            res.append(p)
    return res


def _run(args, timeout=180):
    kw = dict(stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    env = dict(os.environ)
    if STATE['tessdata']:
        env['TESSDATA_PREFIX'] = STATE['tessdata']
    kw['env'] = env
    if os.name == 'nt':
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        kw['startupinfo'] = si
        kw['creationflags'] = 0x08000000            # CREATE_NO_WINDOW
        kw['cwd'] = os.path.dirname(args[0]) or None   # its DLLs sit next to tesseract.exe
    return subprocess.run(args, **kw)


def setup(typed=''):
    """Find a working tesseract. Returns True / False; STATE holds cmd, version or the error."""
    STATE.update(cmd='', tessdata='', version='', error='', tried=[])
    errs = []
    for c in candidates(typed):
        if os.path.isabs(c) and not os.path.exists(c):
            STATE['tried'].append('%s (not found)' % c)
            continue
        td = os.path.join(os.path.dirname(c), 'tessdata')
        STATE['tessdata'] = td if os.path.isdir(td) else ''
        try:
            r = _run([c, '--version'], timeout=30)
            txt = (r.stdout or b'').decode('utf-8', 'replace') + (r.stderr or b'').decode('utf-8', 'replace')
            if r.returncode == 0 and 'tesseract' in txt.lower():
                if STATE['tessdata'] and not os.path.exists(os.path.join(STATE['tessdata'], 'eng.traineddata')):
                    errs.append('%s: tessdata\\eng.traineddata missing' % c)
                    STATE['tried'].append('%s (eng.traineddata missing)' % c)
                    continue
                STATE.update(cmd=c, version=txt.splitlines()[0].strip())
                return True
            errs.append('%s: exit %s %s' % (c, r.returncode, txt.strip()[:120]))
            STATE['tried'].append('%s (exit %s)' % (c, r.returncode))
        except Exception as e:
            errs.append('%s: %s' % (c, e))
            STATE['tried'].append('%s (%s)' % (c, type(e).__name__))
    STATE['tessdata'] = ''
    STATE['error'] = '; '.join(errs) or 'tesseract.exe not found'
    return False


def _args(img_path, out_base, psm, fmt):
    a = [STATE['cmd'], img_path, out_base, '--psm', str(psm), '-l', 'eng']
    if STATE['tessdata']:
        a += ['--tessdata-dir', STATE['tessdata']]
    if fmt == 'tsv':
        a.append('tsv')
    return a


def _ocr(img, psm, fmt):
    if not STATE['cmd'] and not setup():
        raise RuntimeError('Tesseract OCR not available: ' + STATE['error'])
    tmp = tempfile.mkdtemp(prefix='ocr_')
    try:
        ip = os.path.join(tmp, 'p.png')
        img.save(ip)
        base = os.path.join(tmp, 'out')
        r = _run(_args(ip, base, psm, fmt))
        out = base + ('.tsv' if fmt == 'tsv' else '.txt')
        if not os.path.exists(out):
            raise RuntimeError('tesseract failed (exit %s): %s' % (r.returncode, (r.stderr or b'').decode('utf-8', 'replace')[:300]))
        with open(out, encoding='utf-8', errors='replace') as f:
            return f.read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def image_to_string(img, psm=6):
    return _ocr(img, psm, 'txt')


def image_to_data(img, psm=6):
    """Same dict shape as pytesseract.image_to_data(output_type=DICT)."""
    raw = _ocr(img, psm, 'tsv')
    keys = ['level', 'page_num', 'block_num', 'par_num', 'line_num', 'word_num', 'left', 'top', 'width', 'height',
            'conf', 'text']
    d = {k: [] for k in keys}
    rd = csv.reader(io.StringIO(raw), delimiter='\t', quoting=csv.QUOTE_NONE)
    head = next(rd, None)
    for row in rd:
        if len(row) < 11:
            continue
        row = row + [''] * (12 - len(row))
        for k, v in zip(keys, row[:12]):
            if k == 'text':
                d[k].append(v)
            elif k == 'conf':
                try:
                    d[k].append(float(v))
                except ValueError:
                    d[k].append(-1.0)
            else:
                try:
                    d[k].append(int(v))
                except ValueError:
                    d[k].append(0)
    return d
