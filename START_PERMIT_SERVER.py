# -*- coding: utf-8 -*-
"""INFINITY PERMIT SERVER - double-click this file to start the server (Python must be installed).

First start: installs the libraries it needs (internet, about 2 minutes) - only once.
Then the server runs in this window and the browser opens http://localhost:8080.
Keep this window open (you can minimise it) - closing it stops the server.
"""
import os, sys, subprocess, importlib, traceback

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)

NEEDED = [('flask', 'flask>=3.0'), ('waitress', 'waitress>=3.0'), ('openpyxl', 'openpyxl'), ('xlrd', 'xlrd'),
          ('pypdf', 'pypdf>=4.0'), ('pdfplumber', 'pdfplumber'), ('pypdfium2', 'pypdfium2'), ('PIL', 'Pillow'),
          ('numpy', 'numpy')]
OPTIONAL = {'waitress>=3.0', 'xlrd'}           # the server also runs without these (built-in web server / .xlsx only)


def missing(only_required=False):
    out = []
    for mod, pkg in NEEDED:
        if only_required and pkg in OPTIONAL:
            continue
        try:
            importlib.import_module(mod)
        except Exception:
            out.append(pkg)
    return out


def main():
    print('=' * 70)
    print(' INFINITY SOLUTIONS - PERMIT SERVER')
    print(' Folder: ' + HERE)
    print('=' * 70)
    need = missing()
    if need:
        print('First start: installing %s ...' % ', '.join(need))
        print('(needs internet, about 1-3 minutes - only this once)')
        cmd = [sys.executable, '-m', 'pip', 'install', '--disable-pip-version-check', '--user'] + need
        for pkg in need:                                   # one by one: one failing package does not stop the others
            if subprocess.call(cmd[:6] + [pkg]) != 0:
                subprocess.call(cmd[:5] + [pkg])           # without --user (e.g. inside a virtual env)
        try:                                               # a new user site folder is not on sys.path yet
            import site
            us = site.getusersitepackages()
            if us not in sys.path and os.path.isdir(us):
                sys.path.append(us)
        except Exception:
            pass
        importlib.invalidate_caches()
        still = missing(only_required=True)
        if still:
            print('\n*** Could not install: %s' % ', '.join(still))
            print('*** Check the internet connection, or install Python 3.12 from python.org and try again.')
            input('\nPress Enter to close ...')
            return
        print('Libraries installed.\n')
    if not os.path.exists(os.path.join(HERE, 'web', 'index.html')):
        print('*** The folder "web" is missing next to this file - unzip the WHOLE PERMIT_SERVER.zip again.')
        input('\nPress Enter to close ...')
        return
    import server
    sys.argv = [os.path.join(HERE, 'server.py')] + sys.argv[1:]
    server.main()


if __name__ == '__main__':
    try:
        main()
    except SystemExit:
        pass
    except KeyboardInterrupt:
        pass
    except Exception:
        traceback.print_exc()
        print('\n*** The server stopped with an error (above). Send a screenshot of this window.')
        input('\nPress Enter to close ...')
