"""
Watcher: detecta guardados del Excel y regenera index.html + hace git push.
Ejecutar una vez, deja corriendo en segundo plano.
"""
import os, time, subprocess, sys, shutil
from datetime import datetime

EXCEL_PATH   = r'C:\Users\mcasanovas\OneDrive - IVASCULAR, S.L.U\Planning General.xlsm'
REPO_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GEN_SCRIPT   = os.path.join(REPO_DIR, 'scripts', 'gen_planning.py')
POLL_SECS    = 8   # check every 8 seconds
DEBOUNCE     = 5   # wait 5 s after last change before triggering (Excel writes temp files)

def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f'[{ts}] {msg}', flush=True)

def get_mtime():
    try:
        return os.path.getmtime(EXCEL_PATH)
    except FileNotFoundError:
        return None

def regenerate():
    log('Cambio detectado → regenerando HTML...')
    result = subprocess.run(
        [sys.executable, GEN_SCRIPT],
        capture_output=True, text=True, cwd=REPO_DIR
    )
    if result.returncode != 0:
        log(f'ERROR al generar: {result.stderr[:300]}')
        return False
    log(result.stdout.strip().split('\n')[0])
    return True

def git_push():
    cmds = [
        ['git', 'add', 'index.html'],
        ['git', 'commit', '-m', f'auto: update {datetime.now().strftime("%Y-%m-%d %H:%M")}'],
        ['git', 'push'],
    ]
    for cmd in cmds:
        r = subprocess.run(cmd, capture_output=True, text=True, cwd=REPO_DIR)
        if r.returncode != 0 and 'nothing to commit' not in r.stdout + r.stderr:
            log(f'Git error ({" ".join(cmd[:2])}): {r.stderr[:200]}')
            return False
    log('Git push OK → https://mcasanovas23.github.io/planning-dashboard/')
    return True

def main():
    log(f'Watcher iniciado. Monitorizando: {EXCEL_PATH}')
    log(f'Repositorio: {REPO_DIR}')
    last_mtime   = get_mtime()
    pending_mtime = None
    pending_since = None

    while True:
        time.sleep(POLL_SECS)
        current = get_mtime()

        if current is None:
            log('Archivo Excel no encontrado, reintentando...')
            continue

        if current != last_mtime:
            # File changed — start debounce
            if current != pending_mtime:
                pending_mtime = current
                pending_since = time.time()
                log('Cambio detectado, esperando estabilización...')

        if pending_mtime and (time.time() - pending_since) >= DEBOUNCE:
            last_mtime   = pending_mtime
            pending_mtime = None
            pending_since = None
            if regenerate():
                git_push()

if __name__ == '__main__':
    main()
