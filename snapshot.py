"""snapshot.py — freeze the current Rolling Points + Glicko Skill rankings
and per-player history into snapshot.json so the NEXT build's leaderboard
shows ▲/▼ rank arrows for the just-finished kerki.

Run BEFORE adding a new kerki to the spreadsheet (i.e. before running
parse_kerki_from_log.py for the new kerki), so the snapshot captures the
pre-new-cup state. If you forget, the recovery is the same as TyO:
restore kerki.json from git history (commit just before the new cup),
copy its rankings here, rebuild.

Usage: python snapshot.py
"""
import json
import os
import shutil
from datetime import datetime

base = os.path.dirname(os.path.abspath(__file__))
def _p(f): return os.path.join(base, f)


with open(_p('kerki.json'), encoding='utf-8') as f:
    data = json.load(f)

current_kerki = (data.get('meta') or {}).get('last_kerki', 0)

snap_path = _p('snapshot.json')
backup_dir = _p('old snapshots')
if os.path.exists(snap_path):
    with open(snap_path, encoding='utf-8') as fp:
        old = json.load(fp)
    old_kerki = (old.get('_meta') or {}).get('kerki', 'unknown')
    os.makedirs(backup_dir, exist_ok=True)
    backup_path = os.path.join(backup_dir, f'snapshot {old_kerki}.json')
    i = 0
    while os.path.exists(backup_path):
        i += 1
        backup_path = os.path.join(backup_dir, f'snapshot {old_kerki}_{i}.json')
    shutil.copy2(snap_path, backup_path)
    print(f'Backed up -> old snapshots/{os.path.basename(backup_path)}')


# Kerki rankings are keyed by name (no steamid in the JSON).
snap = {
    '_meta':  {
        'kerki': current_kerki,
        'generated': datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
    },
    'ranking':  {p['name']: [p['rank'], p.get('points', 0)]
                 for p in data.get('ranking', {}).get('players', [])},
    'glicko':   {p['name']: [p['rank'], p.get('mu', 1500)]
                 for p in data.get('glicko', {}).get('players', [])},
    'history':  {p['name']: p.get('history', [])
                 for p in data.get('players', [])},
}

tmp = snap_path + '.tmp'
with open(tmp, 'w', encoding='utf-8') as fp:
    json.dump(snap, fp, separators=(',', ':'))
os.replace(tmp, snap_path)

print(
    f'snapshot.json written (kerki #{current_kerki}, '
    f'ranking={len(snap["ranking"])}, glicko={len(snap["glicko"])}, '
    f'history={len(snap["history"])})'
)
