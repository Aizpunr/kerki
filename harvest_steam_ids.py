"""Feed Kerki's log-observed Steam IDs into the ONE shared Zeepkist registry.

Kerki deliberately keeps no steam_ids.json of its own (aizpun 2026-09-21:
"zeepkist players are zeepkist players"). The archived LiveLeaderboardLogger
logs in kerki/logs/ are direct observation, which is the highest-confidence
source there is, so this script offers them to the shared registry.

    python harvest_steam_ids.py            # dry run, prints what it would add
    python harvest_steam_ids.py --write    # apply

ADDITIVE ONLY, BY DESIGN. It never deletes and never rewrites the file from
its own view of the world. seed_steam_ids.py once destroyed 285 entries by
regenerating from a narrower source set; this merges instead.
"""
import argparse
import glob
import json
import os
import re
import sys
from collections import defaultdict

# Zeepkist names carry emoji and CJK; never let the console encoding kill a run.
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except AttributeError:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
LOGS = os.path.join(HERE, 'logs', 'kerki_*.log')
REGISTRY = os.path.normpath(
    os.path.join(HERE, '..', 'zeepkist cotd elo', 'steam_ids.json'))


def strip_tag(name):
    return re.sub(r"^\[[^\]]*\]\s*", "", name or "").strip()


def observed():
    """{sid: {name: cups_seen}} from every archived Kerki log."""
    seen = defaultdict(lambda: defaultdict(set))
    for path in sorted(glob.glob(LOGS)):
        m = re.search(r'kerki_(\d+)\.log', path)
        cup = int(m.group(1)) if m else 0
        with open(path, encoding='utf-8', errors='ignore') as f:
            for line in f:
                r = re.search(r'ROSTER\|(\d+)\|([^|]+)\|', line)
                if r:
                    seen[r.group(1)][r.group(2)].add(cup)
    return seen


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--write', action='store_true',
                    help='apply the additions (default is a dry run)')
    ap.add_argument('--registry', default=REGISTRY)
    args = ap.parse_args()

    if not os.path.exists(args.registry):
        print(f"ERROR: no registry at {args.registry}", file=sys.stderr)
        sys.exit(1)
    with open(args.registry, encoding='utf-8') as f:
        reg = json.load(f)                      # {canonical: "76561..."}

    known_sids = {str(v): k for k, v in reg.items() if v}
    known_names = {k.lower() for k in reg}

    adds, conflicts = {}, []
    for sid, names in sorted(observed().items()):
        if sid in known_sids:
            continue                            # already registered, leave it
        # Prefer the name worn in the most cups, then the longest, as the
        # canonical. Ties are rare and the operator reviews the dry run anyway.
        best = sorted(names.items(), key=lambda kv: (-len(kv[1]), -len(kv[0])))[0][0]
        best = strip_tag(best)
        # Junk names: empty, or punctuation only (one player shows up as "''").
        # A registry entry keyed on those would be unmatchable noise.
        if not best or not any(c.isalnum() for c in best):
            continue
        if best.lower() in known_names:
            conflicts.append((sid, best, reg.get(best)))
            continue
        adds[best] = sid

    print(f"registry: {len(reg)} entries at {args.registry}")
    print(f"kerki logs: {len(observed())} distinct steam ids observed")
    print(f"\nNEW players to add: {len(adds)}")
    for name, sid in sorted(adds.items(), key=lambda kv: kv[0].lower()):
        print(f"  {name:28} {sid}")
    if conflicts:
        print(f"\nSKIPPED, name already taken by another entry ({len(conflicts)}) "
              f"— review by hand:")
        for sid, name, existing in conflicts:
            print(f"  {name:28} observed sid {sid}, registry has {existing}")

    if not args.write:
        print("\nDry run. Re-run with --write to apply.")
        return
    if not adds:
        print("\nNothing to add.")
        return

    reg.update(adds)
    merged = {k: reg[k] for k in sorted(reg, key=str.lower)}
    tmp = args.registry + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(merged, f, indent=2, ensure_ascii=False)
        f.write('\n')
    os.replace(tmp, args.registry)
    print(f"\nWrote {len(adds)} new entries -> {args.registry} "
          f"({len(merged)} total, nothing removed)")
    print("Now re-run build_players_master.py so players.json picks them up.")


if __name__ == '__main__':
    main()
