"""Replay every archived Kerki through the parser and diff it against kerki.json.

Run this after ANY change to parse_kerki_from_log.py. It is the only thing that
says whether a fix aimed at one cup quietly moved another. Winners and finalists
must match exactly; Others points are not compared, because a parser fix may
legitimately correct them (see project_kerki.md).

    python replay.py            all cups
    python replay.py 42 43      just those

FLAGS below is the authority for how each cup is reproduced. Keep it in step with
the reproduction table in project_kerki.md whenever a cup is added or a flag found.
"""
import json
import os
import subprocess
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)

# cup -> extra flags beyond --log/--date/--mappers
FLAGS = {
    35: ["--discovered", "none"],
    36: ["--discovered", "none"],
    37: ["--no-warmup"],
    38: ["--no-warmup"],
    39: ["--maps", "Mount Krithon", "--discovered", "pratia"],
    41: ["--no-warmup", "--baseline", "logs/kerki_41_baseline.json"],
    42: ["--no-warmup"],
    43: ["--no-warmup"],
    44: ["--no-warmup"],
    45: ["--no-warmup", "--exclude-maps", "Timed Variety P", "--skip-rounds", "7,8"],
}

# kerki.json's date is wrong for these; the log is the authority.
DATE = {36: "2026-05-24"}


def iso(d):
    dd, mm, yy = d.split("/")
    return f"20{yy}-{int(mm):02d}-{int(dd):02d}"


def main():
    want = {int(a) for a in sys.argv[1:]} or set(FLAGS)
    reg = json.load(open(os.path.join("..", "zeepkist cotd elo", "players.json"),
                         encoding="utf-8"))
    sid = {}
    for canon, v in reg.items():
        if v.get("steam_id"):
            sid[canon.lower()] = v["steam_id"]
            for a in v.get("aliases") or []:
                sid.setdefault(a.lower(), v["steam_id"])

    pub = {k["id"]: k for k in json.load(open("kerki.json", encoding="utf-8"))["kerkis"]}
    bad = 0
    for n in sorted(want):
        p = pub[n]
        missing = [x for x in p["nuisances"] if x.lower() not in sid]
        mappers = ",".join(sid[x.lower()] for x in p["nuisances"] if x.lower() in sid)
        cmd = [sys.executable, "parse_kerki_from_log.py", str(n),
               "--log", f"logs/kerki_{n}.log",
               "--date", DATE.get(n, iso(p["date"]))] + FLAGS[n]
        if mappers:
            cmd += ["--mappers", mappers]
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        if r.returncode != 0:
            print(f"FAIL #{n}  parser exited {r.returncode}")
            print((r.stdout + r.stderr)[-700:])
            bad += 1
            continue
        got = json.load(open(f"kerki_{n}.json", encoding="utf-8"))
        gw = [w["name"] for w in got["winners"]]
        gf = [f["name"] for f in got["finalists"]]
        ok = gw == p["winners"] and gf == p["finalists"]
        bad += 0 if ok else 1
        print(f"{'OK  ' if ok else 'FAIL'} #{n}  rounds={got['scoring_rounds_count']}")
        if gw != p["winners"]:
            print(f"      winners   published {p['winners']}\n                parsed    {gw}")
        if gf != p["finalists"]:
            print(f"      finalists published {p['finalists']}\n                parsed    {gf}")
        if missing:
            print(f"      (nuisance not in the registry, not excluded: {missing})")

    print(f"\n{len(want) - bad}/{len(want)} reproduce exactly")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
