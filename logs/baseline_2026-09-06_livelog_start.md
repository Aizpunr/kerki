# Kerki 2026-09-06 - championship state at the moment livelog started

Source: 3 in-game "Championship Leaderboard" screenshots taken by aizpun right after
`/livelog start` (the mod was started LATE - rounds before this point are off-log).
Transcribed verbatim; treat as the OFFSET BASELINE for reconstructing the cup.

Context flags:
- NEW per-cup points system in use (values do not match the documented 100/80/70/60/50/40/35/30
  table, and totals exceed the old 750 finalist cap - RoundNzt already at 910).
  aizpun will supply the new system after the cup.
- aizpun's own run at capture time: 00:42.292 (HUD, top-left).
- Page 1 header says Players: 35/64; pages 2-3 say 36/64 (one join between shots).
- Star icon next to justMaki = nuisance/caster marker in the overlay.
- Eye / crossed-eye icon per row = spectating vs racing, not a scoring field.
- "Time" column = that round's time. Blank for some rows (PlusMicron, Heart-TGV,
  SuperHeavy, lost in trees, DaimyoN) - do NOT assume blank == DNF, PlusMicron shows +79.
- "(+N)" = points gained in the round that had just ended.

| # | Player | Time | Points | Round gain |
|---|--------|------|--------|-----------|
| 1 | RoundNzt | 00:43.132 | 910 | +150 |
| 2 | [KURK] Quickracer10 | 00:42.477 | 635 | +125 |
| 3 | Lexer | 00:42.610 | 632 | +110 |
| 4 | Jake | 00:43.026 | 547 | +85 |
| 5 | [CSC] PlusMicron | | 542 | +79 |
| 6 | [CSC]AndMe18 | 00:43.567 | 537 | +20 |
| 7 | [FPV]PandaMane | 00:42.817 | 534 | +92 |
| 8 | Minkus | 00:42.821 | 497 | +63 |
| 9 | agix | 00:42.994 | 459 | +46 |
| 10 | Murrl | 00:43.146 | 458 | +68 |
| 11 | justMaki (star) | 00:43.064 | 440 | +100 |
| 12 | Eclipse135 | 00:42.997 | 427 | +59 |
| 13 | [CSC] redal | 00:43.770 | 381 | +52 |
| 14 | [T7]Heart-TGV | | 355 | +44 |
| 15 | St Nicholas | 00:43.612 | 350 | +73 |
| 16 | DeiRex | 00:43.695 | 345 | +49 |
| 17 | wokonbike | 00:44.041 | 327 | +35 |
| 18 | aizpun | 00:43.924 | 311 | +37 |
| 19 | [CSC] Shadynook | 00:44.924 | 308 | +35 |
| 20 | gilgool | 00:43.484 | 301 | +42 |
| 21 | Matic_D | 00:43.682 | 291 | +35 |
| 22 | Six | 00:43.660 | 279 | +38 |
| 23 | M4DSCOTSM4N | 00:43.586 | 272 | +35 |
| 24 | brrryy | 00:44.014 | 267 | +55 |
| 25 | Ruckooe | 00:44.075 | 262 | +35 |
| 26 | Noxitu | 00:43.777 | 241 | +36 |
| 27 | [CSC] variableferret | 00:44.524 | 230 | +35 |
| 28 | SuperHeavy | | 224 | +20 |
| 29 | [ZET]LILWOOLEY | 00:46.034 | 217 | +35 |
| 30 | Akane | 00:44.978 | 217 | +35 |
| 31 | voidbloom | 00:45.207 | 205 | +40 |
| 32 | Fly8oy | 00:44.664 | 200 | +20 |
| 33 | lost in trees | | 175 | +20 |
| 34 | sleezy | 00:46.065 | 75 | +35 |
| 35 | Ejol3214 | 00:46.456 | 55 | +35 |
| 36 | DaimyoN | | 20 | +20 |

## Aliases to resolve at ingest
- `Jake` = JakeAdjacent (steamid 76561198152852852), already in build_kerki.py CANONICAL.
- Clan tags [KURK] [CSC] [FPV] [T7] [ZET] are prefixes, strip before matching.
