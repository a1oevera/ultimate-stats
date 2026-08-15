# ultimate-stats

Turn a [Statto](https://www.statto.app/) ultimate-frisbee export into a clean,
multi-tab `.xlsx` that imports straight into Google Sheets.

## Usage

```
python3 parse_statto.py [input] [output.xlsx]
```

`input` can be either:
- a `.statto` export (a zip archive containing `data.json`, Statto's native
  download format), or
- a plain `data.json` file extracted from one.

Defaults: input=`data.json`, output=`statto_stats.xlsx`

Example:

```
python3 parse_statto.py Toro_Women_2026_2026-08-14_14-09-12.statto natsstats.xlsx
```

## Output tabs

| Tab | Contents |
|---|---|
| Read Me | Column definitions + assumptions/quirks baked into this file |
| Player Totals | One row per player, whole tournament (live formulas) |
| Player by Game | One row per player per game (live formulas) |
| Games | One row per game: reconstructed score, result, wind |
| Points | One row per real point: lineup, who started O/D, scored/conceded |
| Lineups | Long format, one row per player per point (pivot engine for +/-) |
| Passes | One row per pass, UUIDs resolved to names |
| Possessions | One row per possession |
| Blocks | One row per defensive block |

The summary tabs use `COUNTIFS`/`SUMIFS` referencing the raw tabs, so they
recalculate if you filter or edit the raw data in Sheets — nothing is a
hardcoded aggregate.

## Stat definitions

| Stat | Definition |
|---|---|
| Goal | Receiver of a pass flagged `isAssist` |
| Assist | Thrower of a pass flagged `isAssist` |
| 2nd Assist | Thrower of a pass flagged `isSecondaryAssist` (the hockey assist) |
| D (Block) | A `defensiveBlocks` event credited to the player |
| Throwaway | A pass flagged `isThrowerError` (credited to the thrower) |
| Drop | A pass flagged `isReceiverError` (credited to the receiver) |
| Turnover | Throwaways + Drops |
| Throw | Any pass thrown by the player |
| Completion | A throw that was NOT a throwaway |
| Completion% | Completions / Throws |
| Reception | Receiver on a completed pass (not a throwaway, not a drop) |
| Point Played | The player was one of the 7 on the field for that point |
| +/- | (points won while on field) − (points lost while on field) |

Edit `parse_statto.py` if your team scores things differently.

## Requirements

```
pip install openpyxl
```
