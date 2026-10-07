# PPS Delivery Format Converter

A single-page tool that adds lottery Delivery tabs to an uploaded PPS and lets you download a
new workbook. Open the page and everything runs in the browser. The file is not uploaded to a
server.

**Live:** https://lightchu271828.github.io/PPS-Delivery-Format-Converter/

Pick the lottery first. Kentucky is a four-tab pack. NC, VA, GA, PA, and NH add one Delivery
sheet. Every source tab stays in the workbook.

## Kentucky

Writes `Delivery`, `Odds Table`, `Summary(Delivery)`, and `Prize Breakdown`. Source sheets stay.
Existing delivery tabs are replaced.

Jackpot type is read from the filename and Win Methods, and shown with that exact spelling:

| Type | How it is recognised | Layout |
| --- | --- | --- |
| ChatterJP | `ChatterJP` in the filename or Win Methods | Default 8-price grid + JP |
| MMJ3 | `MMJ3` in the filename or Win Methods | Default 8-price grid + JP |
| MMJ | `MMJ` in the filename or Win Methods | Default 8-price grid + JP |
| SSJ | `SSJ` in the filename or Win Methods | Default 8-price grid + JP |
| Daily Streak | No jackpot data, Daily Streak in the filename / labels | No-JP grid including $3 |
| No jackpot | Blank Progressive Jackpots | Default 8-price grid |

Default ticket prices are `0.5, 1, 2, 5, 10, 20, 30, 50`, all selected. Uncheck any prices you do
not want; Odds and Summary scale from `$1` when it is selected, otherwise from the first remaining
price. A one-price SSJ pack such as X the Money is the exception, not the site default. If jackpot
rows exist but none of MMJ / MMJ3 / SSJ / ChatterJP is in the filename or Win Methods, the build
stops until the filename names the type.

Win Methods FreePlay prizes are rewritten to ordinary `SUMPRODUCT(SUMIF(...))` formulas matched by
Free Plays method identity, not CSE `{=SUM(VLOOKUP(...))}`.

## NC, VA, GA, PA, NH

One `Delivery` sheet is added to the uploaded PPS. The price grid is not used. PA and NH use the
Georgia layout. Frequency rows whose Odds up is 0 are skipped.

| Lottery | Layout |
| --- | --- |
| NC | Frequency order from row 4. Summary in K/L beside the tiers. Bonus / Freeplay show `-` when that category is missing. |
| VA | The NC table with the summary in B/C above it; the tier header sits three rows below Hit Rate. Delivery is placed right after Frequency. Side-bet hit rates go in E/F only when Frequency lists them. |
| GA, PA, NH | Winning tiers sorted by prize, smallest first. Consolidated 1-in-X and `% of Prize Fund`. Bonus and Freeplay lines are always present. |

JP RTP and Total RTP rows appear only for jackpot games. The Progressive Jackpots block is copied
below the total row with its borders and fills, without the json setup form, and JP RTP links to
it. Non-jackpot games have no JP rows. Output name:

`{yymmdd}_{JUR}_{Game}(Delivery)_PPS_{rtp}.xlsx`

Number formats are plain `#,##0`, `#,##0.00`, and `0.00%`. Accounting padding is stripped.

Buy-feature packs are still one workbook per game for now.

## What you need in the file

Frequency must already be calculated and saved in Excel so method names and prizes are cached.
If the page cannot open a workbook, open it in Excel and Save As `.xlsx` first. Combined
multi-price Delivery packs are not inputs — upload the per-price PPS.

## Run locally

```powershell
cd PPS-Delivery-Format-Converter
python -m http.server 8765
```

Open http://127.0.0.1:8765/
