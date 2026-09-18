# PPS Delivery Format Converter

A single-page tool that adds the four Kentucky delivery tabs to an uploaded PPS and lets you
download a new workbook. Open the page and everything runs in the browser. The file is not
uploaded to a server.

**Live:** https://lightchu271828.github.io/PPS-Delivery-Format-Converter/

## What it writes

- `Delivery`
- `Odds Table`
- `Summary(Delivery)`
- `Prize Breakdown`

Source sheets stay. Existing delivery tabs are replaced.

Jackpot type is read from the filename and Win Methods, and shown with that exact spelling:

| Type | How it is recognised | Layout |
| --- | --- | --- |
| ChatterJP | `ChatterJP` in the filename or Win Methods | Default 8-price grid + JP |
| MMJ3 | `MMJ3` in the filename or Win Methods | Default 8-price grid + JP |
| MMJ | `MMJ` in the filename or Win Methods | Default 8-price grid + JP |
| SSJ | `SSJ` in the filename or Win Methods | Default 8-price grid + JP |
| Daily Streak | No jackpot data, Daily Streak in the filename / labels | No-JP grid including $3 |
| No jackpot | Blank Progressive Jackpots | Default 8-price grid |

Default ticket prices are `0.5, 1, 2, 5, 10, 20, 30, 50`, all selected. Uncheck prices you do not want; the grid must still include `$1`. A one-price SSJ pack such as X the Money is the exception, not the site default. If jackpot rows exist but none of MMJ / MMJ3 / SSJ / ChatterJP is in the filename or Win Methods, the build stops until the filename names the type.

Win Methods FreePlay prizes are rewritten to ordinary `SUMPRODUCT(SUMIF(...))` formulas matched by Free Plays method identity, not CSE `{=SUM(VLOOKUP(...))}`.

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
