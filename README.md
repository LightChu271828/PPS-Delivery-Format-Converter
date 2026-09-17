# Kentucky Lottery PPS — delivery tabs

A single-page tool that adds the four Kentucky delivery tabs to an uploaded PPS and lets you
download a new workbook. Open the page and everything runs in the browser. The file is not
uploaded to a server.

**Live:** https://lightchu271828.github.io/ky-pps-delivery/

Same shape as [must-go jackpot designer](https://lightchu271828.github.io/mustgo-jackpot-designer/).

## What it writes

- `Delivery`
- `Odds Table`
- `Summary(Delivery)`
- `Prize Breakdown`

Source sheets stay. Existing delivery tabs are replaced.

Schema is auto-detected from the workbook:

| Schema | How it is recognised | Layout |
| --- | --- | --- |
| SSJ | Jackpots named `small` / `large`, or `SSJ` in the filename / labels | One ticket price |
| MMJ3 | `MMJ3` in the filename / labels, or 1–2 jackpots that are not small/large | Price grid + JP |
| Daily Streak | No jackpot data, Daily Streak in the filename / labels | No-JP grid including $3 |
| No jackpot | Blank Progressive Jackpots | No-JP price grid |

A generic MMJ (not MMJ3) currently uses the MMJ3 layout. Use the schema dropdown to override.

## What you need in the file

Frequency must already be calculated and saved in Excel so method names and prizes are cached.
If the page cannot open a workbook, open it in Excel and Save As `.xlsx` first. Combined
multi-price Delivery packs are not inputs — upload the per-price PPS.

## Run locally

```powershell
cd ky-pps-delivery
python -m http.server 8765
```

Open http://127.0.0.1:8765/
