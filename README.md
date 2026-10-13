# B/L Print Tool

Prepares ocean import shipping documents for batch printing, and pulls the
shipment data out of them into a spreadsheet along the way.

Built for day-to-day freight forwarding work — one vessel at a time,
80–90 shipments per vessel.

## The problem

Carriers send one PDF per document per shipment. A vessel's worth of
paperwork is 200+ loose files. Each shipment needs two sheets of paper:

| | Front | Back |
|---|---|---|
| Sheet 1 | Transport B/L | Waybill |
| Sheet 2 | Pre-Arrival Notice | Arrival Notice (printed later from the TMS) |

Done by hand that means: open a file, print it, take the paper out, put it
back in face-down, open the next file, print it. Times two hundred.

And none of that data ends up anywhere useful — it still gets re-keyed into
the forwarding system by hand afterwards.

## What it does

Pairs the files by bill-of-lading number, merges them in print order, and
writes one PDF per stack. Sheet 1 goes through the printer once with duplex
on; sheet 2 is single-sided so the back stays free.

While it has the PDFs open it also extracts every field worth having —
containers, seals, weights, measurements, shipper, consignee, notify party,
marks, commodity — cross-checks them against the carrier's manifest, and
writes it all to an xlsx with the problems highlighted.

## Quick start

```
pip install pypdf openpyxl xlrd
```

1. Drop the carrier's zip packages and the manifest xls into `1_DROP_ZIPS_HERE`
2. Run `START_HERE.bat`, press `1` to check, then `2` to build
3. Print `2_READY_TO_PRINT\PRINT.pdf` — duplex, flip on **long** edge
4. Pull the pre-arrival notices from Gmail (Apps Script below), drop the
   download into `4_DROP_PREARRIVAL_HERE`, press `4` then `5`
5. Print `PRINT_SHEET2.pdf` — **single sided**

`gmail_grab_prearrival.gs` runs inside a Google Sheet. It searches Gmail by
voyage number, keeps only the newest mail per bill of lading (carriers resend
corrected notices), renames each attachment to its B/L number and drops them
in Drive.

## Folders

```
1_DROP_ZIPS_HERE/         carrier zips + manifest go here
2_READY_TO_PRINT/         the PDFs and the report
3_ARCHIVE/                sources and output, filed by date after each run
4_DROP_PREARRIVAL_HERE/   pre-arrival notices from Drive
```

The bat creates them on first run. Sources are archived after a successful
build so the inbox is empty for the next vessel.

## The report

`_报告.xlsx`, one row per house bill, sorted by HBL:

```
MBL · HBL · containers · shipper · consignee · notify · container no · seal
· type · packages · weight · measurement · marks · description · mode
· origin · destination · ETA · status · sheet no · batch · source files
```

Second tab carries the manifest reconciliation and anything that couldn't be
filed as a row.

### Reading the colors

| | Means |
|---|---|
| Whole row, lilac | Something needs a human — see the status column |
| Whole row, red text | Ramp move, not a port move |
| Whole row, blue fill | Origin port differs from the rest of the vessel |
| Container type cell, orange | Not the usual 40HC |
| Notify cells, yellow | Notify differs from consignee |

Lilac is the one to filter on. The others are "worth knowing", not "go fix it".

## Notes on the PDF parsing

None of the carrier documents are structured. Some things that took a while:

**The waybill table comes out scrambled.** `extract_text()` returns the
columns in the wrong order and layout mode returns nothing at all for the CID
fonts these use. The fix is to collect every text run with its x/y from a
visitor callback and rebuild the lines by coordinate.

**Multi-container waybills use a different layout than single-container
ones** — one is row-wise, the other column-wise. Both are parsed; the counts
are cross-checked against the manifest so a silent mis-parse shows up.

**On the B/L, the field labels are part of the background image,** not text.
There is nothing to search for, so the parties are located by coordinate
region and split on line spacing.

**A consignee block that runs out of room ends in `*`** and continues under
the description of goods. That continuation is stitched back on.

**Container numbers look like B/L numbers** (four letters, seven digits) and
appear earlier in the text, so a naive regex grabs the wrong one.

**One MBL can carry several HBLs.** Keying anything by MBL silently drops the
split bills — they were missing from both the print stack and the report
until this was caught.

## Requirements

- Python 3, `pypdf openpyxl xlrd`
- A printer with automatic duplex, Letter paper
  (the automatic duplexer on these Brother units doesn't take A4, so the
  A4-sized source PDFs are printed with "shrink oversized pages")

## Privacy

Everything runs locally. Nothing is uploaded. The data folders are
gitignored — the documents contain real customer names, addresses and
contact details.
