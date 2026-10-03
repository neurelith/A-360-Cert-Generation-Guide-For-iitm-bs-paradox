# Step 1 & 2 — Data Preparation

Before any code runs, the spreadsheet needs to be set up. Two things must exist for every candidate row: a **unique ID** and a **verification URL**.

---

## The CSV Schema

The pipeline expects a CSV with exactly these 12 columns. Headers must be lowercase with underscores.

| Column | Header | Example | Notes |
|--------|--------|---------|-------|
| A | `cert_id` | `SAAVAN26-W-0001` | Unique code. Used as the PDF filename. |
| B | `type` | `Winner` | One of: `Winner`, `Participant`, `Judge`, `Guest`. Controls which template and HTML email gets used. |
| C | `department` | `Cultural` | Organizing vertical. Not used in the certificate itself, just for internal tracking. |
| D | `event_name` | `Stand-Up Comedy Showdown` | Replaces `<<event_name>>` on the slide and `{{EventName}}` in the email. |
| E | `name` | `Krishna` | Replaces `<<name>>` on the slide and `{{StudentName}}` in the email. |
| F | `email` | `25f2007055@ds.study.iitm.ac.in` | Candidate's actual email. In test mode, delivery goes to the staging addresses instead. |
| G | `position` | `Winner` | Role/rank: `Winner`, `First Runner Up`, `Participant`, `Judge`. Replaces `<<position>>` on winner certs. |
| H | `uuid` | `085AAAF8-1A41-5290-98BD-54E5CC4C39B5` | Unique hex token. Generated via Excel formula, then frozen as a static value. |
| I | `qr_code` | `https://saavan.iitmparadox.org/verify?cert=085AAAF8-...` | Verification link built from the UUID. Gets encoded into the QR code on the certificate. |
| J | `winning_prize` | `INR 5,000 Cash Prize` | Prize description for winners. Use `N/A` for participants/judges. |
| K | `feedback_link` | `https://forms.gle/saavan26feedback` | Google Form for event feedback. |
| L | `discrepancy_link` | `https://forms.gle/saavan26discrepancy` | Form to report name typos or wrong certificates. |

---

## Step 1: Generate UUIDs (Column H)

Each certificate needs a unique, non-guessable identifier. We generate RFC 4122-style hex tokens directly in the spreadsheet so the data team doesn't need to run any code.

### Excel / Google Sheets formula

Paste this into cell `H2` and fill down:

```
=UPPER(CONCATENATE(
  DEC2HEX(RANDBETWEEN(0,4294967295),8), "-",
  DEC2HEX(RANDBETWEEN(0,65535),4), "-",
  DEC2HEX(RANDBETWEEN(16384,20479),4), "-",
  DEC2HEX(RANDBETWEEN(32768,49151),4), "-",
  DEC2HEX(RANDBETWEEN(0,65535),4),
  DEC2HEX(RANDBETWEEN(0,4294967295),8)
))
```

This produces values like `085AAAF8-1A41-5290-98BD-54E5CC4C39B5`.

### Why this works
- `DEC2HEX(RANDBETWEEN(0,4294967295),8)` generates an 8-character hex block.
- The segments follow the `8-4-4-4-12` UUID layout.
- `UPPER()` forces uppercase hex so the printed certificate and QR URL look consistent.

### Freeze it immediately

`RANDBETWEEN` is volatile — it recalculates every time you touch the sheet. If you don't freeze the values, every candidate gets a new UUID whenever someone edits an unrelated cell.

1. Select Column H (all the UUID cells).
2. Copy (`Ctrl+C`).
3. **Paste as Values** — Google Sheets: `Ctrl+Shift+V`. Excel: `Paste Special → Values`.
4. Confirm the cells now show raw text, not formulas.

---

## Step 2: Build Verification URLs (Column I)

Once the UUID is frozen, build the verification link in the `qr_code` column.

### Formula (cell I2, assuming UUID is in H2)

```
="https://saavan.iitmparadox.org/verify?cert=" & H2
```

This is what gets encoded into the QR code printed on the physical certificate. When someone scans it, the verification site looks up the UUID and shows the candidate's details.

---

## Step 3: Generate cert_id (Column A — optional helper)

If you want auto-generated sequential certificate codes:

```
="SAAVAN26-" & IF(B2="Winner","W-",IF(B2="Participant","P-","J-")) & TEXT(ROW()-1,"0000")
```

Produces: `SAAVAN26-W-0001`, `SAAVAN26-P-0002`, `SAAVAN26-J-0003`, etc.

---

## Exporting the CSV

1. **File → Download → Comma-separated values (.csv)**
2. Make sure the encoding is **UTF-8**. This matters for names with accents or non-Latin characters.
3. Delete any trailing blank rows at the bottom. Pandas reads them as `NaN` entries and the generator will crash.
4. If any field contains commas (like `"INR 5,000 Cash Prize"`), the spreadsheet should auto-wrap it in double quotes. Verify this by opening the `.csv` in a text editor.
5. Save as `dummy_data.csv` inside `saavan26/dummy/` (or pass a custom path if using the universal generator).

---

Next: [Certificate Verification & Validation →](02-verification-process.md)
