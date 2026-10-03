# Step 3 — Certificate Generation

This is where the code takes over. `generate_certificates.py` reads the CSV, fills in PowerPoint templates, generates QR codes, and exports PDFs.

```bash
python generate_certificates.py
```

That's it from the user's side. But here's what actually happens underneath.

---

## The Pipeline (inside one script)

```
CSV row                                  Final output
┌──────────────┐                         ┌──────────────┐
│ name         │   1. Pick template      │ SAAVAN26-     │
│ event_name   │   2. Replace tags       │ W-0001.pptx  │
│ position     │   3. Insert QR code     │              │
│ type         │   4. Save PPTX          │ SAAVAN26-     │
│ qr_code      │   5. Convert to PDF     │ W-0001.pdf   │
└──────────────┘                         └──────────────┘
     repeated for every row in the CSV
```

---

## 1. Template Selection

The `type` column in the CSV determines which PowerPoint file gets used:

| Type | Template file | What changes |
|------|--------------|--------------|
| `winner` | `General_Event_Winner_Template.pptx` | Name, position (rank), event name |
| `participant` | `General_Event_Participant_Template.pptx` | Name, event name |
| `judge` / `guest` | `Guest_Template.pptx` | Name, event name |

Each template was designed by the fest's design team. They contain placeholder text like `<<name>>`, `<<position>>`, and `<<event_name>>` that the script replaces with actual data.

---

## 2. Placeholder Replacement (the tricky part)

### The problem

PowerPoint doesn't store text the way you'd expect. Internally, a `.pptx` is a zip of XML files. A single visible string like `<<name>>` can get split across multiple XML "runs":

```xml
<!-- You typed <<name>> but PowerPoint stored it as: -->
<a:r><a:t>&lt;&lt;na</a:t></a:r>
<a:r><a:t>me&gt;&gt;</a:t></a:r>
```

If you search inside each run for `<<name>>`, you find nothing. Neither run contains the complete tag.

### How we handle it

The `replace_text_in_slide()` function does a two-pass approach:

**Pass 1 (fast path):** Check if any single run contains the full tag. If yes, replace it directly. This preserves the exact font/size/color formatting on that run.

**Pass 2 (fallback):** If the tag is fragmented across runs, read the full paragraph text (`paragraph.text`), do the replacement, write the result into `runs[0]`, and blank out all subsequent runs (`run.text = ""`). This fixes the split but inherits formatting from the first run only.

### Nested shapes

Templates sometimes put text inside grouped shapes or table cells. Standard `slide.shapes` iteration misses those. The `get_all_text_frames()` function walks the full shape tree recursively:

```python
def get_all_text_frames(shapes):
    text_frames = []
    for shape in shapes:
        if shape.has_text_frame:
            text_frames.append(shape.text_frame)
        if shape.shape_type == 6:  # group shape
            text_frames.extend(get_all_text_frames(shape.shapes))
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text_frame:
                        text_frames.append(cell.text_frame)
    return text_frames
```

---

## 3. Long Name Handling

Templates are designed with sample names like "John Doe". Real data has names like *"Karri Nagendra Sai Jaya Rami Reddy"* (36 characters). Without handling, the name either wraps to a second line (breaking the layout) or clips off the edge of the slide.

Two defenses:

**Disable line wrapping:**
```python
text_frame.word_wrap = False
```

**Scale the font proportionally:**
```python
if len(name) > 20:
    original_pt = run.font.size.pt or 50.58
    new_pt = max(28.0, original_pt * (20.0 / len(name)))
    run.font.size = Pt(new_pt)
```

The formula shrinks the font inversely with name length but never goes below 28pt (still readable when printed).

---

## 4. QR Code Generation

Every certificate gets a QR code pointing to its verification URL. The `insert_qr_code()` function:

**Generates the QR matrix** with high error correction (Level H — 30% recovery). This means the QR stays scannable even if part of it is obscured or printed at a weird angle.

**Strips the white background.** A stock QR code has a white square behind it. Slapping that on a textured or gradient certificate background looks terrible. We convert to RGBA and zero out white pixels:

```python
img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
pix = img.load()
for y in range(img.size[1]):
    for x in range(img.size[0]):
        r, g, b, a = pix[x, y]
        if r > 200 and g > 200 and b > 200:
            pix[x, y] = (255, 255, 255, 0)  # transparent
```

**Places it at exact coordinates.** Each certificate type has a calibrated QR position (in typographical points, converted to inches for the PPTX API):

| Type | left (pt) | top (pt) | size (pt) |
|------|-----------|----------|-----------|
| Winner | 683.15 | 127.55 | 49.25 |
| Participant | 683.15 | 145.74 | 49.25 |
| Judge/Guest | 683.15 | 89.08 | 49.25 |

These coordinates were found by inspecting the original templates with `inspect_pptx.py` and `analyze_shapes_pt.py` (in the scratch folder).

---

## 5. PDF Conversion (PowerPoint COM)

Python libraries can create and edit `.pptx` files, but none of them can render a PPTX to PDF accurately. There's no reliable cross-platform way to do this. So we use the real PowerPoint application through Windows COM automation.

```python
pythoncom.CoInitialize()
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
try:
    for pptx_path in generated_files:
        deck = powerpoint.Presentations.Open(abs_path, WithWindow=False)
        deck.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
        deck.Close()
finally:
    powerpoint.Quit()
    pythoncom.CoUninitialize()
```

Key details:

- `WithWindow=False` runs headless — no GUI window pops up.
- We open PowerPoint once and reuse it for the entire batch. Opening/closing per file is ~10x slower.
- The `try/finally` block ensures PowerPoint gets killed even if a file is corrupt. Without this, you get invisible `POWERPNT.EXE` zombie processes eating memory (see [troubleshooting](05-troubleshooting.md#1-zombie-powerpointexe-processes)).
- `pythoncom.CoInitialize()` / `CoUninitialize()` are required on Windows for COM in Python. Skip them and you get cryptic `CoInitialize has not been called` errors.

---

## Output

After `generate_certificates.py` finishes:

```
generated_certificates/
├── pptx/
│   ├── SAAVAN26-W-0001.pptx
│   ├── SAAVAN26-W-0002.pptx
│   ├── SAAVAN26-P-0001.pptx
│   ├── SAAVAN26-P-0002.pptx
│   └── SAAVAN26-J-0001.pptx
└── pdf/
    ├── SAAVAN26-W-0001.pdf   ← these get emailed
    ├── SAAVAN26-W-0002.pdf
    ├── SAAVAN26-P-0001.pdf
    ├── SAAVAN26-P-0002.pdf
    └── SAAVAN26-J-0001.pdf
```

Open a few PDFs and verify:
- Name is on one line, not wrapping.
- QR code scans and opens the right verification URL.
- No leftover `<<placeholder>>` text.
- Transparent QR — no white square on the certificate background.

Then move on to [email dispatch →](03-email-dispatch.md)
