# Troubleshooting

Things that have gone wrong during real batch runs, why they happened, and how to fix them.

---

## 1. Zombie PowerPoint.exe Processes

**Symptom:** After running `generate_certificates.py`, files are locked (`PermissionError: [WinError 32]`). Or the next run fails with a COM error like `0x80010105`.

**Why:** If the script crashes before calling `powerpoint.Quit()`, the PowerPoint process stays alive invisibly in the background. It holds file locks on any presentations it had open. Stack up a few of these and Windows starts running out of desktop heap.

**Fix:** Kill all orphaned PowerPoint processes from PowerShell:
```powershell
Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
```

Then re-run the script. The code already uses `try/finally` to clean up, but if Python itself segfaults or gets killed, the `finally` block doesn't run.

---

## 2. Placeholder Tags Not Getting Replaced

**Symptom:** The generated PDF still shows `<<name>>` or `<<event_name>>` instead of actual values.

**Why:** PowerPoint internally splits text across multiple XML "runs." If someone edited the template and touched the placeholder text, PowerPoint may have fragmented `<<name>>` into `<<na` + `me>>` across separate runs. The simple `run.text.replace()` approach finds nothing.

**Check:** Open the template in `inspect_pptx.py` and look at the raw text per shape. If the placeholder appears correct in the output, the issue is elsewhere (wrong tag spelling, case mismatch). If it looks fragmented, the two-pass replacement in `replace_text_in_slide()` should handle it — but only if the placeholder appears somewhere in the full `paragraph.text`.

**Fix:** Open the template in PowerPoint, select the placeholder text, delete it entirely, and retype it in one go without backspacing or reformatting. Save. This forces PowerPoint to store it as a single run.

---

## 3. Windows Console Encoding Crash

**Symptom:**
```
UnicodeEncodeError: 'charmap' codec can't encode character '\u2705'
```

**Why:** Windows cmd.exe and older PowerShell versions use legacy code pages (CP1252). The script prints emoji status indicators (🚀, ✅, ❌) that can't be encoded in CP1252.

**Fix:** Already handled. Every script includes this at the top:
```python
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        pass
```

If you still hit this, run from Windows Terminal (which defaults to UTF-8) instead of the legacy `cmd.exe`.

---

## 4. SES "Email address is not verified"

**Symptom:**
```
MessageRejected: Email address is not verified.
The following identities failed the check in region AP-SOUTH-1: user@gmail.com
```

**Why:** Your SES account is still in sandbox mode. In sandbox, you can only send to email addresses that have been individually verified in the SES console.

**Fix:** Either:
- Verify the recipient address in AWS SES console → Verified Identities.
- Or request production access (SES console → Account dashboard → Request production access). This removes the restriction entirely.

During development, use `--to` to send only to your own verified address.

---

## 5. SES Throttling

**Symptom:**
```
454 Throttling: Maximum sending rate exceeded
```

**Why:** SES has a per-second sending rate limit. Fresh accounts start at 1-14 emails/second. Blasting 500 emails in a tight loop without pauses exceeds this.

**Fix:** The `send_emails.py` script doesn't have a built-in delay by default. If you hit throttling, add a sleep between sends by modifying the loop, or pace your batch. Alternatively, request a sending rate increase from the SES console.

---

## 6. Long Names Breaking the Certificate Layout

**Symptom:** A name wraps to two lines, pushing the signature block down and overlapping with the border.

**Why:** The text frame's default behavior is to wrap text. Combined with a long name and a large font size, it overflows.

**Fix:** Already handled in the code. `word_wrap` is disabled and font size scales down for names over 20 characters. But if you see this on a new template, check:
1. The template's text frame is wide enough for ~40 characters at the minimum font size (28pt).
2. The `replace_text_in_slide()` function's name detection is matching the right tag (`<<name>>`).

---

## 7. QR Code Has a White Square Background

**Symptom:** The QR code renders correctly but has an ugly white rectangle behind it, clashing with the certificate's colored/textured background.

**Why:** The Pillow alpha transparency mask didn't run, or the threshold is wrong.

**Check:** Open the intermediate QR image (before it gets deleted — comment out the `os.remove()` call temporarily) and verify it has a transparent background. Open it in an image editor that shows transparency as a checkerboard.

**Fix:** The pixel-level white stripping checks `r > 200 and g > 200 and b > 200`. If the QR background is slightly off-white (e.g., light gray), lower the threshold.

---

## 8. COM Error: "CoInitialize has not been called"

**Symptom:**
```
pywintypes.com_error: (-2147221008, 'CoInitialize has not been called.', None, None)
```

**Why:** COM automation requires each thread to initialize its own COM apartment. If you're running the conversion in a worker thread without calling `pythoncom.CoInitialize()`, this happens.

**Fix:** Call `pythoncom.CoInitialize()` at the start of every thread that touches COM, and `pythoncom.CoUninitialize()` when done. The main-thread code already does this.

---

## 9. PDF Not Found During Email Sending

**Symptom:**
```
⚠️ Warning: PDF file 'generated_certificates/pdf/SAAVAN26-W-0001.pdf' not found.
Email will be sent without attachment.
```

**Why:** Either `generate_certificates.py` wasn't run first, or it failed partway through and didn't produce all PDFs.

**Fix:** Run `generate_certificates.py` first. Check the `generated_certificates/pdf/` directory and compare file count to CSV row count. If some are missing, check the generation logs for which files failed.
