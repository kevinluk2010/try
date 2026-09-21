# Auto-save confidential Outlook mail as PDF

**Short answer: yes, on classic Outlook for Windows — with two real caveats.**

1. Outlook has no built-in "save as PDF" for messages, so the PDF has to be
   produced indirectly (this repo saves the message as MHTML and has Word
   export it, which preserves the From/To/Sent/Subject header block).
2. If "confidential" means *rights-protected* (encrypted, "Do Not Forward",
   a Purview label that applies encryption), automatic export is a different
   question — see [Rights-protected mail](#rights-protected-mail).

---

## Which Outlook are you on?

| Client | Works? | Route |
|---|---|---|
| Classic Outlook for Windows | Yes | [VBA macro](#option-a--vba-macro-classic-outlook-for-windows) |
| New Outlook for Windows | No VBA, no COM add-ins | [Power Automate](#option-b--power-automate-any-outlook) |
| Outlook on the web | No VBA | [Power Automate](#option-b--power-automate-any-outlook) |
| Outlook for Mac | No VBA | [Power Automate](#option-b--power-automate-any-outlook) |

Classic Outlook is the one with **File → Options → Trust Center** and an
**Alt+F11** VBA editor. If Alt+F11 does nothing, you're on new Outlook.

---

## What counts as "confidential"?

Three different things get called this, and they behave differently:

| Meaning | How to detect | Exportable? |
|---|---|---|
| Outlook sensitivity flag (`Sensitivity = Confidential`) | `MailItem.Sensitivity = 3` | Yes — it's just a flag |
| Purview / AIP label named "Confidential", no encryption | `msip_labels` MAPI property | Yes |
| Purview label **with encryption**, IRM, "Do Not Forward", S/MIME encrypted | `MailItem.Permission <> 0` | Usually no, and see below |

The macro matches the first two and **skips the third by default**.
Run `WhyNotSelected` (below) on a real message to see which bucket yours lands in.

---

## Option A — VBA macro (classic Outlook for Windows)

### Install

1. **Alt+F11** → in the Project pane expand
   `Project1` → `Microsoft Outlook Objects` → double-click **ThisOutlookSession**.
2. Paste the entire contents of
   [`outlook-vba/ThisOutlookSession.vba`](outlook-vba/ThisOutlookSession.vba).
3. Edit the **CONFIG** block at the top — at minimum `SAVE_FOLDER`.
4. **Ctrl+S**, then close the editor.
5. Allow macros to run — pick one:
   - **Recommended:** sign the project. Run `SelfCert.exe` from your Office
     install folder, create a certificate, then in the VBA editor
     **Tools → Digital Signature → Choose** your cert. Then set
     **File → Options → Trust Center → Trust Center Settings → Macro Settings**
     to *Notifications for digitally signed macros only*.
   - **Quicker, weaker:** set Macro Settings to
     *Enable all macros* (this also enables any other macro that lands in your
     profile — your IT policy may forbid it, and may overwrite the setting).
6. **Restart Outlook.** `Application_Startup` only runs at launch. (To test
   without restarting, put the cursor in `HookEvents` and press **F5**.)

### Configuration

All of it is at the top of the `.vba` file:

| Setting | Default | What it does |
|---|---|---|
| `SAVE_FOLDER` | `%USERPROFILE%\Documents\Confidential PDFs` | Destination. `%VARS%` expand; created if missing. A OneDrive or UNC path works. |
| `MATCH_SENSITIVITY_FLAG` | `True` | Match the classic Confidential flag |
| `MATCH_LABEL_NAMES` | `Confidential` | `;`-separated, case-insensitive substring, so it also matches `Confidential \ Internal` |
| `MATCH_LABEL_GUIDS` | *(empty)* | For tenants whose labels don't stamp a readable name |
| `SKIP_RIGHTS_PROTECTED` | `True` | Skip encrypted / IRM mail |
| `EXPORT_ON_OPEN` | `True` | Fire when you open a message in its own window |
| `EXPORT_ON_PREVIEW` | `False` | Fire when you select it in the reading pane |
| `EXPORT_ON_ARRIVAL` | `False` | Fire when it lands in the Inbox, opened or not |
| `LOG_FILE` | *(empty)* | Set a path to log every decision — the first thing to turn on when debugging |

You asked for "whenever I open" — that's `EXPORT_ON_OPEN`, the default.
Worth knowing: it only captures mail you actually open, so anything you read
in the reading pane or never click is missed. If the real goal is *a PDF of
every confidential message*, set `EXPORT_ON_ARRIVAL = True` instead.

### Filenames and duplicates

`2026-09-21 143207 - Jane Chan - Q3 budget review.pdf`

The name is derived deterministically from received time, sender and subject,
and the macro skips a message whose PDF already exists. So re-opening the same
mail ten times still gives you one file, and this survives an Outlook restart.

### Manual commands

Two extra macros you can run from the VBA editor or pin to the ribbon
(**File → Options → Customize Ribbon → Macros**):

- **`ExportSelectedToPdf`** — exports whatever is selected, *ignoring* the
  confidential test. Use it to backfill mail that arrived before you installed this.
- **`WhyNotSelected`** — shows a message's sensitivity, permission, message
  class and raw label string, and whether the macro considers it confidential.
  Start here when something isn't being picked up.

### If it doesn't work

| Symptom | Cause |
|---|---|
| Nothing happens at all | Macro security, or Outlook wasn't restarted. Set `LOG_FILE` and check whether `hooked;` is written at startup. |
| `Word cannot open this file` / file-block error | Word's Trust Center blocks MHTML. **Word → File → Options → Trust Center → Trust Center Settings → File Block Settings** → clear *Open* for **Web Pages and Web Archives**. |
| Word windows flash open | Expected if Word was already running — the macro reuses your instance and won't close it. |
| Reading-pane preview feels sluggish | `EXPORT_ON_PREVIEW = True` runs Word on every selection change. Leave it `False`. |
| Confidential mail ignored | Run `WhyNotSelected`. Most likely it's rights-protected and being skipped by design. |

---

## Option B — Power Automate (any Outlook)

Server-side, so it works with new Outlook, the web, and your phone, and doesn't
depend on Outlook being open. It triggers **on arrival**, not on open — there's
no "user opened a message" trigger in the cloud.

1. **When a new email arrives (V3)** — Outlook 365 connector.
   Add a condition on **Importance** / a subject or header filter, or filter
   downstream on the `msip_labels` internet header.
2. **Export email (V2)** — pass the message ID; returns the message as `.eml`.
3. **Create file** (OneDrive for Business / SharePoint) — save the `.eml` into a
   staging folder.
4. **Convert file** (OneDrive for Business) — target type **PDF**.
   The underlying Graph conversion handles `.eml` and `.msg`.
5. **Create file** — write the PDF into your designated folder.
6. **Delete file** — remove the staged `.eml`.

Rights-protected mail fails at step 4 here too.

---

## Rights-protected mail

If the message is encrypted or carries a protective label, automatically writing
it to an unprotected PDF on disk removes the control the sender applied: the file
then sits outside DLP, retention, eDiscovery and revocation. Depending on your
tenant, the export will simply fail; depending on your employer, doing it
deliberately may breach policy.

`SKIP_RIGHTS_PROTECTED = True` is the default for that reason. Turning it off
won't defeat encryption — it just lets the attempt through — and it's worth
clearing with whoever owns your information-protection policy first.

Plain `Sensitivity = Confidential` carries no such restriction. It's a visual
marker, and exporting it is unremarkable.

---

## Files

Both go in the same place — **Alt+F11 → Project1 → Microsoft Outlook Objects →
ThisOutlookSession**. Pick one, not both.

| File | Lines | Use it when |
|---|---|---|
| [`outlook-vba/Minimal.vba`](outlook-vba/Minimal.vba) | ~105 | You just want: open a Confidential message → PDF in a folder. Matches the classic sensitivity flag only. |
| [`outlook-vba/ThisOutlookSession.vba`](outlook-vba/ThisOutlookSession.vba) | ~470 | You need Purview/AIP label matching, rights-protected handling, arrival/preview triggers, logging, or the diagnostic macros. |

Start with `Minimal.vba`. Move up if it doesn't catch your mail — that usually
means your "Confidential" is a Purview label rather than the sensitivity flag,
which only the full version reads.
