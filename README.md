# MSG → Distributor Mapper

Matches incoming `.msg` emails to a distributor by looking up the sender's
email against a reference sheet (`Distributor`, `Dist_Acc_No`, `Distributor Email`).

## How it works

1. The reference workbook lives in the repo at `data/reference.xlsx` and
   loads automatically — no need to upload it every session.
2. Upload one or more `.msg` files.
3. For each file, the sender's email is shown as plain text, and every
   matching reference row is shown as three dropdowns (Distributor Email /
   Match Dist / Match Dist Acc No) — even a single match still shows as a
   dropdown, and multiple candidates (e.g. shared domain) are all listed.
4. Each attachment shows as a file card (icon, name, type). Click it to
   open an Excel-style view: row numbers, column letters, and a tab per
   sheet.
5. Each sheet has a **Map columns** step — rename or skip columns — with
   the last-used mapping for that distributor + sheet pre-filled. Save the
   mapping to reuse it next time, or download just that sheet with the
   mapped headers applied.
6. Hit **Clear All** to wipe the session instantly.

## Privacy

Uploaded `.msg` files and attachment bytes are processed **entirely in
memory** (`io.BytesIO`) — never written to disk. The one exception is
`data/column_mappings.json`, which intentionally persists so column
mappings are remembered across sessions (see below).

## Column mapping memory

Saved mappings are stored in `data/column_mappings.json`, keyed by
`distributor::sheet name`. This file persists for the life of the running
app instance. On Streamlit Cloud, a fresh deploy (new container) resets
it unless you commit the updated file back to the repo — for now, treat
it as living memory for the current session/instance, with periodic
manual backups if you want it permanent across deploys.

## Updating the reference sheet

Replace `data/reference.xlsx` in the repo and push — the app picks it up
on the next reload. No need to touch any code.

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy on Streamlit Cloud

1. Push this repo to GitHub.
2. In Streamlit Cloud → **New app** → point to this repo / `streamlit_app.py`.
3. **Advanced settings** → set Python version to **3.12** (already pinned via
   `runtime.txt`, but double-check the dropdown — mismatched versions have
   caused segfaults with C-extension packages like `openpyxl` in past apps).

## Roadmap (iterating incrementally)

- [ ] Batch upload via `.zip`/`.rar` of `.msg` files
- [ ] Persist a JSON mapping-memory file so repeat senders auto-resolve
- [ ] Manual override UI for unmatched senders
- [ ] Support attachment type detection (xlsx / csv / xlsb) and preview
