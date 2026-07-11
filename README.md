# MSG → Distributor Mapper

Matches incoming `.msg` emails to a distributor by looking up the sender's
email against a reference sheet (`Distributor`, `Dist_Acc_No`, `Distributor Email`).

## How it works

1. The reference workbook lives in the repo at `data/reference.xlsx` and
   loads automatically — no need to upload it every session. (You can
   still upload a one-off override sheet via the expander if needed.)
2. Upload one or more `.msg` files.
3. The app extracts each sender's email and matches it:
   - **Exact match** against the reference email column (comma-separated
     multi-email cells are split and each checked individually).
   - **Domain fallback** (e.g. `@taisei-musen.com.hk`) if no exact match.
4. Review Matched / Unmatched tables (Distributor + Dist_Acc_No shown for
   every match) and download the full mapping as Excel.
5. Hit **Clear All** to wipe the session instantly.

## Privacy

Uploaded `.msg` files and their bytes are processed **entirely in memory**
(`io.BytesIO`) — nothing is ever written to disk or to a cache file.
Streamlit's own upload buffer holds the file only for the life of the
browser session; closing the tab or hitting **Clear All** discards it.

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
