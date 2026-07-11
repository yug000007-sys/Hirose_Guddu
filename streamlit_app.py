"""
MSG -> Distributor Mapper (v3)
--------------------------------
- Reference workbook bundled in repo (data/reference.xlsx), no re-upload needed.
- .msg files processed entirely in memory. Nothing written to disk except the
  column-mapping memory file (data/column_mappings.json), which is meant to persist.
- Match results shown as three independent dropdowns (Distributor Email,
  Match Dist, Match Dist Acc No), always dropdowns even with a single option.
- Attachments shown as file cards (icon + name + type). Clicking opens an
  Excel-style sheet view (row numbers, column letters, sheet tabs).
- Each sheet has a column-mapping step (rename/skip columns), remembered per
  distributor + sheet name, with a "download this sheet only" export.
"""

import io
import json
import re
from pathlib import Path

import extract_msg
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="MSG -> Distributor Mapper", layout="wide")

EMAIL_RE = re.compile(r"[\w\.\-\+]+@[\w\-]+\.[\w\.\-]+")
BUNDLED_REFERENCE_PATH = Path(__file__).parent / "data" / "reference.xlsx"
MAPPING_MEMORY_PATH = Path(__file__).parent / "data" / "column_mappings.json"

MIME_TYPES = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
    ".xlsb": "application/vnd.ms-excel.sheet.binary.macroenabled.12",
    ".csv": "text/csv",
}

EXCEL_ICON_COLORS = {".xlsx": "#1D6F42", ".xls": "#1D6F42", ".xlsb": "#1D6F42", ".csv": "#217346"}


def guess_mime(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    return MIME_TYPES.get(ext, "application/octet-stream")


def col_letter(n: int) -> str:
    """0-indexed column number -> Excel-style letter (A, B, ..., Z, AA, ...)."""
    letters = ""
    n += 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ----------------------------- reference matching -----------------------------

def extract_sender_email(file_bytes: bytes) -> tuple[str, str, str, list[tuple[str, bytes]]]:
    bio = io.BytesIO(file_bytes)
    msg = extract_msg.Message(bio)
    raw_sender = msg.sender or ""
    match = EMAIL_RE.search(raw_sender)
    clean_email = match.group(0).lower().strip() if match else ""
    subject = msg.subject or ""
    attachments = [(a.longFilename, a.data) for a in msg.attachments if a.longFilename and a.data]
    msg.close()
    return raw_sender, clean_email, subject, attachments


@st.cache_data(show_spinner=False)
def load_reference(file_bytes: bytes) -> pd.DataFrame:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    frames = []

    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [str(h).strip() if h else "" for h in rows[0]]
        df = pd.DataFrame(rows[1:], columns=headers)

        dist_col = next((c for c in df.columns if c.lower().startswith("distributor") and "email" not in c.lower()), None)
        email_col = next((c for c in df.columns if "email" in c.lower()), None)
        acc_col = next((c for c in df.columns if "dist_acc" in c.lower() or c.lower() == "account"), None)
        region_col = next((c for c in df.columns if c.lower() == "region"), None)

        if not dist_col or not email_col:
            continue

        keep = df[[dist_col, email_col]].copy()
        keep.columns = ["Distributor", "Distributor Email"]
        keep["Dist_Acc_No"] = df[acc_col] if acc_col else ""
        keep["Region"] = df[region_col] if region_col else ""
        keep["Source Sheet"] = ws.title
        frames.append(keep)

    if not frames:
        return pd.DataFrame(columns=["Distributor", "Distributor Email", "Dist_Acc_No", "Region", "Source Sheet"])

    ref = pd.concat(frames, ignore_index=True)
    ref = ref.dropna(subset=["Distributor Email"])
    ref["Distributor Email"] = ref["Distributor Email"].astype(str)
    ref = ref.assign(**{"Distributor Email": ref["Distributor Email"].str.split(",")}).explode("Distributor Email")
    ref["Distributor Email"] = ref["Distributor Email"].str.strip().str.lower()
    ref = ref[ref["Distributor Email"] != ""]
    ref["Domain"] = ref["Distributor Email"].str.extract(r"@(.+)$")

    return ref.reset_index(drop=True)


def get_candidates(sender_email: str, ref: pd.DataFrame) -> list[dict]:
    """Return every matching reference row (exact matches first, then domain
    matches), each as a dict. Empty list if nothing matches."""
    if not sender_email:
        return []

    exact = ref[ref["Distributor Email"] == sender_email]
    domain = sender_email.split("@")[-1]
    dom_match = ref[(ref["Domain"] == domain) & (ref["Distributor Email"] != sender_email)]

    candidates = []
    for _, row in exact.iterrows():
        candidates.append({
            "Distributor": row["Distributor"], "Dist_Acc_No": row["Dist_Acc_No"],
            "Region": row["Region"], "Distributor Email": row["Distributor Email"],
            "Match Type": "Exact email match",
        })
    for _, row in dom_match.iterrows():
        candidates.append({
            "Distributor": row["Distributor"], "Dist_Acc_No": row["Dist_Acc_No"],
            "Region": row["Region"], "Distributor Email": row["Distributor Email"],
            "Match Type": "Domain match",
        })
    return candidates


# ----------------------------- sheet reading -----------------------------

def _detect_header_row(raw: pd.DataFrame, scan_rows: int = 15) -> int:
    """Pick the most likely header row within the first `scan_rows` rows.
    A real header row is mostly filled AND made of actual text labels
    (e.g. 'Location', 'Report Date') rather than short numeric/flag values
    (e.g. a '1,2,3...' index row, or a 'Y/N' row) - those can be just as
    full as the real header, so fullness alone isn't enough."""
    total_cols = raw.shape[1] or 1
    for i in range(min(scan_rows, len(raw))):
        row = raw.iloc[i]
        non_null = row.notna()
        fill_ratio = non_null.sum() / total_cols
        if fill_ratio < 0.5:
            continue
        values = [str(v) for v in row[non_null]]
        avg_len = sum(len(v) for v in values) / len(values) if values else 0
        if avg_len >= 4:
            return i

    # fallback: most-filled row in the scan window
    best_idx, best_count = 0, -1
    for i in range(min(scan_rows, len(raw))):
        count = raw.iloc[i].notna().sum()
        if count > best_count:
            best_idx, best_count = i, count
    return best_idx


def read_all_sheets_raw(filename: str, file_bytes: bytes) -> dict[str, pd.DataFrame] | None:
    """Read every sheet with no header assumption (header=None), so the
    caller can detect/choose the real header row."""
    ext = Path(filename).suffix.lower()
    bio = io.BytesIO(file_bytes)
    try:
        if ext == ".csv":
            return {"Sheet1": pd.read_csv(bio, header=None, dtype=str)}
        elif ext == ".xlsb":
            return pd.read_excel(bio, sheet_name=None, engine="pyxlsb", header=None, dtype=str)
        elif ext in (".xlsx", ".xls"):
            return pd.read_excel(bio, sheet_name=None, header=None, dtype=str)
    except Exception:
        return None
    return None


def apply_header_row(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    """Slice a header=None dataframe into a proper dataframe using the given
    row index as column headers."""
    headers = raw.iloc[header_row].tolist()
    headers = [
        f"Column {i+1}" if (h is None or (isinstance(h, float) and pd.isna(h)) or str(h).strip() == "")
        else str(h).strip()
        for i, h in enumerate(headers)
    ]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)
    df = clean_numeric_strings(df)
    return df


def clean_numeric_strings(df: pd.DataFrame) -> pd.DataFrame:
    """Cells were read as strings, so numbers like 84480*5.86 come back as
    '495052.80000000005' (binary float noise). Round any numeric-looking
    cell to 9 decimals to strip that noise while keeping real precision
    (e.g. a genuine 4-decimal unit cost stays untouched)."""
    df = df.copy()
    for col in df.columns:
        series = df[col]

        def _clean(v):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                return v
            s = str(v).strip()
            if s == "":
                return v
            try:
                f = float(s)
            except (ValueError, TypeError):
                return v
            r = round(f, 9)
            if r == int(r):
                return str(int(r))
            return str(r)

        df[col] = series.map(_clean)
    return df


# ----------------------------- mapping memory -----------------------------

def load_mapping_memory() -> dict:
    if MAPPING_MEMORY_PATH.exists():
        try:
            return json.loads(MAPPING_MEMORY_PATH.read_text())
        except Exception:
            return {}
    return {}


def save_mapping_memory(memory: dict):
    MAPPING_MEMORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    MAPPING_MEMORY_PATH.write_text(json.dumps(memory, indent=2, ensure_ascii=False))


def mapping_key(distributor: str, sheet_name: str) -> str:
    return f"{distributor or 'Unknown'}::{sheet_name}"


# --------------------------- session state ---------------------------------

if "results" not in st.session_state:
    st.session_state.results = None
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "expanded_attachment" not in st.session_state:
    st.session_state.expanded_attachment = None
if "mapping_memory" not in st.session_state:
    st.session_state.mapping_memory = load_mapping_memory()


def clear_all():
    st.session_state.results = None
    st.session_state.expanded_attachment = None
    st.session_state.uploader_key += 1


# ------------------------------- UI ---------------------------------------

st.title("MSG -> Distributor Mapper")
st.caption("Upload .msg files to match senders to distributors, preview attachments, and map/export sheet columns. "
           "Msg files stay in memory only; nothing is written to disk except saved column mappings.")

top_l, top_r = st.columns([4, 1])
with top_r:
    st.button("Clear All", on_click=clear_all, use_container_width=True, type="secondary")

with st.expander("Reference sheet options (using bundled sheet by default)"):
    override_ref = st.file_uploader(
        "Override with a different reference workbook for this session only",
        type=["xlsx"], key=f"ref_override_{st.session_state.uploader_key}",
    )

if override_ref is not None:
    ref_bytes = override_ref.getvalue()
elif BUNDLED_REFERENCE_PATH.exists():
    ref_bytes = BUNDLED_REFERENCE_PATH.read_bytes()
else:
    ref_bytes = None

msg_files = st.file_uploader(
    "MSG files (multiple allowed)", type=["msg"], accept_multiple_files=True,
    key=f"msg_uploader_{st.session_state.uploader_key}",
)

if ref_bytes is None:
    st.error("No reference sheet found. Add one at data/reference.xlsx in the repo, or upload an override above.")
elif msg_files:
    ref_df = load_reference(ref_bytes)

    if ref_df.empty:
        st.error("Couldn't find a sheet with Distributor + Distributor Email columns in the reference workbook.")
    else:
        results = []
        with st.spinner("Processing .msg files (in memory, nothing saved to disk)..."):
            for uploaded in msg_files:
                file_bytes = uploaded.getvalue()
                raw_sender, sender_email, subject, attachments = extract_sender_email(file_bytes)
                candidates = get_candidates(sender_email, ref_df)

                results.append({
                    "MSG File": uploaded.name,
                    "Subject": subject,
                    "Sender": raw_sender,
                    "Sender Email": sender_email,
                    "Candidates": candidates,
                    "Attachments Data": attachments,
                })

        st.session_state.results = results

if st.session_state.results:
    st.divider()

    for f_idx, item in enumerate(st.session_state.results):
        st.markdown(f"**{item['MSG File']}**")
        st.write(f"Sender email - {item['Sender Email'] or '\u2014'}")

        candidates = item["Candidates"]
        if candidates:
            c1, c2, c3 = st.columns(3)
            emails = [c["Distributor Email"] for c in candidates]
            dists = [c["Distributor"] for c in candidates]
            accs = [c["Dist_Acc_No"] for c in candidates]
            with c1:
                st.selectbox("Distributor Email", emails, key=f"email_dd_{f_idx}")
            with c2:
                st.selectbox("Match Dist", dists, key=f"dist_dd_{f_idx}")
            with c3:
                st.selectbox("Match Dist Acc No", accs, key=f"acc_dd_{f_idx}")
            selected_distributor = dists[0]
        else:
            c1, c2, c3 = st.columns(3)
            with c1:
                st.selectbox("Distributor Email", ["No match found"], key=f"email_dd_{f_idx}", disabled=True)
            with c2:
                st.selectbox("Match Dist", ["No match found"], key=f"dist_dd_{f_idx}", disabled=True)
            with c3:
                st.selectbox("Match Dist Acc No", ["No match found"], key=f"acc_dd_{f_idx}", disabled=True)
            selected_distributor = "Unknown"

        attachments = item["Attachments Data"]
        if attachments:
            att_cols = st.columns(min(len(attachments), 4))
            for a_idx, (att_name, att_bytes) in enumerate(attachments):
                ext = Path(att_name).suffix.lower()
                card_key = f"card_{f_idx}_{a_idx}"
                with att_cols[a_idx % len(att_cols)]:
                    clicked = st.button(
                        f"\U0001F4C4 {att_name}\n{ext.lstrip('.').upper()} File",
                        key=card_key, use_container_width=True,
                    )
                    if clicked:
                        current = st.session_state.expanded_attachment
                        target = (f_idx, a_idx)
                        st.session_state.expanded_attachment = None if current == target else target

            if st.session_state.expanded_attachment and st.session_state.expanded_attachment[0] == f_idx:
                _, a_idx = st.session_state.expanded_attachment
                att_name, att_bytes = attachments[a_idx]
                raw_sheets = read_all_sheets_raw(att_name, att_bytes)

                if raw_sheets is None:
                    st.warning(f"Could not read {att_name} as a spreadsheet.")
                else:
                    sheet_names = list(raw_sheets.keys())
                    tabs = st.tabs(sheet_names)

                    for tab, sheet_name in zip(tabs, sheet_names):
                        with tab:
                            raw_df = raw_sheets[sheet_name]
                            detected_row = _detect_header_row(raw_df)

                            hr_key = f"header_row_{f_idx}_{a_idx}_{sheet_name}"
                            if hr_key not in st.session_state:
                                st.session_state[hr_key] = detected_row

                            st.caption("Click a row below to use it as the header row (auto-detected row is pre-selected).")

                            preview_rows = min(15, len(raw_df))
                            preview_df = raw_df.iloc[:preview_rows].copy()
                            preview_df.columns = [col_letter(i) for i in range(preview_df.shape[1])]
                            preview_df.index = range(preview_rows)

                            sel_key = f"header_select_{f_idx}_{a_idx}_{sheet_name}"
                            event = st.dataframe(
                                preview_df, use_container_width=True, height=280,
                                on_select="rerun", selection_mode="single-row", key=sel_key,
                            )

                            selected_rows = event.selection.rows if event and event.selection else []
                            if selected_rows:
                                st.session_state[hr_key] = selected_rows[0]

                            header_row = st.session_state[hr_key]
                            st.write(f"Using row {header_row} as header.")

                            df = apply_header_row(raw_df, int(header_row))
                            display_df = df.copy()
                            display_df.columns = [f"{col_letter(i)}: {c}" for i, c in enumerate(df.columns)]
                            display_df.index = range(header_row + 2, header_row + 2 + len(display_df))
                            st.dataframe(display_df, use_container_width=True, height=280)

                            mkey = mapping_key(selected_distributor, sheet_name)
                            saved_map = st.session_state.mapping_memory.get(mkey, {})

                            with st.expander(f"Map columns - {sheet_name}"):
                                new_names = []
                                for i, col in enumerate(df.columns):
                                    saved_name = saved_map.get(str(i), saved_map.get(col, col))
                                    new_name = st.text_input(
                                        f"{col_letter(i)}: {col}", value=saved_name,
                                        key=f"map_{f_idx}_{a_idx}_{sheet_name}_{i}",
                                    )
                                    new_names.append(new_name)
                                new_map = {str(i): name for i, name in enumerate(new_names)}

                                mc1, mc2 = st.columns(2)
                                with mc1:
                                    if st.button(f"Save mapping for {selected_distributor}", key=f"save_map_{f_idx}_{a_idx}_{sheet_name}"):
                                        st.session_state.mapping_memory[mkey] = new_map
                                        save_mapping_memory(st.session_state.mapping_memory)
                                        st.success("Mapping saved.")

                                with mc2:
                                    mapped_df = df.copy()
                                    mapped_df.columns = new_names
                                    sheet_out = io.BytesIO()
                                    with pd.ExcelWriter(sheet_out, engine="openpyxl") as writer:
                                        mapped_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
                                    sheet_out.seek(0)
                                    st.download_button(
                                        "Download this sheet",
                                        data=sheet_out,
                                        file_name=f"{Path(att_name).stem}_{sheet_name}_mapped.xlsx",
                                        mime=MIME_TYPES[".xlsx"],
                                        key=f"dl_sheet_{f_idx}_{a_idx}_{sheet_name}",
                                    )

                    st.download_button(
                        f"Download original {att_name}",
                        data=att_bytes, file_name=att_name, mime=guess_mime(att_name),
                        key=f"dl_orig_{f_idx}_{a_idx}",
                    )
        else:
            st.caption("No attachments found.")

        st.divider()

    # full mapping export (sender/match summary only, no attachment bytes)
    export_rows = []
    for f_idx, item in enumerate(st.session_state.results):
        candidates = item["Candidates"]
        best = candidates[0] if candidates else {"Distributor": "", "Dist_Acc_No": "", "Region": "", "Distributor Email": "", "Match Type": "Unmatched"}
        export_rows.append({
            "MSG File": item["MSG File"], "Sender Email": item["Sender Email"],
            "Attachments": ", ".join(n for n, _ in item["Attachments Data"]),
            **best,
        })
    export_df = pd.DataFrame(export_rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Mapping")
    output.seek(0)

    st.download_button(
        "Download full mapping summary as Excel",
        data=output, file_name="msg_distributor_mapping.xlsx",
        mime=MIME_TYPES[".xlsx"],
    )
else:
    st.info("Upload .msg files to see match results.")
