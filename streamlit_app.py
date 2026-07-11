"""
MSG -> Distributor Mapper (v2)
--------------------------------
- Reference workbook is bundled in the repo (data/reference.xlsx) so you
  don't need to re-upload it every time. You can still override it for a
  one-off run with a different sheet.
- .msg files are processed entirely in memory (io.BytesIO). Nothing is
  written to disk or cache at any point.
- "Clear All" button wipes the in-memory session state immediately.
"""

import io
import re
from pathlib import Path

import extract_msg
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="MSG -> Distributor Mapper", layout="wide")

EMAIL_RE = re.compile(r"[\w\.\-\+]+@[\w\-]+\.[\w\.\-]+")
BUNDLED_REFERENCE_PATH = Path(__file__).parent / "data" / "reference.xlsx"

MIME_TYPES = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xls": "application/vnd.ms-excel",
    ".xlsb": "application/vnd.ms-excel.sheet.binary.macroenabled.12",
    ".csv": "text/csv",
}


def guess_mime(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    return MIME_TYPES.get(ext, "application/octet-stream")


# ----------------------------- helpers -----------------------------------

def extract_sender_email(file_bytes: bytes) -> tuple[str, str, str, list[tuple[str, bytes]]]:
    """Return (display_sender, clean_email, subject, attachments) from
    in-memory .msg bytes. Nothing touches disk. attachments is a list of
    (filename, file_bytes) tuples."""
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
    """Load every sheet of the reference workbook, normalize columns,
    and explode multi-email cells into one row per email."""
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


def match_sender(sender_email: str, ref: pd.DataFrame) -> dict:
    if not sender_email:
        return {"Distributor": "", "Dist_Acc_No": "", "Region": "", "Distributor Email": "", "Match Type": "No sender email found"}

    exact = ref[ref["Distributor Email"] == sender_email]
    if not exact.empty:
        row = exact.iloc[0]
        return {"Distributor": row["Distributor"], "Dist_Acc_No": row["Dist_Acc_No"], "Region": row["Region"], "Distributor Email": row["Distributor Email"], "Match Type": "Exact email match"}

    domain = sender_email.split("@")[-1]
    dom_match = ref[ref["Domain"] == domain]
    if not dom_match.empty:
        row = dom_match.iloc[0]
        return {"Distributor": row["Distributor"], "Dist_Acc_No": row["Dist_Acc_No"], "Region": row["Region"], "Distributor Email": row["Distributor Email"], "Match Type": "Domain match"}

    return {"Distributor": "", "Dist_Acc_No": "", "Region": "", "Distributor Email": "", "Match Type": "Unmatched"}


# --------------------------- session state ---------------------------------

if "results" not in st.session_state:
    st.session_state.results = None
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


def clear_all():
    st.session_state.results = None
    st.session_state.uploader_key += 1  # forces file_uploader widgets to reset


# ------------------------------- UI ---------------------------------------

st.title("MSG -> Distributor Mapper")
st.caption("Reference sheet is preloaded from the repo. Upload .msg files to match senders to distributors. "
           "Nothing is written to disk or cache — everything stays in memory for this session only.")

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
    ref_source_label = f"Uploaded override: {override_ref.name}"
elif BUNDLED_REFERENCE_PATH.exists():
    ref_bytes = BUNDLED_REFERENCE_PATH.read_bytes()
    ref_source_label = "Bundled reference sheet (data/reference.xlsx)"
else:
    ref_bytes = None
    ref_source_label = None

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
                file_bytes = uploaded.getvalue()  # stays in RAM only
                raw_sender, sender_email, subject, attachments = extract_sender_email(file_bytes)
                match = match_sender(sender_email, ref_df)

                results.append({
                    "MSG File": uploaded.name,
                    "Subject": subject,
                    "Sender": raw_sender,
                    "Sender Email": sender_email,
                    "Attachments": ", ".join(name for name, _ in attachments),
                    "Attachments Data": attachments,
                    **match,
                })

        st.session_state.results = pd.DataFrame(results)

if st.session_state.results is not None:
    result_df = st.session_state.results
    st.divider()

    for idx, row in result_df.iterrows():
        is_matched = row["Match Type"] in ["Exact email match", "Domain match"]
        label = f"{row['MSG File']}" + (f" — {row['Distributor']}" if is_matched else " — Unmatched")

        with st.popover(label, use_container_width=True):
            st.write(
                f"Sender email - {row['Sender Email'] or '—'}  |  "
                f"Distributor Email - {row['Distributor Email'] if is_matched else '—'}  |  "
                f"Match Dist - {row['Distributor'] if is_matched else '—'}  |  "
                f"Match Dist Acc No - {row['Dist_Acc_No'] if is_matched else '—'}"
            )
            if not is_matched:
                st.caption(f"⚠ {row['Match Type']}")

            attachments = row["Attachments Data"]
            if attachments:
                st.write("Attachments:")
                cols = st.columns(len(attachments))
                for i, (att_name, att_bytes) in enumerate(attachments):
                    with cols[i]:
                        st.download_button(
                            f"⬇ {att_name}",
                            data=att_bytes,
                            file_name=att_name,
                            mime=guess_mime(att_name),
                            key=f"att_{idx}_{i}",
                        )
            else:
                st.caption("No attachments found.")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.drop(columns=["Attachments Data"]).to_excel(writer, index=False, sheet_name="Mapping")
    output.seek(0)

    st.download_button(
        "Download mapping as Excel",
        data=output,
        file_name="msg_distributor_mapping.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload .msg files to see match results.")
