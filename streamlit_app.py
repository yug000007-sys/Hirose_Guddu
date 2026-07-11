"""
MSG → Distributor Mapper
-------------------------
Upload a reference workbook (Distributor / Dist_Acc_No / Distributor Email)
and a batch of .msg files. The app extracts the sender email from each .msg
and matches it against the reference sheet (exact email match first, then
domain fallback), producing a downloadable Excel mapping report.
"""

import io
import re
import tempfile
from pathlib import Path

import extract_msg
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="MSG → Distributor Mapper", layout="wide")

EMAIL_RE = re.compile(r"[\w\.\-\+]+@[\w\-]+\.[\w\.\-]+")


# ----------------------------- helpers -----------------------------------

def extract_sender_email(msg_path: str) -> tuple[str, str]:
    """Return (display_sender, clean_email) from a .msg file."""
    msg = extract_msg.Message(msg_path)
    raw_sender = msg.sender or ""
    match = EMAIL_RE.search(raw_sender)
    clean_email = match.group(0).lower().strip() if match else ""
    return raw_sender, clean_email


def load_reference(file) -> pd.DataFrame:
    """Load every sheet of the reference workbook, normalize columns,
    and explode multi-email cells into one row per email."""
    wb = openpyxl.load_workbook(file, data_only=True)
    frames = []

    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [str(h).strip() if h else "" for h in rows[0]]
        df = pd.DataFrame(rows[1:], columns=headers)

        # find the relevant columns by fuzzy name match
        dist_col = next((c for c in df.columns if c.lower().startswith("distributor") and "email" not in c.lower()), None)
        email_col = next((c for c in df.columns if "email" in c.lower()), None)
        acc_col = next((c for c in df.columns if "dist_acc" in c.lower() or c.lower() == "account"), None)
        region_col = next((c for c in df.columns if c.lower() == "region"), None)

        if not dist_col or not email_col:
            continue  # sheet doesn't look like a distributor sheet, skip

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

    # explode comma-separated email lists into one row per email
    ref["Distributor Email"] = ref["Distributor Email"].astype(str)
    ref = ref.assign(**{"Distributor Email": ref["Distributor Email"].str.split(",")}).explode("Distributor Email")
    ref["Distributor Email"] = ref["Distributor Email"].str.strip().str.lower()
    ref = ref[ref["Distributor Email"] != ""]
    ref["Domain"] = ref["Distributor Email"].str.extract(r"@(.+)$")

    return ref.reset_index(drop=True)


def match_sender(sender_email: str, ref: pd.DataFrame) -> dict:
    """Exact email match first, then domain fallback."""
    if not sender_email:
        return {"Distributor": "", "Dist_Acc_No": "", "Region": "", "Match Type": "No sender email found"}

    exact = ref[ref["Distributor Email"] == sender_email]
    if not exact.empty:
        row = exact.iloc[0]
        return {
            "Distributor": row["Distributor"],
            "Dist_Acc_No": row["Dist_Acc_No"],
            "Region": row["Region"],
            "Match Type": "Exact email match",
        }

    domain = sender_email.split("@")[-1]
    dom_match = ref[ref["Domain"] == domain]
    if not dom_match.empty:
        row = dom_match.iloc[0]
        return {
            "Distributor": row["Distributor"],
            "Dist_Acc_No": row["Dist_Acc_No"],
            "Region": row["Region"],
            "Match Type": "Domain match",
        }

    return {"Distributor": "", "Dist_Acc_No": "", "Region": "", "Match Type": "Unmatched"}


# ------------------------------- UI ---------------------------------------

st.title("📧 MSG → Distributor Mapper")
st.caption("Match incoming .msg files to a distributor using the reference sheet's sender emails.")

col1, col2 = st.columns(2)
with col1:
    ref_file = st.file_uploader("Reference workbook (Distributor / Dist_Acc_No / Distributor Email)", type=["xlsx"])
with col2:
    msg_files = st.file_uploader("MSG files (multiple allowed)", type=["msg"], accept_multiple_files=True)

if ref_file and msg_files:
    ref_df = load_reference(ref_file)

    if ref_df.empty:
        st.error("Couldn't find a sheet with Distributor + Distributor Email columns. Check the reference file.")
    else:
        st.success(f"Loaded {len(ref_df)} distributor-email rows from reference sheet.")

        results = []
        with st.spinner("Processing .msg files..."):
            for uploaded in msg_files:
                with tempfile.NamedTemporaryFile(suffix=".msg", delete=False) as tmp:
                    tmp.write(uploaded.getbuffer())
                    tmp_path = tmp.name

                raw_sender, sender_email = extract_sender_email(tmp_path)
                match = match_sender(sender_email, ref_df)

                # also try to get attachment names + subject for context
                msg = extract_msg.Message(tmp_path)
                attachments = ", ".join(a.longFilename for a in msg.attachments if a.longFilename)

                results.append({
                    "MSG File": uploaded.name,
                    "Subject": msg.subject or "",
                    "Sender": raw_sender,
                    "Sender Email": sender_email,
                    "Attachments": attachments,
                    **match,
                })
                Path(tmp_path).unlink(missing_ok=True)

        result_df = pd.DataFrame(results)

        # badge-style summary
        matched = (result_df["Match Type"] != "Unmatched").sum()
        total = len(result_df)
        st.markdown(f"**Matched:** {matched} / {total}")

        def highlight_match(row):
            if row["Match Type"] == "Unmatched" or row["Match Type"] == "No sender email found":
                return ["background-color: #ffe0e0"] * len(row)
            elif row["Match Type"] == "Domain match":
                return ["background-color: #fff6d6"] * len(row)
            else:
                return ["background-color: #e2f7e2"] * len(row)

        st.dataframe(result_df.style.apply(highlight_match, axis=1), use_container_width=True)

        # download as Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result_df.to_excel(writer, index=False, sheet_name="Mapping")
        output.seek(0)

        st.download_button(
            "⬇️ Download mapping as Excel",
            data=output,
            file_name="msg_distributor_mapping.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Upload the reference workbook and at least one .msg file to begin.")
