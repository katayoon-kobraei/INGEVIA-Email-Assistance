import json
import os
import subprocess
from datetime import datetime, timedelta

import pandas as pd
import streamlit as st

from src.config import OUTPUT_ROOT
from src.output.project_folders import list_existing_projects, discover_trabajos_years

TASK_NAME = "Email AI Assistant"

st.set_page_config(page_title="Email Assistant - Debug Dashboard", page_icon=":wrench:", layout="wide")
st.title("Email Assistant — Debug Dashboard")


@st.cache_data(ttl=30)
def load_index():
    index_path = os.path.join(OUTPUT_ROOT, "index.csv")
    if not os.path.exists(index_path):
        return pd.DataFrame()
    return pd.read_csv(index_path, encoding="utf-8-sig")


@st.cache_data(ttl=30)
def load_quarantine_details():
    """Scan every metadata.json under OUTPUT_ROOT for quarantined attachments."""
    records = []
    if not os.path.isdir(OUTPUT_ROOT):
        return pd.DataFrame(records)
    for dirpath, _, filenames in os.walk(OUTPUT_ROOT):
        if "metadata.json" not in filenames:
            continue
        try:
            with open(os.path.join(dirpath, "metadata.json"), "r", encoding="utf-8") as f:
                meta = json.load(f)
        except Exception:
            continue
        for att in meta.get("attachments", []):
            if att.get("status") == "quarantined":
                records.append({
                    "Subject": meta.get("subject", ""),
                    "Filename": att.get("filename", ""),
                    "Reason": att.get("reason", ""),
                    "Email Folder": dirpath,
                })
    return pd.DataFrame(records)


def get_scheduler_status():
    try:
        result = subprocess.run(
            [
                "powershell", "-Command",
                f"Get-ScheduledTaskInfo -TaskName '{TASK_NAME}' | "
                f"Select-Object LastRunTime, LastTaskResult, NextRunTime | ConvertTo-Json",
            ],
            capture_output=True, text=True, timeout=10,
        )
        if result.returncode != 0 or not result.stdout.strip():
            return None
        return json.loads(result.stdout)
    except Exception:
        return None


# --- Scheduler health ----------------------------------------------------

st.subheader("Scheduler status")
status = get_scheduler_status()
if status:
    cols = st.columns(3)
    cols[0].metric("Last run", str(status.get("LastRunTime", "—")))
    result_code = status.get("LastTaskResult")
    cols[1].metric("Last result", "OK" if result_code == 0 else f"Code {result_code}")
    cols[2].metric("Next run", str(status.get("NextRunTime", "—")))
else:
    st.warning(f"Could not read Task Scheduler status for task '{TASK_NAME}'.")

if st.button("Run pipeline now"):
    with st.spinner("Running pipeline..."):
        try:
            from src.pipeline import run as run_pipeline
            run_pipeline()
            st.success("Pipeline run complete.")
            load_index.clear()
            load_quarantine_details.clear()
        except Exception as e:
            st.error(f"Pipeline run failed: {e}")

st.divider()

# --- Summary metrics -------------------------------------------------------

df = load_index()

if df.empty:
    st.info("No emails processed yet (index.csv not found or empty).")
else:
    total = len(df)
    entrante = (df["Direction"] == "ENTRANTE").sum()
    saliente = (df["Direction"] == "SALIENTE").sum()
    unsorted_df = df[df["Project Folder"] == "UNSORTED"]
    quarantine_df = load_quarantine_details()

    cols = st.columns(5)
    cols[0].metric("Total processed", total)
    cols[1].metric("Inbound (ENTRANTE)", int(entrante))
    cols[2].metric("Outbound (SALIENTE)", int(saliente))
    cols[3].metric("Unsorted", len(unsorted_df))
    cols[4].metric("Quarantined attachments", len(quarantine_df))

    st.divider()

    # --- Unsorted / failed classification ---------------------------------

    st.subheader("Needs review (classification failed)")
    if unsorted_df.empty:
        st.success("Nothing unsorted right now.")
    else:
        st.dataframe(unsorted_df, use_container_width=True)

    st.divider()

    # --- Recent activity ----------------------------------------------------

    st.subheader("Recent activity")
    col1, col2, col3 = st.columns(3)
    projects = sorted(df["Project Folder"].dropna().unique().tolist())
    project_filter = col1.multiselect("Project", projects)
    direction_filter = col2.multiselect("Direction", sorted(df["Direction"].dropna().unique().tolist()))
    days_back = col3.slider("Days back", 1, 90, 14)

    filtered = df.copy()
    filtered["Date"] = pd.to_datetime(filtered["Date"], errors="coerce")
    cutoff = datetime.now() - timedelta(days=days_back)
    filtered = filtered[filtered["Date"] >= cutoff]

    if project_filter:
        filtered = filtered[filtered["Project Folder"].isin(project_filter)]
    if direction_filter:
        filtered = filtered[filtered["Direction"].isin(direction_filter)]

    st.dataframe(filtered.sort_values("Date", ascending=False), use_container_width=True)

    st.divider()

    # --- Quarantine viewer -----------------------------------------------

    st.subheader("Quarantined attachments")
    if quarantine_df.empty:
        st.success("Nothing quarantined.")
    else:
        st.dataframe(quarantine_df, use_container_width=True)

    st.divider()

    # --- Known projects -----------------------------------------------------

    st.subheader("Known project folders")
    all_years = discover_trabajos_years(OUTPUT_ROOT)
    project_rows = [
        {"Year": year, "Project Folder": name}
        for year in all_years
        for name in list_existing_projects(OUTPUT_ROOT, [year])
    ]
    st.write(f"{len(project_rows)} project folder(s) on record across {len(all_years)} year(s):")
    st.dataframe(pd.DataFrame(project_rows), use_container_width=True)