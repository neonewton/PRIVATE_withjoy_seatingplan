import streamlit as st
import pandas as pd
import io
from datetime import datetime

SAMPLE_CSV_URL = "https://raw.githubusercontent.com/neonewton/PRIVATE_withjoy_seatingplan/main/guest-list.csv"


# -----------------------------
# Helper: clean "No"
# -----------------------------
def clean_no(value):
    if pd.isna(value):
        return value
    return "" if "no" in str(value).lower() else value


# -----------------------------
# Core Seating Plan Generator
# -----------------------------
def generate_seating_plan(df, table_size=10):

    # --- CLEAN COLUMN NAMES ---
    df.columns = df.columns.str.strip()

    # --- COMBINE NAME ---
    df["full_name"] = (
        df["first name"].fillna("") + " " + df["last name"].fillna("")
    ).str.strip()

    # --- RSVP FILTERING ---
    rsvp = df["rsvp"].astype(str)
    declined_mask = rsvp.str.contains("Regretfully Decline", case=False, na=False)
    blank_mask = rsvp.str.strip().eq("") | df["rsvp"].isna()
    attending_mask = ~(declined_mask | blank_mask)

    attending = df[attending_mask].copy()
    pending = df[blank_mask].copy()
    declined = df[declined_mask].copy()

    # --- TAG CLEANING ---
    attending["tag_group_raw"] = (
        attending["tags"]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.replace(r"\s+", "", regex=True)
        .str.strip()
    )

    no_tag_mask = attending["tag_group_raw"].isin(["", "nan", "none"])
    pending_tags = attending[no_tag_mask].copy()
    attending = attending[~no_tag_mask].copy()

    attending["tag_group"] = attending["tag_group_raw"]

    # -----------------------------
    # TABLE ASSIGNMENT
    # -----------------------------
    table_number_by_index = {}
    next_table_number = 1

    tag_groups = sorted(attending["tag_group"].unique())

    for tg in tag_groups:

        group_df = attending[attending["tag_group"] == tg]
        group_indices = list(group_df.index)
        group_size = len(group_indices)

        if group_size == 0:
            continue

        # Allow 11-seater only if exactly 11 guests
        local_cap = table_size + 1 if group_size == table_size + 1 else table_size

        # Split into party or single
        party_str = group_df["party"].astype(str).str.strip()
        has_party = party_str != ""
        parties_df = group_df[has_party]
        singles_df = group_df[~has_party]

        tables_for_tag = []

        # --- PARTIES FIRST ---
        for party_id, pgroup in parties_df.groupby("party"):
            idxs = list(pgroup.index)
            size = len(idxs)

            placed = False
            for tbl in tables_for_tag:
                if len(tbl) + size <= local_cap:
                    tbl.extend(idxs)
                    placed = True
                    break

            if not placed:
                tables_for_tag.append(idxs)

        # --- THEN SINGLES ---
        for idx in singles_df.index:
            placed = False
            for tbl in tables_for_tag:
                if len(tbl) < local_cap:
                    tbl.append(idx)
                    placed = True
                    break

            if not placed:
                tables_for_tag.append([idx])

        # Assign table numbers
        for tbl in tables_for_tag:
            tbl_num = next_table_number
            for idx in tbl:
                table_number_by_index[idx] = tbl_num
            next_table_number += 1

    # -----------------------------
    # FINALIZE TABLE NUMBERS
    # -----------------------------
    attending["table"] = attending.index.map(table_number_by_index)
    attending = attending[attending["table"].notna()].copy()
    attending["table"] = attending["table"].astype(int)

    # -----------------------------
    # PREPARE EXCEL EXPORT COLUMNS
    # -----------------------------
    meal_col = "meal"
    baby_col = "baby chair"
    carpark_col = "do you need a car park coupon? 您需要停车券吗？"
    other_req_col = (
        "if you have any other comments or requests not mentioned above, "
        "feel free to leave them here. 如果您有其他未提及的备注或需求，也欢迎在此填写."
    )
    comments_col = "comments"

    def combine_remarks(row):
        parts = []
        if pd.notna(row.get(other_req_col, "")) and row.get(other_req_col, "").strip():
            parts.append(row[other_req_col])
        if pd.notna(row.get(comments_col, "")) and row.get(comments_col, "").strip():
            parts.append(row[comments_col])
        return " | ".join(parts)

    attending["remarks"] = attending.apply(combine_remarks, axis=1)

    # Clean NO Values
    for col in [meal_col, baby_col, carpark_col]:
        attending[col] = attending[col].apply(clean_no)

    # -----------------------------
    # BUILD VERTICAL SEATING PLAN
    # -----------------------------
    rows = []
    max_rows = table_size

    columns_main = [
        "Table", "Name", "Meal preference", "Baby chair",
        "Car park coupon", "Remarks", "Tags"
    ]

    for tid in sorted(attending["table"].unique()):
        tid = int(tid)

        # Header
        rows.append(pd.DataFrame([[f"Table #{tid}"] + [""] * 6], columns=columns_main))

        # Subheader
        rows.append(pd.DataFrame([[""] + columns_main[1:]], columns=columns_main))

        # Guest rows
        tdf = (
            attending[attending["table"] == tid]
            .sort_values(["tag_group", "party", "full_name"])
            .reset_index(drop=True)[
                ["full_name", meal_col, baby_col, carpark_col, "remarks", "tags"]
            ]
        )

        tdf = tdf.rename(columns={
            "full_name": "Name",
            meal_col: "Meal preference",
            baby_col: "Baby chair",
            carpark_col: "Car park coupon",
            "remarks": "Remarks",
            "tags": "Tags"
        })

        # Pad
        if len(tdf) < max_rows:
            pad_rows = max_rows - len(tdf)
            pad_df = pd.DataFrame([[""] * 6] * pad_rows, columns=columns_main[1:])
            tdf = pd.concat([tdf, pad_df], ignore_index=True)

        # Insert row numbers
        tdf.insert(0, "Table", list(range(1, len(tdf) + 1)))
        rows.append(tdf)

        # Separator
        rows.append(pd.DataFrame([[""] * 7], columns=columns_main))

    seating_plan = pd.concat(rows, ignore_index=True)
    seating_plan["Table"] = seating_plan["Table"].astype(str)

    # -----------------------------
    # BUILD EXCEL OUTPUT
    # -----------------------------
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        seating_plan.to_excel(writer, sheet_name="SeatingPlan", index=False)
        pending.to_excel(writer, sheet_name="Pending_RSVP", index=False)
        declined.to_excel(writer, sheet_name="Declined", index=False)
        pending_tags.to_excel(writer, sheet_name="Pending_Tags", index=False)

    return buffer.getvalue(), attending, seating_plan


# -----------------------------
# STREAMLIT UI
# -----------------------------
st.title("💒 Wedding Seating Plan Generator")

if "df" not in st.session_state:
    st.session_state.df = None

uploaded = st.file_uploader("Upload your guest-list CSV", type=["csv"])
if uploaded:
    st.session_state.df = pd.read_csv(uploaded)
    st.success("CSV loaded successfully!")

if st.button("Use Sample Data"):
    try:
        st.session_state.df = pd.read_csv("guest-list.csv")
        st.success("Sample CSV loaded!")
    except Exception as e:
        st.error(f"Sample load failed: {e}")

if st.session_state.df is None:
    st.info("Upload a CSV or click sample to continue.")
    st.stop()

df = st.session_state.df

if st.button("Generate Seating Plan"):

    excel_bytes, attending_df, seating_plan_df = generate_seating_plan(df)

    st.metric("🎉 Joyfully Accepted", len(attending_df))

    st.subheader("📋 Table Summary")
    summary = (
        attending_df.groupby("table")
        .agg(
            guests=("full_name", "count"),
            tag_group=("tag_group", lambda x: x.mode().iloc[0])
        )
        .reset_index()
    )
    st.dataframe(summary, width="stretch")

    st.subheader("🪑 Full Seating Plan")
    st.dataframe(seating_plan_df, width="stretch", height=600)

    filename = f"Wedding_SeatingPlan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "📥 Download Seating Plan Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("Seating Plan generated and ready for download!")