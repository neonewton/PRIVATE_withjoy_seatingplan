import streamlit as st
import pandas as pd
import io
from datetime import datetime

SAMPLE_CSV_URL = "https://raw.githubusercontent.com/neonewton/PRIVATE_withjoy_seatingplan/main/guest-list.csv"


# -----------------------------
# Helper: clean "No"
# -----------------------------
def clean_no(value):
    """Blank out any content containing 'no' (case-insensitive)."""
    if pd.isna(value):
        return value
    s = str(value).lower()
    if "no" in s:
        return ""
    return value


# -----------------------------
# Core Seating Plan Generator
# -----------------------------
def generate_seating_plan(df, table_size=10):

    # --- 1. Clean column names and normalize ---
    df.columns = df.columns.str.strip()

    # --- 2. Combine name ---
    df["full_name"] = (
        df["first name"].fillna("") + " " + df["last name"].fillna("")
    ).str.strip()

    # --- 3. RSVP Processing ---
    rsvp_str = df["rsvp"].astype(str)

    declined_mask = rsvp_str.str.contains("Regretfully Decline", case=False, na=False)
    blank_mask = rsvp_str.str.strip().eq("") | df["rsvp"].isna()

    attending_mask = ~(declined_mask | blank_mask)

    attending = df[attending_mask].copy()
    pending = df[blank_mask].copy()
    declined = df[declined_mask].copy()

    # --- 4. Tag-group cleanup ---
    # Normalize tags for grouping
    attending["tag_group_raw"] = (
        attending["tags"]
        .fillna("")
        .astype(str)
        .str.replace(r"\s+", "", regex=True)  # remove all whitespace
        .str.lower()
        .str.strip()
    )

    # Guests with no usable tags → Pending_Tags
    no_tag_mask = attending["tag_group_raw"].isin(["", "nan", "none"])
    pending_tags = attending[no_tag_mask].copy()
    attending = attending[~no_tag_mask].copy()

    # Standardize tag_group names
    attending["tag_group"] = attending["tag_group_raw"]

    # --- 5. Table assignment by tag-groups ---
    table_number_by_index = {}
    next_table_number = 1

    tag_groups = list(attending["tag_group"].unique())

    # Place 'uncategorised' or blanks last if any
    if "uncategorised" in tag_groups:
        tag_groups = [tg for tg in tag_groups if tg != "uncategorised"] + ["uncategorised"]

    for tg in tag_groups:
        group_df = attending[attending["tag_group"] == tg]
        group_indices = list(group_df.index)
        group_size = len(group_indices)

        if group_size == 0:
            continue

        # Allow 11 pax table only when group size is exactly 11
        local_cap = table_size
        if group_size == table_size + 1:
            local_cap = table_size + 1

        # Party assignment
        party_str = group_df["party"].astype(str).str.strip()
        has_party = party_str != ""
        parties_df = group_df[has_party]
        singles_df = group_df[~has_party]

        tables_for_tag = []

        # --- Assign parties first ---
        for party_id, pgroup in parties_df.groupby("party"):
            indices = list(pgroup.index)
            size = len(indices)

            placed = False
            for tbl in tables_for_tag:
                if len(tbl) + size <= local_cap:
                    tbl.extend(indices)
                    placed = True
                    break

            if not placed:
                tables_for_tag.append(indices)

        # --- Assign singles next ---
        for idx in singles_df.index:
            placed = False
            for tbl in tables_for_tag:
                if len(tbl) < local_cap:
                    tbl.append(idx)
                    placed = True
                    break
            if not placed:
                tables_for_tag.append([idx])

        # Assign global table numbers
        for tbl in tables_for_tag:
            tbl_num = next_table_number
            for idx in tbl:
                table_number_by_index[idx] = tbl_num
            next_table_number += 1

    attending["table"] = attending.index.map(table_number_by_index)

    # --- 6. Prepare Excel output fields ---
    meal_col = "meal"
    baby_col = "baby chair"
    carpark_col = "do you need a car park coupon? 您需要停车券吗？"
    other_req_col = (
        "if you have any other comments or requests not mentioned above, "
        "feel free to leave them here. 如果您有其他未提及的备注或需求，也欢迎在此填写."
    )
    comments_col = "comments"

    # Combine R + S into "Remarks"
    def combine_remarks(row):
        parts = []
        r_val = row.get(other_req_col)
        s_val = row.get(comments_col)
        if pd.notna(r_val) and str(r_val).strip():
            parts.append(str(r_val).strip())
        if pd.notna(s_val) and str(s_val).strip():
            parts.append(str(s_val).strip())
        return " | ".join(parts) if parts else ""

    attending["remarks"] = attending.apply(combine_remarks, axis=1)

    # Clean "No"
    for col in [meal_col, baby_col, carpark_col]:
        attending[col] = attending[col].apply(clean_no)

    # --- 7. Build vertical seating plan ---
    table_ids = sorted(attending["table"].unique())
    rows = []
    max_rows = table_size

    columns_main = ["Table", "Name", "Meal preference", "Baby chair",
                    "Car park coupon", "Remarks", "Tags"]

    for tid in table_ids:

        # Header
        header_df = pd.DataFrame([[f"Table #{int(tid)}"] + [""] * 6], columns=columns_main)
        rows.append(header_df)

        # Subheader
        sub_df = pd.DataFrame([[""] + columns_main[1:]], columns=columns_main)
        rows.append(sub_df)

        # Guests in this table
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
            "tags": "Tags",
        })

        # Pad to exactly 10 rows
        if len(tdf) < max_rows:
            pad_df = pd.DataFrame([[""] * 6] * (max_rows - len(tdf)), columns=columns_main[1:])
            tdf = pd.concat([tdf, pad_df], ignore_index=True)
        else:
            tdf = tdf.iloc[:max_rows]

        # Insert row numbers
        tdf.insert(0, "Table", list(range(1, max_rows + 1)))
        rows.append(tdf)

        # Blank row separator
        rows.append(pd.DataFrame([[""] * 7], columns=columns_main))

    seating_plan = pd.concat(rows, ignore_index=True)
    seating_plan["Table"] = seating_plan["Table"].astype(str)

    # --- Build Excel ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        seating_plan.to_excel(writer, sheet_name="SeatingPlan", index=False)
        pending.to_excel(writer, sheet_name="Pending_RSVP", index=False)
        declined.to_excel(writer, sheet_name="Declined", index=False)
        pending_tags.to_excel(writer, sheet_name="Pending_Tags", index=False)

    return buffer.getvalue(), attending, seating_plan


# -----------------------------
# Streamlit Frontend
# -----------------------------
st.title("💒 Wedding Seating Plan Generator")

# Session state
if "df" not in st.session_state:
    st.session_state.df = None

# Upload CSV
uploaded = st.file_uploader("Upload your guest-list CSV", type=["csv"])
if uploaded:
    st.session_state.df = pd.read_csv(uploaded)
    st.success("CSV loaded successfully!")

# Sample CSV
if st.button("Use Sample Data"):
    try:
        st.session_state.df = pd.read_csv("guest-list.csv")
        st.success("Sample CSV loaded successfully!")
    except Exception as e:
        st.error(f"Failed to load sample CSV: {e}")

# Stop if no data
if st.session_state.df is None:
    st.info("Upload a CSV or click 'Use Sample Data' to begin.")
    st.stop()

df = st.session_state.df

# Generate button
if st.button("Generate Seating Plan"):
    excel_bytes, attending_df, seating_plan_df = generate_seating_plan(df)

    # Joyfully accepted
    st.metric("🎉 Joyfully Accepted", len(attending_df))

    # Table summary
    st.subheader("📋 Table Summary")
    summary = (
        attending_df.groupby("table")
        .agg(
            guests=("full_name", "count"),
            tag_group=("tag_group", lambda x: x.mode().iloc[0] if not x.mode().empty else "")
        )
        .reset_index()
        .rename(columns={"table": "Table Number"})
    )
    st.dataframe(summary, width="stretch")

    # Seating preview
    st.subheader("🪑 Full Seating Plan")
    st.dataframe(seating_plan_df, width="stretch", height=600)

    # Download
    filename = f"Wedding_SeatingPlan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    st.download_button(
        "📥 Download Seating Plan Excel",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
