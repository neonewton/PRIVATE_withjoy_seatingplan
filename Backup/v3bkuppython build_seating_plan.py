import streamlit as st
import pandas as pd
import io
from datetime import datetime

# ---------------------------------------------------------
# Helper: Clean "No"
# ---------------------------------------------------------
def clean_no(value):
    if pd.isna(value):
        return value
    if "no" in str(value).lower():
        return ""
    return value


# ---------------------------------------------------------
# Core seating plan generator (modified to work in Streamlit)
# ---------------------------------------------------------
def generate_seating_plan(df, table_size=10):

    # Combine first + last name
    df["full_name"] = (
        df["first name"].fillna("") + " " + df["last name"].fillna("")
    ).str.strip()

    # RSVP handling
    rsvp_str = df["rsvp"].astype(str)
    declined_mask = rsvp_str.str.contains("Regretfully Decline", case=False, na=False)
    blank_mask = rsvp_str.str.strip().eq("") | df["rsvp"].isna()
    attending_mask = ~(declined_mask | blank_mask)

    attending = df[attending_mask].copy()
    pending = df[blank_mask].copy()
    declined = df[declined_mask].copy()

    # Tags → group name
    attending["tag_group"] = (
        attending["tags"]
        .fillna("Uncategorised")
        .astype(str)
        .str.strip()
    )
    attending.loc[attending["tag_group"] == "", "tag_group"] = "Uncategorised"

    # Table assignment logic
    table_number_by_index = {}
    next_table_number = 1

    tag_groups = list(attending["tag_group"].unique())
    if "Uncategorised" in tag_groups:
        tag_groups = [tg for tg in tag_groups if tg != "Uncategorised"] + ["Uncategorised"]

    for tg in tag_groups:
        group_df = attending[attending["tag_group"] == tg].copy()

        # party / non-party split
        party_str = group_df["party"].astype(str).str.strip()
        has_party = party_str != ""
        parties_df = group_df[has_party]
        singles_df = group_df[~has_party]

        tables_for_tag = []

        # Place full parties first
        for party_id, p_group in parties_df.groupby("party"):
            indices = list(p_group.index)
            size = len(indices)

            placed = False
            for tbl in tables_for_tag:
                if len(tbl) + size <= table_size:
                    tbl.extend(indices)
                    placed = True
                    break

            if not placed:
                tables_for_tag.append(indices)

        # Place singles into remaining tables or new ones
        for idx in singles_df.index:
            placed = False
            for tbl in tables_for_tag:
                if len(tbl) < table_size:
                    tbl.append(idx)
                    placed = True
                    break
            if not placed:
                tables_for_tag.append([idx])

        # Assign table numbers globally
        for tbl in tables_for_tag:
            tbl_num = next_table_number
            for idx in tbl:
                table_number_by_index[idx] = tbl_num
            next_table_number += 1

    attending["table"] = attending.index.map(table_number_by_index)

    # Columns
    meal_col = "meal"
    baby_col = "baby chair"
    carpark_col = "do you need a car park coupon? 您需要停车券吗？"
    other_req_col = (
        "if you have any other comments or requests not mentioned above, "
        "feel free to leave them here. 如果您有其他未提及的备注或需求，也欢迎在此填写."
    )
    comments_col = "comments"

    # Combine R+S → remarks
    def combine_remarks(row):
        parts = []
        if pd.notna(row.get(other_req_col)) and str(row.get(other_req_col)).strip():
            parts.append(str(row.get(other_req_col)).strip())
        if pd.notna(row.get(comments_col)) and str(row.get(comments_col)).strip():
            parts.append(str(row.get(comments_col)).strip())
        return " | ".join(parts)

    attending["remarks"] = attending.apply(combine_remarks, axis=1)

    # Clean "No" fields
    for col in [meal_col, baby_col, carpark_col]:
        attending[col] = attending[col].apply(clean_no)

    # Build vertical output
    table_ids = sorted(attending["table"].dropna().unique())
    rows = []
    max_rows = table_size

    columns_main = [
        "Table",
        "Name",
        "Meal preference",
        "Baby chair",
        "Car park coupon",
        "Remarks",
        "Tags",
    ]

    for tid in table_ids:
        # Header row
        header_df = pd.DataFrame(
            [[f"Table #{int(tid)}"] + [""] * 6],
            columns=columns_main,
        )
        rows.append(header_df)

        # Subheader
        subheader_df = pd.DataFrame(
            [[""] + columns_main[1:]],
            columns=columns_main,
        )
        rows.append(subheader_df)

        # Data rows
        tdf = (
            attending[attending["table"] == tid]
            .sort_values(["tag_group", "party", "full_name"])
            .reset_index(drop=True)[
                ["full_name", meal_col, baby_col, carpark_col, "remarks", "tags"]
            ]
        )

        tdf = tdf.rename(
            columns={
                "full_name": "Name",
                meal_col: "Meal preference",
                baby_col: "Baby chair",
                carpark_col: "Car park coupon",
                "remarks": "Remarks",
                "tags": "Tags",
            }
        )

        # pad to 10 rows
        current_len = len(tdf)
        if current_len < max_rows:
            pad_df = pd.DataFrame(
                [[""] * 6] * (max_rows - current_len),
                columns=["Name", "Meal preference", "Baby chair",
                         "Car park coupon", "Remarks", "Tags"],
            )
            tdf = pd.concat([tdf, pad_df], ignore_index=True)

        tdf.insert(0, "Table", list(range(1, max_rows + 1)))
        rows.append(tdf)

        # blank separator
        sep = pd.DataFrame([[""] * 7], columns=columns_main)
        rows.append(sep)

    seating_df = pd.concat(rows, ignore_index=True)

    # Create Excel file in memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        seating_df.to_excel(writer, sheet_name="SeatingPlan", index=False)
        pending.to_excel(writer, sheet_name="Pending_RSVP", index=False)
        declined.to_excel(writer, sheet_name="Declined", index=False)

    return buffer.getvalue(), attending

# ---------------------------------------------------------
# Streamlit Frontend
# ---------------------------------------------------------
st.title("💒 Wedding Seating Plan Generator")

uploaded = st.file_uploader("Upload your guest-list CSV file", type=["csv"])

if uploaded:
    df = pd.read_csv(uploaded)
    st.success("CSV loaded successfully!")

    if st.button("Generate Seating Plan"):
        excel_bytes, attending_df = generate_seating_plan(df)

        # ======== TABLE SUMMARY SECTION =========
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

        st.dataframe(summary, use_container_width=True)

        # ======== DOWNLOAD BUTTON =========
        filename = f"Wedding_SeatingPlan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        st.download_button(
            label="📥 Download Seating Plan Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

