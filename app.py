import streamlit as st
import pandas as pd
import io
from datetime import datetime

# -----------------------------
# Helper: clean "No"
# -----------------------------
def clean_no(value):
    """
    Any cell content that contains 'no' (case-insensitive)
    will be blanked out in the Excel output.
    Applied ONLY to: meal, baby chair, car park.
    """
    if pd.isna(value):
        return value
    s = str(value)
    if "no" in s.lower():
        return ""
    return value


# -----------------------------
# Core Seating Plan Generator
# -----------------------------
def generate_seating_plan(df, table_size=10):
    """
    df: raw guest-list DataFrame
    table_size: normal max size per table (10),
                but tag-groups with exactly 11 guests
                are allowed to form a special 11-pax table.
    Returns: (excel_bytes, attending_df)
    """

    # 1–2. Combine first + last name
    df["full_name"] = (
        df["first name"].fillna("") + " " + df["last name"].fillna("")
    ).str.strip()

    # 3. Process RSVP (Column N)
    rsvp_str = df["rsvp"].astype(str)

    declined_mask = rsvp_str.str.contains(
        "Regretfully Decline", case=False, na=False
    )
    blank_mask = rsvp_str.str.strip().eq("") | df["rsvp"].isna()
    attending_mask = ~(declined_mask | blank_mask)

    attending = df[attending_mask].copy()
    pending = df[blank_mask].copy()
    declined = df[declined_mask].copy()

    # 4. Determine tag-group from Column C (full tags string)
    #    Missing/empty tags -> 'Uncategorised'
    attending["tag_group"] = (
        attending["tags"]
        .fillna("Uncategorised")
        .astype(str)
        .str.strip()
    )
    attending.loc[attending["tag_group"] == "", "tag_group"] = "Uncategorised"

    # 5. Table assignment per tag_group
    table_number_by_index = {}
    next_table_number = 1

    # Order tag groups; keep 'Uncategorised' last if present
    tag_groups = list(attending["tag_group"].unique())
    if "Uncategorised" in tag_groups:
        tag_groups = [tg for tg in tag_groups if tg != "Uncategorised"] + ["Uncategorised"]

    for tg in tag_groups:
        group_df = attending[attending["tag_group"] == tg].copy()
        group_indices = group_df.index.tolist()
        group_size = len(group_indices)

        if group_size == 0:
            continue

        # Local capacity:
        # - Normally 10
        # - If this tag-group has exactly 11 guests, allow a special 11-pax table.
        local_cap = table_size
        if group_size == table_size + 1:  # e.g. 11 when table_size = 10
            local_cap = table_size + 1

        # Split into parties (non-empty 'party') and singles (empty 'party')
        party_str = group_df["party"].astype(str).str.strip()
        has_party = party_str != ""
        parties_df = group_df[has_party]
        singles_df = group_df[~has_party]

        # Tables for this tag_group: each is a list of row indices
        tables_for_tag = []

        # 5a. Place parties first, each party must stay together
        for party_id, p_group in parties_df.groupby("party"):
            indices = list(p_group.index)
            size = len(indices)

            placed = False
            # Try to place into an existing table of same tag-group
            for tbl in tables_for_tag:
                if len(tbl) + size <= local_cap:
                    tbl.extend(indices)
                    placed = True
                    break

            if not placed:
                # Create a new table for this party
                tables_for_tag.append(indices)

        # 5b. Place singles (no party) into any table with space, else new table
        for idx in singles_df.index:
            placed = False
            for tbl in tables_for_tag:
                if len(tbl) < local_cap:
                    tbl.append(idx)
                    placed = True
                    break
            if not placed:
                tables_for_tag.append([idx])

        # Assign global table numbers for these tag_group tables
        for tbl in tables_for_tag:
            tbl_num = next_table_number
            for idx in tbl:
                table_number_by_index[idx] = tbl_num
            next_table_number += 1

    # 6. Attach table numbers back to attending DataFrame
    attending["table"] = attending.index.map(table_number_by_index)

    # 7. Prepare columns for Excel output
    meal_col = "meal"  # Column O
    baby_col = "baby chair"  # Column P
    carpark_col = "do you need a car park coupon? 您需要停车券吗？"  # Column Q
    other_req_col = (
        "if you have any other comments or requests not mentioned above, "
        "feel free to leave them here. 如果您有其他未提及的备注或需求，也欢迎在此填写."
    )  # Column R
    comments_col = "comments"  # Column S

    # Combine R and S into one "Remarks" column
    def combine_remarks(row):
        parts = []
        r_val = row.get(other_req_col)
        s_val = row.get(comments_col)
        if pd.notna(r_val) and str(r_val).strip() != "":
            parts.append(str(r_val).strip())
        if pd.notna(s_val) and str(s_val).strip() != "":
            parts.append(str(s_val).strip())
        return " | ".join(parts) if parts else ""

    attending["remarks"] = attending.apply(combine_remarks, axis=1)

    # 8. Clean "No" from meal, baby chair, car park
    for col in [meal_col, baby_col, carpark_col]:
        attending[col] = attending[col].apply(clean_no)

    # --- Build VERTICAL SeatingPlan sheet ---
    table_ids = sorted(attending["table"].dropna().unique())
    rows = []
    max_rows = table_size  # 10 visible row numbers per table

    # All rows in SeatingPlan will share these exact columns
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
        # 1) Header row: "Table #X" in first col only
        header_df = pd.DataFrame(
            [[f"Table #{int(tid)}"] + [""] * (len(columns_main) - 1)],
            columns=columns_main,
        )
        rows.append(header_df)

        # 2) Subheader row: column labels
        subheader_df = pd.DataFrame(
            [[""] + columns_main[1:]],  # first cell blank, then labels
            columns=columns_main,
        )
        rows.append(subheader_df)

        # 3) Data rows for this table
        tdf = (
            attending[attending["table"] == tid]
            .sort_values(["tag_group", "party", "full_name"], na_position="first")
            .reset_index(drop=True)[
                ["full_name", meal_col, baby_col, carpark_col, "remarks", "tags"]
            ]
        )

        # Rename to match columns_main[1:]
        data_df = tdf.rename(
            columns={
                "full_name": "Name",
                meal_col: "Meal preference",
                baby_col: "Baby chair",
                carpark_col: "Car park coupon",
                "remarks": "Remarks",
                "tags": "Tags",
            }
        )

        # Pad to exactly max_rows rows (e.g. 10 guests per table)
        current_len = len(data_df)
        if current_len < max_rows:
            pad_rows = max_rows - current_len
            pad_df = pd.DataFrame(
                [[""] * (len(columns_main) - 1)] * pad_rows,
                columns=columns_main[1:],  # exclude "Table"
            )
            data_df = pd.concat([data_df, pad_df], ignore_index=True)
        else:
            data_df = data_df.iloc[:max_rows]

        # Insert row numbers 1–10 in Column A ("Table" col) for guest rows
        data_df.insert(0, "Table", list(range(1, max_rows + 1)))

        rows.append(data_df)

        # 4) Blank separator row between tables
        sep_df = pd.DataFrame(
            [[""] * len(columns_main)], columns=columns_main
        )
        rows.append(sep_df)

    # Stack all tables vertically
    seating_plan = pd.concat(rows, ignore_index=True)

    # -------- Build Excel in memory --------
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        seating_plan.to_excel(writer, sheet_name="SeatingPlan", index=False)
        pending.to_excel(writer, sheet_name="Pending_RSVP", index=False)
        declined.to_excel(writer, sheet_name="Declined", index=False)

    return buffer.getvalue(), attending, seating_plan


# -----------------------------
# Streamlit Frontend
# -----------------------------
st.title("💒 Wedding Seating Plan Generator")

uploaded = st.file_uploader("Upload your guest-list CSV file", type=["csv"])

if uploaded:
    df = pd.read_csv(uploaded)
    st.success("CSV loaded successfully!")

    if st.button("Generate Seating Plan"):
        excel_bytes, attending_df, seating_plan_df = generate_seating_plan(df)

        # --- TABLE SUMMARY ---
        st.subheader("📋 Table Summary")

        summary = (
            attending_df.groupby("table")
            .agg(
                guests=("full_name", "count"),
                tag_group=("tag_group", lambda x: x.mode().iloc[0] if not x.mode().empty else "")
            )
            .reset_index()
            .rename(columns={"table": "Table Number"})
            .sort_values("Table Number")
        )

        st.dataframe(summary, use_container_width=True)

        # --- FULL SEATING PLAN PREVIEW ---
        st.subheader("🪑 Full Seating Plan (Same as Excel)")

        st.dataframe(
            seating_plan_df,
            use_container_width=True,
            height=600
        )

        # --- DOWNLOAD BUTTON ---
        filename = f"Wedding_SeatingPlan_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        st.download_button(
            label="📥 Download Seating Plan Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload your guest-list CSV to begin.")
