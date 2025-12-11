import pandas as pd
import time


INPUT_CSV = "guest-list.csv"          # your raw export
OUTPUT_XLSX = "Wedding_seating_plan.xlsx"     # output file
TABLE_SIZE = 10                       # max people per table


def extract_category_from_tags(tags: str) -> str:
    """
    Turn the 'tags' column into a simpler category.

    Examples from your sheet:
      'Groom_Nelton, ArmyFriends'           -> 'ArmyFriends'
      'Groom_Nelton, Family'                -> 'Family'
      'Groom_Nelton, LKCMedWorkColleagues'  -> 'LKCMedWorkColleagues'
      'Table 1_VIP, Groom_Nelton'           -> 'Groom_Nelton'
      'Table 2_VIP, Bride_Steph'            -> 'Bride_Steph'
      'Groom_Nelton'                        -> 'Groom_Nelton'
    """
    if pd.isna(tags):
        return "Uncategorised"

    parts = [p.strip() for p in str(tags).split(",") if p.strip()]
    if not parts:
        return "Uncategorised"

    # If it starts with "Table", use the last part (e.g. Groom_Nelton / Bride_Steph)
    if parts[0].startswith("Table ") and len(parts) > 1:
        return parts[-1]

    # In most of your rows, the last part is the "where they are from" piece
    return parts[-1]


def build_seating(input_path: str, output_path: str, table_size: int = 10):
    df = pd.read_csv(input_path)

    # A + B -> Combined name
    df["full_name"] = (
        df["first name"].fillna("") + " " + df["last name"].fillna("")
    ).str.strip()

    # N -> RSVP
    rsvp_str = df["rsvp"].astype(str)

    declined_mask = rsvp_str.str.contains(
        "Regretfully Decline", case=False, na=False
    )
    blank_mask = rsvp_str.str.strip().eq("") | df["rsvp"].isna()
    attending_mask = ~(declined_mask | blank_mask)

    attending = df[attending_mask].copy()
    pending = df[blank_mask].copy()
    declined = df[declined_mask].copy()

    # C -> category (from 'tags')
    attending["category"] = attending["tags"].apply(extract_category_from_tags)

    # M -> party; everyone in same party sits together
    attending["party_id"] = attending["party"]
    missing_party = (
        attending["party_id"].isna()
        | (attending["party_id"].astype(str).str.strip() == "")
    )
    attending.loc[missing_party, "party_id"] = (
        "solo-" + attending[missing_party].index.astype(str)
    )

    # Decide order of categories (largest first so they get nice blocks of tables)
    cat_order = (
        attending.groupby("category")["full_name"]
        .count()
        .sort_values(ascending=False)
        .index.tolist()
    )

    # Build party objects: each party is an indivisible block
    parties = []
    for (category, party_id), group in attending.groupby(["category", "party_id"]):
        parties.append(
            {
                "category": category,
                "party_id": party_id,
                "size": len(group),
                "indices": group.index.tolist(),
            }
        )

    # Seating assignment
    row_to_table = {}  # row index -> table number
    tables = []        # each: {id, category, remaining, indices}
    current_table = 0

    for cat in cat_order:
        # Parties of this category, largest first
        cat_parties = [p for p in parties if p["category"] == cat]
        cat_parties.sort(key=lambda p: p["size"], reverse=True)

        # existing tables already used by this category
        cat_tables = [t for t in tables if t.get("category") == cat]

        for p in cat_parties:
            placed = False

            # 1) Try to place into an existing table of same category
            for t in cat_tables:
                if t["remaining"] >= p["size"]:
                    t["remaining"] -= p["size"]
                    t["indices"].extend(p["indices"])
                    for idx in p["indices"]:
                        row_to_table[idx] = t["id"]
                    placed = True
                    break

            if placed:
                continue

            # 2) Otherwise, reuse any table that still has spare seats
            for t in tables:
                if t["remaining"] >= p["size"]:
                    t["remaining"] -= p["size"]
                    t["indices"].extend(p["indices"])
                    for idx in p["indices"]:
                        row_to_table[idx] = t["id"]
                    placed = True
                    break

            if placed:
                continue

            # 3) Otherwise, create a new table just for this party/category
            current_table += 1
            t = {
                "id": current_table,
                "category": cat,
                "remaining": table_size - p["size"],
                "indices": list(p["indices"]),
            }
            tables.append(t)
            cat_tables.append(t)
            for idx in p["indices"]:
                row_to_table[idx] = current_table

    # Attach table numbers back to attending DataFrame
    attending["table"] = attending.index.map(row_to_table)

    # Columns you want in the output
    meal_col = "meal"  # O
    baby_col = "baby chair"  # P
    carpark_col = "do you need a car park coupon? 您需要停车券吗？"  # Q
    other_req_col = (
        "if you have any other comments or requests not mentioned above, "
        "feel free to leave them here. 如果您有其他未提及的备注或需求，也欢迎在此填写."
    )  # R
    comments_col = "comments"  # S

    # Long form (one row per guest, easier for checking)
    long_cols = [
        "table",
        "category",
        "party_id",
        "full_name",
        meal_col,
        baby_col,
        carpark_col,
        other_req_col,
        comments_col,
    ]

    seating_long = attending[long_cols].sort_values(
        ["table", "category", "party_id", "full_name"]
    )

    # Wide / landscape form: each table as 6 columns, up to 10 rows
    table_ids = sorted(seating_long["table"].unique())
    blocks = []
    max_rows = table_size

    for tid in table_ids:
        tdf = (
            seating_long[seating_long["table"] == tid]
            .reset_index(drop=True)[
                [
                    "full_name",
                    meal_col,
                    baby_col,
                    carpark_col,
                    other_req_col,
                    comments_col,
                ]
            ]
        )

        # pad to exactly max_rows = 10 rows
        if len(tdf) < max_rows:
            pad = pd.DataFrame(index=range(max_rows - len(tdf)), columns=tdf.columns)
            tdf = pd.concat([tdf, pad], ignore_index=True)
        else:
            tdf = tdf.iloc[:max_rows]

        # MultiIndex columns: top level = table number, second = field
        tdf.columns = pd.MultiIndex.from_product(
            [
                [f"Table {tid}"],
                [
                    "Name",
                    "Meal",
                    "Baby chair",
                    "Car park coupon",
                    "Other requests",
                    "Internal comments",
                ],
            ]
        )
        blocks.append(tdf)

    seating_wide = pd.concat(blocks, axis=1)

    # Write everything to Excel
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        # Landscape seating plan
        seating_wide.to_excel(writer, sheet_name="SeatingPlan")
        # For checking / filtering / manual tweaks
        seating_long.to_excel(writer, sheet_name="Attending_Long", index=False)
        # Parked aside
        pending.to_excel(writer, sheet_name="Pending_RSVP", index=False)
        declined.to_excel(writer, sheet_name="Declined", index=False)

    print(f"Done. Wrote seating plan to: {output_path}")


if __name__ == "__main__":
    build_seating(INPUT_CSV, OUTPUT_XLSX, TABLE_SIZE)
