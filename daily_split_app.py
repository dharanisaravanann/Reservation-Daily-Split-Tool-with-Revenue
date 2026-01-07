import os
import streamlit as st
import pandas as pd
from io import BytesIO

TEMPLATE_COLUMNS = [
    "Reservation Number",
    "Apartment",
    "Guest Name",
    "Channel",
    "Arrival",
    "Departure",
    "Booking Date",
    "Base Revenue",
    "Total Revenue",
    "Room Revenue",
    "SC on Room Revenue",
    "VAT on Room Rev",
    "VAT on SC",
    "Cleaning Fees Without VAT",
    "VAT on Cleaning Fees",
    "Tourism Dirham Fees",
    "Cleaning Fees",
]

def build_template_excel() -> BytesIO:
    template_df = pd.DataFrame(columns=TEMPLATE_COLUMNS)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        template_df.to_excel(writer, sheet_name="Template", index=False)
    buffer.seek(0)
    return buffer

def split_reservations_daily(df: pd.DataFrame) -> pd.DataFrame:
    """
    1 row per night.
    Splits all revenue/fee columns evenly across nights by dividing by total nights (per ORIGINAL row).
    Also keeps sums consistent after rounding by pushing rounding remainder to the last night.
    """
    df = df.copy()

    # Clean column names
    df.columns = (
        df.columns.astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    required = ["Reservation Number", "Arrival", "Departure", "Booking Date"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Convert date columns
    df["Arrival"] = pd.to_datetime(df["Arrival"], dayfirst=True, errors="coerce")
    df["Departure"] = pd.to_datetime(df["Departure"], dayfirst=True, errors="coerce")
    df["Booking Date"] = pd.to_datetime(df["Booking Date"], dayfirst=True, errors="coerce")

    # Unique id per input row (critical fix)
    df["_line_id"] = range(len(df))

    # Nights per original row
    df["_total_nights"] = (df["Departure"] - df["Arrival"]).dt.days
    df["_total_nights"] = df["_total_nights"].fillna(0).astype(int)

    # Build nightly ranges: Arrival -> day before Departure
    df["Stay_dates"] = [
        pd.date_range(start, end - pd.Timedelta(days=1), freq="D")
        if pd.notna(start) and pd.notna(end) and end > start
        else pd.DatetimeIndex([])
        for start, end in zip(df["Arrival"], df["Departure"])
    ]

    # Explode
    df_daily = df.explode("Stay_dates").reset_index(drop=True)
    df_daily = df_daily[df_daily["Stay_dates"].notna()].copy()

    # Excel epoch
    EXCEL_EPOCH = pd.Timestamp("1899-12-30")

    # Night Date as Excel serial
    stay_dt = pd.to_datetime(df_daily["Stay_dates"], errors="coerce")
    df_daily["Date"] = (stay_dt - EXCEL_EPOCH).dt.days

    # Booking Date as Excel serial
    booking_dt = pd.to_datetime(df_daily["Booking Date"], errors="coerce")
    df_daily["Booking Date"] = (booking_dt - EXCEL_EPOCH).dt.days

    # 1 row = 1 night
    df_daily["Nights"] = 1

    # Money columns (raw names BEFORE renaming)
    money_cols = [
        "Base Revenue",
        "Total Revenue",
        "Room Revenue",
        "SC on Room Revenue",
        "VAT on Room Rev",
        "VAT on SC",
        "Cleaning Fees Without VAT",
        "VAT on Cleaning Fees",
        "Tourism Dirham Fees",
        "Cleaning Fees",
    ]

    # Convert money to numeric
    for col in money_cols:
        if col in df_daily.columns:
            df_daily[col] = pd.to_numeric(df_daily[col], errors="coerce").fillna(0.0)

    # Divide by nights per ORIGINAL row (critical fix)
    # Avoid divide-by-zero: if _total_nights is 0 (bad dates), keep 0
    denom = df_daily["_total_nights"].replace(0, pd.NA)

    for col in money_cols:
        if col in df_daily.columns:
            df_daily[col] = (df_daily[col] / denom).fillna(0.0)

    # OPTIONAL (recommended): make sums match exactly to 2dp after rounding
    # We round per-night values to 2dp, then push rounding remainder to the last night of each _line_id
    def _fix_rounding(group: pd.DataFrame) -> pd.DataFrame:
        # identify last night row in this line
        last_idx = group.index.max()

        for col in money_cols:
            if col not in group.columns:
                continue

            original_total = group[col].sum()  # this is already "distributed" total, equals original money for the line
            rounded = group[col].round(2)
            diff = (original_total - rounded.sum()).round(2)

            group[col] = rounded
            if pd.notna(diff) and diff != 0:
                group.loc[last_idx, col] = (group.loc[last_idx, col] + diff).round(2)

        return group

    df_daily = df_daily.groupby("_line_id", group_keys=False).apply(_fix_rounding)

    # Drop helper column
    df_daily.drop(columns=["Stay_dates"], inplace=True)

    # Rename Channel -> Sub Channel
    if "Channel" in df_daily.columns:
        df_daily = df_daily.rename(columns={"Channel": "Sub Channel"})

    # Rename AFTER splitting
    rename_map = {
        "Base Revenue": "Base Revenue per Night",
        "Total Revenue": "Total Revenue per Night",
        "Room Revenue": "Room Revenue per Night",
        "SC on Room Revenue": "Service Charge per Night",
        "VAT on Room Rev": "VAT on Room Revenue per Night",
        "VAT on SC": "VAT on Service Charge per Night",
        "Cleaning Fees Without VAT": "Cleaning Fees (Excl VAT) per Night",
        "VAT on Cleaning Fees": "VAT on Cleaning Fees per Night",
        "Tourism Dirham Fees": "Tourism Dirham Fees per Night",
        "Cleaning Fees": "Cleaning Fees per Night",
    }
    df_daily = df_daily.rename(columns=rename_map)

    # Drop Arrival/Departure from nightly output
    for col in ["Arrival", "Departure"]:
        if col in df_daily.columns:
            df_daily.drop(columns=[col], inplace=True)

    # Output order
    desired_cols = [
        "Reservation Number",
        "Apartment",
        "Guest Name",
        "Sub Channel",
        "Date",
        "Booking Date",
        "Nights",
        "Base Revenue per Night",
        "Total Revenue per Night",
        "Room Revenue per Night",
        "Service Charge per Night",
        "VAT on Room Revenue per Night",
        "VAT on Service Charge per Night",
        "Cleaning Fees (Excl VAT) per Night",
        "VAT on Cleaning Fees per Night",
        "Tourism Dirham Fees per Night",
        "Cleaning Fees per Night",
    ]
    desired_cols = [c for c in desired_cols if c in df_daily.columns]

    # Remove internal helper columns
    for c in ["_line_id", "_total_nights"]:
        if c in df_daily.columns:
            df_daily.drop(columns=[c], inplace=True)

    return df_daily[desired_cols]


# ---------------- STREAMLIT APP ----------------

st.title("Reservation Daily Split Tool with Revenue (DATEVALUE)")
st.write(
    "Upload a reservations Excel file (.xlsx) and this tool will:\n"
    "- Create **1 row per night**\n"
    "- Convert **Date** (night date) and **Booking Date** to **Excel DATEVALUE serial numbers**\n"
    "- Divide **all revenue/fee columns** evenly across nights\n"
    "- Return an Excel file with two sheets: **Original Data** + **Reservations Daily Split**\n"
    "- Provide a Free **Excel Template** to paste data in correct format"
)

st.subheader("Step 0 — Download the Free Excel Template")
st.write(
    "If your column names are different, download this template, "
    "paste your data under the headers, then upload it."
)

template_buffer = build_template_excel()
st.download_button(
    label="📥 Download Excel Template",
    data=template_buffer,
    file_name="reservations_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.subheader("Step 1 — Upload your Excel file (.xlsx)")
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_input = pd.read_excel(uploaded_file)
        df_input.columns = (
            df_input.columns.astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

        st.subheader("Preview of uploaded data")
        st.dataframe(df_input.head(), use_container_width=True)

        df_output = split_reservations_daily(df_input)

        st.subheader("Preview of daily split (first 20 rows)")
        st.dataframe(df_output.head(20), use_container_width=True)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_input.to_excel(writer, sheet_name="Original Data", index=False)
            df_output.to_excel(writer, sheet_name="Reservations Daily Split", index=False)

        buffer.seek(0)

        st.download_button(
            label="📥 Download Excel (Original + Daily Split)",
            data=buffer,
            file_name="reservations_with_daily_split_DATEVALUE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
else:
    st.info("Please upload an Excel file to begin.")

