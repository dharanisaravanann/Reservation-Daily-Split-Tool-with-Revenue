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

st.write("RUNNING FILE:", os.path.abspath(__file__))


def split_reservations_daily(df: pd.DataFrame) -> pd.DataFrame:
    """
    1 row per night.
    Splits all revenue/fee columns evenly across nights by dividing by total nights.
    Outputs:
      - Date = Excel DATEVALUE serial number for each night
      - Booking Date = Excel DATEVALUE serial number
    """
    df = df.copy()
    df.columns = df.columns.str.strip()

    # Required columns
    required = ["Reservation Number", "Arrival", "Departure", "Booking Date"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")
    # Convert key date columns to datetime
    df["Arrival"] = pd.to_datetime(df["Arrival"], dayfirst=True, errors="coerce")
    df["Departure"] = pd.to_datetime(df["Departure"], dayfirst=True, errors="coerce")
    df["Booking Date"] = pd.to_datetime(df["Booking Date"], dayfirst=True, errors="coerce")

    # Build nightly date ranges: Arrival -> day before Departure
    df["Stay_dates"] = [
        pd.date_range(start, end - pd.Timedelta(days=1), freq="D")
        if pd.notna(start) and pd.notna(end) and end > start
        else pd.DatetimeIndex([])
        for start, end in zip(df["Arrival"], df["Departure"])
    ]

    # Explode: one row per night
    df_daily = df.explode("Stay_dates").reset_index(drop=True)
    df_daily = df_daily[df_daily["Stay_dates"].notna()].copy()

    # Excel epoch (DATEVALUE base)
    EXCEL_EPOCH = pd.Timestamp("1899-12-30")

    # Date (night date) as Excel serial number
    stay_dt = pd.to_datetime(df_daily["Stay_dates"], errors="coerce")
    df_daily["Date"] = (stay_dt - EXCEL_EPOCH).dt.days

    # Booking Date as Excel serial number too
    booking_dt = pd.to_datetime(df_daily["Booking Date"], errors="coerce")
    df_daily["Booking Date"] = (booking_dt - EXCEL_EPOCH).dt.days

    # Drop helper column
    df_daily.drop(columns=["Stay_dates"], inplace=True)

    # Total nights per reservation (count nightly rows)
    total_nights = df_daily.groupby("Reservation Number")["Date"].transform("size")

    # Each row = 1 night
    df_daily["Nights"] = 1

    # All revenue/fee columns to split per night
    money_cols = [
        "Base Revenue",
        "Total Revenue",
        "Room Revenue",
        "SC on Room Revenue",
        "VAT on Room Rev",
        "VAT on SC Per Night",
        "Cleaning Fees Without VAT",
        "VAT on Cleaning Fees",
        "Tourism Dirham Fees ",
        "Cleaning Fees",
    ]

    # Convert to numeric + split per night
    for col in money_cols:
        if col in df_daily.columns:
            df_daily[col] = pd.to_numeric(df_daily[col], errors="coerce")
            df_daily[col] = (df_daily[col] / total_nights).round(2)

    # Rename Channel -> Sub Channel (if needed)
    if "Channel" in df_daily.columns:
        df_daily = df_daily.rename(columns={"Channel": "Sub Channel"})

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

    # drop Arrival/Departure from the nightly output
    for col in ["Arrival", "Departure"]:
        if col in df_daily.columns:
            df_daily.drop(columns=[col], inplace=True)

    # Output column order
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
    "Cleaning Fees per Night"
    ] + [c for c in money_cols if c in df_daily.columns]

    # Keep only columns that exist
    desired_cols = [c for c in desired_cols if c in df_daily.columns]
    return df_daily[desired_cols]

def build_template_excel() -> BytesIO:
    template_df = pd.DataFrame(columns=TEMPLATE_COLUMNS)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        template_df.to_excel(writer, sheet_name="Template", index=False)

    buffer.seek(0)
    return buffer

# ---------------- STREAMLIT APP ----------------

st.title("Reservation Daily Split Tool with Revenue(DATEVALUE)")
st.write(
    "Upload a reservations Excel file (.xlsx) and this tool will:\n"
    "- Creates **1 row per night**\n"
    "- Converts **Date** (night date) and **Booking Date** to **Excel DATEVALUE serial numbers**\n"
    "- Divide **all revenue/fee columns** evenly across nights\n"
    "- Return an Excel file with two sheets: **Original Data** + **Reservations Daily Split**"
    "- **Free Excel Template to paste data in correct format**"
)


st.subheader("Step 0 — Download the upload template")

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

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_input = pd.read_excel(uploaded_file)
        df_input.columns = df_input.columns.str.strip()

        st.subheader("Preview of uploaded data")
        st.dataframe(df_input.head(), use_container_width=True)

        df_output = split_reservations_daily(df_input)

        st.subheader("Preview of daily split (first 20 rows)")
        st.dataframe(df_output.head(20), use_container_width=True)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            # Original sheet (raw)
            df_input.to_excel(writer, sheet_name="Original Data", index=False)

            # Daily split sheet
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
