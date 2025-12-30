import os
import streamlit as st
import pandas as pd
from io import BytesIO


# --- Debug: show exactly which file Streamlit is running ---
st.write("RUNNING FILE:", os.path.abspath(__file__))


def split_reservations(df: pd.DataFrame) -> pd.DataFrame:
    """
    1 row per night.
    Splits revenue/fee columns evenly across nights by dividing by total nights.
    Date + Booking Date are converted to Excel serial date values (DATEVALUE style).
    """
    df = df.copy()
    df.columns = df.columns.str.strip()

    # Required columns check
    required = ["Reservation Number", "Arrival", "Departure", "Booking Date"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Convert date columns to datetime
    df["Arrival"] = pd.to_datetime(df["Arrival"], dayfirst=True, errors="coerce")
    df["Departure"] = pd.to_datetime(df["Departure"], dayfirst=True, errors="coerce")
    df["Booking Date"] = pd.to_datetime(df["Booking Date"], dayfirst=True, errors="coerce")

    # Create stay dates list (Arrival to day before Departure)
    df["Stay_dates"] = [
        pd.date_range(start, end - pd.Timedelta(days=1), freq="D")
        if pd.notna(start) and pd.notna(end) and end > start
        else pd.DatetimeIndex([])
        for start, end in zip(df["Arrival"], df["Departure"])
    ]

    # Explode to one row per night
    df_daily = df.explode("Stay_dates").reset_index(drop=True)
    df_daily = df_daily[df_daily["Stay_dates"].notna()].copy()

    # Excel DATEVALUE epoch
    EXCEL_EPOCH = pd.Timestamp("1899-12-30")

    # Create Date (stay date) as Excel serial number
    df_daily["DateValue"] = (
        pd.to_datetime(df_daily["Stay_dates"], errors="coerce") - EXCEL_EPOCH
    ).dt.days

    # Rename DateValue -> Date (as you want)
    df_daily = df_daily.rename(columns={"DateValue": "Date"})

    # Booking Date as Excel serial number too
    df_daily["Booking Date"] = (
        pd.to_datetime(df_daily["Booking Date"], errors="coerce") - EXCEL_EPOCH
    ).dt.days

    # Drop helper column
    df_daily.drop(columns=["Stay_dates"], inplace=True)

    # Total nights per reservation (count of nightly rows)
    total_nights = df_daily.groupby("Reservation Number")["Date"].transform("size")

    # Each row is one night
    df_daily["Nights"] = 1

    # Columns to split per night
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

    # Split values per night
    for col in money_cols:
        if col in df_daily.columns:
            df_daily[col] = pd.to_numeric(df_daily[col], errors="coerce")
            df_daily[col] = (df_daily[col] / total_nights).round(2)

    # Rename columns AFTER splitting
    rename_map = {
        "Base Revenue": "Base Revenue per Night",
        "Total Revenue": "Total Revenue per Night",
        "Room Revenue": "Room Revenue per Night",
        "SC on Room Revenue": "SC on Room Revenue per Night",
        "VAT on Room Rev": "VAT on Room Rev per Night",
        "VAT on SC": "VAT on SC per Night",
        "Cleaning Fees Without VAT": "Cleaning Fees Without VAT per Night",
        "VAT on Cleaning Fees": "VAT on Cleaning Fees per Night",
        "Tourism Dirham Fees": "Tourism Dirham Fees per Night",
        "Cleaning Fees": "Cleaning Fees per Night",
    }
    df_daily = df_daily.rename(columns=rename_map)

    # Rename Channel -> Sub Channel if exists
    if "Channel" in df_daily.columns:
        df_daily = df_daily.rename(columns={"Channel": "Sub Channel"})

    # Drop Arrival/Departure (optional)
    for col in ["Arrival", "Departure"]:
        if col in df_daily.columns:
            df_daily.drop(columns=[col], inplace=True)

    # Keep desired columns in order
    desired_cols = [
        "Reservation Number",
        "Apartment",
        "Guest Name",
        "Sub Channel",
        "Date",          # Excel serial stay date
        "Booking Date",  # Excel serial booking date
        "Nights",
        "Base Revenue per Night",
        "Total Revenue per Night",
        "Room Revenue per Night",
        "SC on Room Revenue per Night",
        "VAT on Room Rev per Night",
        "VAT on SC per Night",
        "Cleaning Fees Without VAT per Night",
        "VAT on Cleaning Fees per Night",
        "Tourism Dirham Fees per Night",
        "Cleaning Fees per Night",
    ]
    desired_cols = [c for c in desired_cols if c in df_daily.columns]
    return df_daily[desired_cols]


# ---------------- STREAMLIT APP ----------------

st.title("Reservation Daily Split Tool")

st.write(
    "Upload a reservations Excel file (.xlsx) and this tool will:\n"
    "- Split each booking into daily rows (1 row per night)\n"
    "- Split all revenue/fee columns evenly across nights\n"
    "- Convert Date + Booking Date into Excel date values (DATEVALUE style)\n"
    "- Return an Excel file with two sheets: Original Data + Reservations Daily Split."
)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_input = pd.read_excel(uploaded_file)
        df_input.columns = df_input.columns.str.strip()

        st.subheader("Preview of uploaded data")
        st.dataframe(df_input.head(), use_container_width=True)

        df_output = split_reservations(df_input)

        st.subheader("Preview of daily split (first 20 rows)")
        st.dataframe(df_output.head(20), use_container_width=True)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_original = df_input.copy()
            df_original.columns = df_original.columns.str.strip()
            df_original.to_excel(writer, sheet_name="Original Data", index=False)
            df_output.to_excel(writer, sheet_name="Reservations Daily Split", index=False)

        buffer.seek(0)

        st.download_button(
            label="📥 Download Excel (Original + Daily Split)",
            data=buffer,
            file_name="reservations_with_daily_split.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
else:
    st.info("Please upload an Excel file to begin.")
