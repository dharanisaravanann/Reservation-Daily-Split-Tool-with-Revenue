import os
import streamlit as st
import pandas as pd
from io import BytesIO

st.write("RUNNING FILE:", os.path.abspath(__file__))


def to_excel_serial(series: pd.Series) -> pd.Series:
    """
    Convert a column to Excel DATEVALUE serial numbers.
    - If values are already numeric (e.g., 45292), keep them.
    - Otherwise parse as datetime and convert to serial using Excel epoch.
    Returns pandas nullable integer (Int64).
    """
    EXCEL_EPOCH = pd.Timestamp("1899-12-30")
    s = series.copy()

    # 1) Try numeric first (already Excel serials)
    num = pd.to_numeric(s, errors="coerce")
    is_num = num.notna()

    # 2) Parse non-numeric values as datetime
    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")

    out = pd.Series(pd.NA, index=s.index, dtype="float")
    out[is_num] = num[is_num]
    out[~is_num] = (dt[~is_num] - EXCEL_EPOCH).dt.days

    return out.astype("Int64")


def reservation_with_revenue(df: pd.DataFrame) -> pd.DataFrame:
    """
    Keeps 1 row per reservation.
    Converts Arrival/Departure/Booking Date to Excel DATEVALUE serial numbers.
    Cleans revenue columns to numeric.
    """
    df = df.copy()
    df.columns = df.columns.str.strip()

    # Convert date columns to Excel serial numbers (robust)
    for col in ["Arrival", "Departure", "Booking Date"]:
        if col in df.columns:
            df[col] = to_excel_serial(df[col])

    # Revenue-related columns
    revenue_cols = [
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

    # Ensure numeric (keep as floats; fill blanks with 0)
    for col in revenue_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Rename Channel -> Sub Channel
    if "Channel" in df.columns and "Sub Channel" not in df.columns:
        df = df.rename(columns={"Channel": "Sub Channel"})

    desired_cols = [
        "Reservation Number",
        "Apartment",
        "Guest Name",
        "Sub Channel",
        "Arrival",        # Excel serial
        "Departure",      # Excel serial
        "Booking Date",   # Excel serial
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

    # Keep only columns that exist (prevents crashes)
    desired_cols = [c for c in desired_cols if c in df.columns]
    return df[desired_cols]


# ---------------- STREAMLIT APP ----------------

st.title("Reservation Revenue Summary Tool (DATEVALUE)")

st.write(
    "Upload a reservations Excel file (.xlsx) and this tool will:\n"
    "- Keep **one row per reservation**\n"
    "- Convert **Arrival / Departure / Booking Date** into **Excel DATEVALUE serial numbers** (e.g., 01/01/2024 → 45292)\n"
    "- Clean revenue/fee columns to numeric\n"
    "- Return an Excel file with two sheets: **Original Data** + **Reservation Revenue Summary**"
)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_input = pd.read_excel(uploaded_file)
        df_input.columns = df_input.columns.str.strip()

        st.subheader("Preview of uploaded data")
        st.dataframe(df_input.head(), use_container_width=True)

        df_output = reservation_with_revenue(df_input)

        st.subheader("Preview of reservation revenue summary (first 20 rows)")
        st.dataframe(df_output.head(20), use_container_width=True)

        # Build Excel in memory
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            # Keep Original Data as-is (raw upload)
            df_input.to_excel(writer, sheet_name="Original Data", index=False)

            # Output summary
            df_output.to_excel(writer, sheet_name="Reservation Revenue Summary", index=False)

        buffer.seek(0)

        st.download_button(
            label="📥 Download Excel (Original + Revenue Summary)",
            data=buffer,
            file_name="reservation_revenue_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")

else:
    st.info("Please upload an Excel file to begin.")
