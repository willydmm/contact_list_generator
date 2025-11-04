import streamlit as st
import pandas as pd
import re
from datetime import datetime

st.set_page_config(
    page_title="Contact List Converter", page_icon="📇", layout="centered"
)

st.title("📇 Contact List Converter")
st.caption(
    "Upload an Excel file, choose a sheet, and convert it to a clean CSV (phone, full_name, nick_name)."
)

# --- Controls
uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])
sheet_name = st.text_input(
    "Sheet name", value="Sheet2", help="Exact name of the sheet to read"
)


# --- Helpers
# Validate Phone Number
def validate_phone_number(phone_number: str) -> str:
    if pd.isna(phone_number):
        return ""
    phone_number = str(phone_number).strip()
    phone_number = re.sub(r"[^\d+]", "", phone_number)
    phone_number = re.sub(r"^\+62", "62", phone_number)
    phone_number = re.sub(r"^0", "62", phone_number)
    return phone_number


# Validate Ticket Number
def validate_ticket_number(ticket_number: str) -> str:
    if pd.isna(ticket_number):
        return ""
    ticket_number = str(ticket_number).strip()
    ticket_number = re.sub(r"[\\D]", "", ticket_number)
    return ticket_number


# Convert DataFrame to CSV bytes
def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


# --- Main Logic
if uploaded and sheet_name:
    try:
        df = pd.read_excel(uploaded, sheet_name=sheet_name, dtype=str)

        # Transform
        transformed_df = pd.DataFrame()
        transformed_df["phone"] = df["Phone"].map(validate_phone_number)

        ticket_clean = df["No Tiket"].map(validate_ticket_number)
        customer_clean = df["Customer Name"].astype(str).fillna("").str.strip()
        customer_clean = customer_clean.str.replace(",", " ", regex=False)

        transformed_df["nick_name"] = customer_clean

        korkel_clean = df["Korkel"].astype(str).fillna("").str.strip()
        korkel_clean = korkel_clean.str.replace(",", " ", regex=False)

        transformed_df["full_name"] = (
            "IR1 - " + korkel_clean + " - " + ticket_clean + " - " + customer_clean
        )

        valid_rows = (transformed_df["phone"] != "") & (
            transformed_df["full_name"] != ""
        )
        valid_df = transformed_df.loc[valid_rows, ["phone", "full_name", "nick_name"]]
        invalid_df = transformed_df.loc[
            ~valid_rows, ["phone", "full_name", "nick_name"]
        ]

        total = len(transformed_df)
        valid_count = len(valid_df)
        invalid_count = len(invalid_df)

        # --- Display Results
        c1, c2, c3 = st.columns(3)
        c1.metric("Total rows", total)
        c2.metric("Valid", valid_count)
        c3.metric("Invalid", invalid_count)
        # -- Preview and Download
        st.divider()
        st.subheader("Preview")
        st.write("**Valid rows**")
        st.dataframe(valid_df.head(50), use_container_width=True)
        if invalid_count > 0:
            st.write("**Invalid rows**")
            st.dataframe(invalid_df.head(50), use_container_width=True)

        st.divider()
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        valid_name = f"Contact_List_{timestamp_str}.csv"
        invalid_name = f"invalid_output_{timestamp_str}.csv"

        st.download_button(
            "⬇️ Download valid CSV",
            data=to_csv_bytes(valid_df),
            file_name=valid_name,
            mime="text/csv",
            use_container_width=True,
        )
        st.download_button(
            "⬇️ Download invalid CSV",
            data=to_csv_bytes(invalid_df),
            file_name=invalid_name,
            mime="text/csv",
            use_container_width=True,
        )

    except KeyError as e:
        st.error(
            f"Missing expected column in the sheet: {e}. Required columns: 'Phone', 'No Tiket', 'Customer Name'."
        )
    except ValueError as e:
        st.error(f"Sheet error: {e}")
    except Exception as e:
        st.exception(e)
else:
    st.info("Upload an .xlsx file and set the sheet name (default: 'Sheet2').")
