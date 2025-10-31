import pandas as pd
import re
import sys
from datetime import datetime

# Configure input file path

if len(sys.argv) > 1:
    xlsx_path = sys.argv[1]
else:
    sys.exit("Error: please provide the input file path.")

sheet_name = "Sheet2"


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
    ticket_number = re.sub(r"[\D]", "", ticket_number)
    return ticket_number


# Read Excel file
df = pd.read_excel(xlsx_path, sheet_name=sheet_name, dtype=str)

# Transform columns
transformed_df = pd.DataFrame()
transformed_df["phone"] = df["Phone"].map(validate_phone_number)

# Create full_name = "IR1 - No Tiket - Customer Name"
ticket_clean = df["No Tiket"].map(validate_ticket_number)
customer_clean = df["Customer Name"].astype(str).fillna("").str.strip()

# Remove commas from Customer Name
customer_clean = customer_clean.str.replace(",", " ", regex=False)

# Keep nick_name as the customer name
transformed_df["nick_name"] = customer_clean

# Add prefix "IR1 - "
transformed_df["full_name"] = "IR1 - " + ticket_clean + " - " + customer_clean

# Set valid df
valid_rows = (transformed_df["phone"] != "") & (transformed_df["full_name"] != "")
valid_df = transformed_df.loc[valid_rows, ["phone", "full_name", "nick_name"]]

# Print count before saving
valid_count = len(valid_df)
print(f"Saving {valid_count} contact{'s' if valid_count != 1 else ''}...")

# Format today's date and time
timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")  # → 20251030_142530

# Build filename with timestamp
output_filename = f"Contact_List_{timestamp_str}.csv"

# Save to CSV
valid_df.to_csv(output_filename, index=False)

# Save invalid df
invalid_df = transformed_df.loc[~valid_rows, ["phone", "full_name", "nick_name"]]
invalid_df.to_csv("invalid_output.csv", index=False)
