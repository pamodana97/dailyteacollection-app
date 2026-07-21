import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile
import os

# -------------------------------------------------------
# Page Configuration
# -------------------------------------------------------

st.set_page_config(
    page_title="Tea Collection System",
    page_icon="🍃",
    layout="wide"
)

# -------------------------------------------------------
# Title
# -------------------------------------------------------

st.title("🍃 Tea Collection Entry System")

st.markdown("---")

# -------------------------------------------------------
# Upload Excel
# -------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload Monthly Tea Collection Excel File",
    type=["xlsx"]
)

# -------------------------------------------------------
# If no file uploaded
# -------------------------------------------------------

if uploaded_file is None:

    st.info("Please upload your Excel file to begin.")

    st.stop()

# -------------------------------------------------------
# Save uploaded file temporarily
# -------------------------------------------------------

temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

temp_file.write(uploaded_file.read())

temp_file.close()

# -------------------------------------------------------
# Load Workbook
# -------------------------------------------------------

workbook = load_workbook(temp_file.name)

sheet = workbook.active

# -------------------------------------------------------
# Display Success
# -------------------------------------------------------

st.success("Excel file loaded successfully.")

# -------------------------------------------------------
# Read Sheet into DataFrame
# -------------------------------------------------------

data = sheet.values

columns = next(data)

df = pd.DataFrame(data, columns=columns)

# -------------------------------------------------------
# Extract Customer List
# -------------------------------------------------------

customer_df = df.iloc[:, 0:2].copy()

customer_df.columns = ["Customer Code", "Customer Name"]

customer_df = customer_df.dropna(subset=["Customer Code"])

customer_df["Customer Code"] = customer_df["Customer Code"].astype(str)

customer_dict = dict(
    zip(
        customer_df["Customer Code"],
        customer_df["Customer Name"]
    )
)

# -------------------------------------------------------
# Find Customer Row
# -------------------------------------------------------

def find_customer_row(sheet, customer_code):
    """
    Returns the Excel row number of the selected customer.
    """

    for row in range(2, sheet.max_row + 1):

        value = sheet.cell(row=row, column=1).value

        if str(value).strip() == str(customer_code).strip():

            return row

    return None

# -------------------------------------------------------
# Find Date Column
# -------------------------------------------------------

def find_date_column(sheet, selected_date):
    """
    Returns the Excel column number for the selected day.
    """

    for col in range(1, sheet.max_column + 1):

        value = sheet.cell(row=1, column=col).value

        if value == selected_date:

            return col

    return None

# -------------------------------------------------------
# Entry Form
# -------------------------------------------------------

st.markdown("---")
st.subheader("Tea Collection Entry")

# Date
selected_date = st.selectbox(
    "Date",
    list(range(1, 32))
)

# Customer Row
col1, col2 = st.columns([1, 2])

with col1:
    customer_code = st.selectbox(
        "Customer Code",
        customer_df["Customer Code"]
    )

with col2:
    customer_name = customer_dict.get(customer_code, "")

    st.text_input(
        "Customer Name",
        value=customer_name,
        disabled=True
    )

# Amount
amount = st.number_input(
    "Amount (Kg)",
    min_value=0.0,
    step=0.5,
    format="%.2f"
)

# -------------------------------------------------------
# Submit Entry
# -------------------------------------------------------

if st.button("Submit Entry", use_container_width=True):

    customer_row = find_customer_row(sheet, customer_code)

    date_column = find_date_column(sheet, selected_date)

    if customer_row is None:

        st.error("Customer not found.")

    elif date_column is None:

        st.error("Date column not found.")

    else:

        # Update Excel Cell
        sheet.cell(
            row=customer_row,
            column=date_column
        ).value = amount

        # Reload dataframe
        data = sheet.values
        columns = next(data)
        df = pd.DataFrame(data, columns=columns)

        st.success("Entry saved successfully!")

        st.dataframe(
            df,
            use_container_width=True,
            height=500
        )

# -------------------------------------------------------
# Display Excel
# -------------------------------------------------------

st.markdown("---")

st.subheader("Current Excel Sheet")

table_placeholder = st.empty()

table_placeholder.dataframe(
    df,
    use_container_width=True,
    height=500
)
