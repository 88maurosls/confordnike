import streamlit as st
import pandas as pd

# Title of the Streamlit app
st.title("Excel Size Data Transformation")

# File uploader to allow the user to upload an Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read the uploaded Excel file
        excel_data = pd.ExcelFile(uploaded_file)

        # Display sheet names to let the user choose
        sheet_name = st.selectbox("Select the sheet to process", excel_data.sheet_names)

        # Load the selected sheet
        sheet_data = excel_data.parse(sheet_name)

        # Display the first few rows of the uploaded data
        st.write("Preview of the uploaded data:")
        st.dataframe(sheet_data.head())

        # Identify size columns (from "3.5" to "15")
        size_columns = sheet_data.loc[:, "3.5":"15"]

        # Other columns to keep constant for each size row
        other_columns = sheet_data.drop(columns=size_columns.columns)

        # Transform the table to create one row per size
        melted_data = sheet_data.melt(
            id_vars=other_columns.columns,
            value_vars=size_columns.columns,
            var_name="Size",
            value_name="Quantity"
        )

        # Remove rows with NaN quantities to keep only relevant size rows
        final_data = melted_data.dropna(subset=["Quantity"])

        # Remove rows where all columns except "Size" and "Quantity" are NaN
        final_data = final_data.dropna(how="all", subset=other_columns.columns)

        # Sort the data alphabetically based on the specified column
        sort_column = st.selectbox("Select the column to sort by", final_data.columns, index=0)
        final_data = final_data.sort_values(by=sort_column, ascending=True)

        # Display the transformed data
        st.write("Transformed Data:")
        st.dataframe(final_data)

        # Download button to download the transformed Excel file
        @st.cache_data
        def convert_df_to_excel(df):
            """Convert the DataFrame to Excel bytes."""
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Transformed_Data')
            return output.getvalue()

        excel_bytes = convert_df_to_excel(final_data)

        st.download_button(
            label="Download Transformed Data as Excel",
            data=excel_bytes,
            file_name="Transformed_Size_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")
