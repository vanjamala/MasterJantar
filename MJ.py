import streamlit as st
import pandas as pd
from io import BytesIO

# Streamlit code for file upload
st.title("Učitajte MasterTeam evidenciju")  # Title in Croatian

# File uploader for users to upload their own files
uploaded_file = st.file_uploader("Učitajte Excel datoteku", type=["xls", "xlsx"])  # Upload instruction in Croatian

if uploaded_file is not None:
    # Load the uploaded Excel file
    df = pd.read_excel(uploaded_file, header=3)
    
    
    # Data processing steps (same as your original code)
    df = df.drop(columns=[df.columns[0]])  #  Drop first column if it's empty
    df = df.reset_index(drop=True)
    
    # Convert "Rbr" column to numeric, removing non-numeric rows
    df['Rbr'] = pd.to_numeric(df['Rbr'], errors='coerce')
    df = df.dropna(subset=['Rbr'])
    df = df[df['Rbr'] <= 1000]  # Drop rows where "Rbr" is > 1000
    df = df.reset_index(drop=True)
    
    # Remove columns that are "Unnamed" or empty
    df = df.loc[:, ~df.columns.str.contains('^Unnamed|^$', na=False)]
    df = df.reset_index(drop=True)

    # Identify unique persons based on the first two columns (Rbr and PREZIME i IME)
    personal_data_columns = ["Rbr", "PREZIME i IME"]
    
    # Extract day columns (Su 1 to Pe 31)
    day_columns = [col for col in df.columns if any(str(i) in col for i in range(1, 32))]
    
    # Melt the day columns into rows (long format)
    melted_data = pd.melt(df, id_vars=personal_data_columns, value_vars=day_columns,
                          var_name="Day", value_name="Value")
    
    # Clean the "Day" column
    melted_data['Day'] = melted_data['Day'].str.extract('(\d+)')

    # Save the transformed data to a new Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        melted_data.to_excel(writer, index=False, sheet_name="Sheet1")
    output.seek(0)

    # Provide a download button for the processed Excel file
    st.download_button(
        label="Preuzmite obrađenu Excel datoteku",  # Download instruction in Croatian
        data=output,
        file_name="transformed_data_stacked.xlsx",
        mime="application/vnd.ms-excel"
    )
