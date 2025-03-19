import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🎈 Provjera sati")
st.write("Provjeri sate rada.")

# Upload MasterTeam, Jantar and PN files
uploaded_masterteam = st.file_uploader("Učitajte MasterTeam evidenciju", type=["xls", "xlsx"])
uploaded_jantar = st.file_uploader("Učitajte Jantar Team datoteku", type=["xls", "xlsx"])
uploaded_pn = st.file_uploader("Učitajte datoteku službenih putovanja", type=["xls", "xlsx"])  # Upload instruction in Croatian

# Process MasterTeam file
if uploaded_masterteam is not None and st.button("Obradi MasterTeam"):
    # df_master is loaded from the Excel file
    df_master = pd.read_excel(uploaded_masterteam, header=3)
    df_master = df_master.drop(columns=[df_master.columns[0]])

    # Reset index after dropping rows to avoid index gaps
    df_master = df_master.reset_index(drop=True)

    # Remove rows where the "Rbr" column is not a number
    df_master['Rbr'] = pd.to_numeric(df_master['Rbr'], errors='coerce')  # Convert "Rbr" to numeric, invalid entries become NaN
    df_master = df_master.dropna(subset=['Rbr'])  # Drop rows where "Rbr" is NaN

    # Reset the index after dropping rows
    df_master = df_master.reset_index(drop=True)

    # Drop rows where "Rbr" or any numeric columns are greater than 1000
    df_master = df_master[df_master['Rbr'] <= 1000]

    # Reset the index after dropping rows
    df_master = df_master.reset_index(drop=True)

    # Drop columns with names that contain "Unnamed" or are blank
    df_master = df_master.loc[:, ~df_master.columns.str.contains('^Unnamed|^$', na=False)]

    # Reset the index after dropping the columns
    df_master = df_master.reset_index(drop=True)

    # Identify unique persons based on the first two columns (Rbr and PREZIME i IME)
    personal_data_columns = ["Rbr", "PREZIME i IME"]

    # Extract only the day columns (Su 1 to Pe 31)
    day_columns = [col for col in df_master.columns if any(str(i) in col for i in range(1, 32))]

    # Melt the day columns into rows (long format)
    melted_master = pd.melt(df_master, id_vars=personal_data_columns, value_vars=day_columns,
                        var_name="Day", value_name="Value")

    # Clean up the "Day" column to only include the day number (e.g., '1', '2', etc.)
    melted_master['Day'] = melted_master['Day'].str.extract('(\d+)')

    # Save as Excel
    output_master = BytesIO()
    with pd.ExcelWriter(output_master, engine='xlsxwriter') as writer:
        melted_master.to_excel(writer, index=False, sheet_name="MasterTeam")
    output_master.seek(0)

    # Download button
    st.download_button("Preuzmite obrađenu MasterTeam datoteku", data=output_master,
                       file_name="transformed_masterteam.xlsx", mime="application/vnd.ms-excel")


# Process Jantar file
# ---- PROCESS JANTAR FILE ----
if uploaded_jantar and st.button("Obradi Jantar"):
    df_J = pd.read_excel(uploaded_jantar, header=None)

    metadata = {}
    all_data = []
    current_section = None

    # Loop through rows
    for index, row in df_J.iterrows():
        first_col = str(row.iloc[0]).strip()
        second_col = row.iloc[1] if len(row) > 1 else None

        if first_col in ["Korisnik", "Razdoblje", "Odjel", "Raspored", "Kartica korisnika"]:
            metadata[first_col] = second_col
        elif first_col in ["Suma", "Saldo za razdoblje", "Radna obveza"]:
            metadata[first_col] = second_col
        elif first_col in ["Prekovremeno", "Stimulacija", "Stanje", "Prijenos", "Godišnji", "Stari godišnji",
                           "Dvokratni rad", "Broj obroka", "Broj prijevoza"]:
            metadata[first_col] = second_col
        elif first_col in ["Statistika", "Vrijeme", "Ukupno", "Vremenski razrez", "Vrijeme"]:
            continue  # Skip these
        elif not first_col:
            continue
        elif first_col == "Dan":
            current_section = {**metadata}
            continue
        elif current_section:
            row_data = row.tolist()
            while len(row_data) < 8:
                row_data.append(None)

            if all(x is None or pd.isna(x) for x in row_data[:8]):
                continue

            combined_row = {
                **current_section,
                "Dan": row_data[0],
                "Datum": row_data[1],
                "Početak": row_data[2],
                "Unnamed 1": row_data[3],
                "Kraj": row_data[4],
                "Unnamed 2": row_data[5],
                "Ukupno": row_data[6],
                "Statistika": row_data[7]
            }
            all_data.append(combined_row)

    df_J_cleaned = pd.DataFrame(all_data)
    df_J_cleaned['Korisnik'] = df_J_cleaned['Korisnik'].str.strip().str.upper()
    df_J_cleaned['Datum'] = df_J_cleaned['Datum'].fillna(method='ffill')

    # Save & Download
    output_jantar = BytesIO()
    with pd.ExcelWriter(output_jantar, engine='xlsxwriter') as writer:
        df_J_cleaned.to_excel(writer, index=False, sheet_name="Jantar")
    output_jantar.seek(0)

    st.download_button("Preuzmite obrađenu Jantar datoteku", data=output_jantar,
                       file_name="transformed_jantar.xlsx", mime="application/vnd.ms-excel")

if uploaded_pn is not None and st.button("Obradite datoteku putnih naloga"):  # Combine the file upload and button click
    # Load the uploaded Excel file and skip the first 3 rows
    df_pn = pd.read_excel(uploaded_pn, header=3)

    # Remove the "SVEUKUPNO" rows
    df_pn = df_pn[df_pn['Broj PN\n'] != 'SVEUKUPNO']

    # Convert 'Dat. Polaska' and 'Dat. Povratka' to datetime format (if they are not already in datetime format)
    df_pn["Dat. Polaska"] = pd.to_datetime(df_pn["Dat. Polaska"], errors='coerce')
    df_pn["Dat. Povratka"] = pd.to_datetime(df_pn["Dat. Povratka"], errors='coerce')

    # Expand each row into multiple rows for each day in the GO period
    expanded_rows = []
    for _, row in df_pn.iterrows():
        # Generate all dates between 'Dat. Polaska' and 'Dat. Povratka'
        if pd.notna(row["Dat. Polaska"]) and pd.notna(row["Dat. Povratka"]):
            date_range = pd.date_range(row["Dat. Polaska"], row["Dat. Povratka"])  # Generate all dates
            for date in date_range:
                expanded_rows.append({
                    "Prezime Ime": row['Prezime i ime'],  # Use the correct column name
                    "Datum": date.strftime("%d.%m.%Y"),
                    "Razlog odsutnosti": row["Zadatak službenog puta"]
                })

    # Convert the list of rows into a DataFrame
    df_expanded = pd.DataFrame(expanded_rows)

    # Provide a download button for the processed Excel file
    output_pn = BytesIO()
    with pd.ExcelWriter(output_pn, engine='xlsxwriter') as writer:
        df_expanded.to_excel(writer, index=False, sheet_name="Processed Data")
    output_pn.seek(0)

    # Provide a download button for the processed data
    st.download_button(
        label="Preuzmite obrađenu datoteku putnih naloga",  # Download instruction in Croatian
        data=output_pn,
        file_name="processed_pn_data.xlsx",
        mime="application/vnd.ms-excel"
    )

# Check if all three files are uploaded and the button is clicked
if uploaded_masterteam is not None and uploaded_jantar is not None and uploaded_pn is not None and st.button('Spoji podatke i pripremi izvještaj'):
    df_master = pd.read_excel(uploaded_masterteam, header=3)
    # df_master is loaded from the Excel file
    df_master = df_master.drop(columns=[df_master.columns[0]])

    # Reset index after dropping rows to avoid index gaps
    df_master = df_master.reset_index(drop=True)

    # Remove rows where the "Rbr" column is not a number
    df_master['Rbr'] = pd.to_numeric(df_master['Rbr'], errors='coerce')  # Convert "Rbr" to numeric, invalid entries become NaN
    df_master = df_master.dropna(subset=['Rbr'])  # Drop rows where "Rbr" is NaN

    # Reset the index after dropping rows
    df_master = df_master.reset_index(drop=True)

    # Drop rows where "Rbr" or any numeric columns are greater than 1000
    df_master = df_master[df_master['Rbr'] <= 1000]

    # Reset the index after dropping rows
    df_master = df_master.reset_index(drop=True)

    # Drop columns with names that contain "Unnamed" or are blank
    df_master = df_master.loc[:, ~df_master.columns.str.contains('^Unnamed|^$', na=False)]

    # Reset the index after dropping the columns
    df_master = df_master.reset_index(drop=True)

    # Identify unique persons based on the first two columns (Rbr and PREZIME i IME)
    personal_data_columns = ["Rbr", "PREZIME i IME"]

    # Extract only the day columns (Su 1 to Pe 31)
    day_columns = [col for col in df_master.columns if any(str(i) in col for i in range(1, 32))]

    # Melt the day columns into rows (long format)
    melted_master = pd.melt(df_master, id_vars=personal_data_columns, value_vars=day_columns,
                        var_name="Day", value_name="Value")

    # Clean up the "Day" column to only include the day number (e.g., '1', '2', etc.)
    melted_master['Day'] = melted_master['Day'].str.extract('(\d+)')

    # The melted_master dataframe is now ready for further processing


    # Extract columns
    personal_data_columns = ["Rbr", "PREZIME i IME"]
    day_columns = [col for col in df_master.columns if any(str(i) in col for i in range(1, 32))]
    
    # Reshape data
    melted_master = pd.melt(df_master, id_vars=personal_data_columns, value_vars=day_columns,
                            var_name="Day", value_name="Value")
    melted_master['Day'] = melted_master['Day'].str.extract('(\d+)')
    #transform Jantar data
    df_J = pd.read_excel(uploaded_jantar, header=None)

    metadata = {}
    all_data = []
    current_section = None

    # Loop through rows
    for index, row in df_J.iterrows():
        first_col = str(row.iloc[0]).strip()
        second_col = row.iloc[1] if len(row) > 1 else None

        if first_col in ["Korisnik", "Razdoblje", "Odjel", "Raspored", "Kartica korisnika"]:
            metadata[first_col] = second_col
        elif first_col in ["Suma", "Saldo za razdoblje", "Radna obveza"]:
            metadata[first_col] = second_col
        elif first_col in ["Prekovremeno", "Stimulacija", "Stanje", "Prijenos", "Godišnji", "Stari godišnji",
                           "Dvokratni rad", "Broj obroka", "Broj prijevoza"]:
            metadata[first_col] = second_col
        elif first_col in ["Statistika", "Vrijeme", "Ukupno", "Vremenski razrez", "Vrijeme"]:
            continue  # Skip these
        elif not first_col:
            continue
        elif first_col == "Dan":
            current_section = {**metadata}
            continue
        elif current_section:
            row_data = row.tolist()
            while len(row_data) < 8:
                row_data.append(None)

            if all(x is None or pd.isna(x) for x in row_data[:8]):
                continue

            combined_row = {
                **current_section,
                "Dan": row_data[0],
                "Datum": row_data[1],
                "Početak": row_data[2],
                "Unnamed 1": row_data[3],
                "Kraj": row_data[4],
                "Unnamed 2": row_data[5],
                "Ukupno": row_data[6],
                "Statistika": row_data[7]
            }
            all_data.append(combined_row)

    df_J_cleaned = pd.DataFrame(all_data)
    df_J_cleaned['Korisnik'] = df_J_cleaned['Korisnik'].str.strip().str.upper()
    df_J_cleaned['Datum'] = df_J_cleaned['Datum'].fillna(method='ffill')
    
    # Load the uploaded Excel file and skip the first 3 rows
    df_pn = pd.read_excel(uploaded_pn, header=3)

    # Remove the "SVEUKUPNO" rows
    df_pn = df_pn[df_pn['Broj PN\n'] != 'SVEUKUPNO']

    # Convert 'Dat. Polaska' and 'Dat. Povratka' to datetime format (if they are not already in datetime format)
    df_pn["Dat. Polaska"] = pd.to_datetime(df_pn["Dat. Polaska"], errors='coerce')
    df_pn["Dat. Povratka"] = pd.to_datetime(df_pn["Dat. Povratka"], errors='coerce')

    # Expand each row into multiple rows for each day in the GO period
    expanded_rows = []
    for _, row in df_pn.iterrows():
        # Generate all dates between 'Dat. Polaska' and 'Dat. Povratka'
        if pd.notna(row["Dat. Polaska"]) and pd.notna(row["Dat. Povratka"]):
            date_range = pd.date_range(row["Dat. Polaska"], row["Dat. Povratka"])  # Generate all dates
            for date in date_range:
                expanded_rows.append({
                    "Prezime Ime": row['Prezime i ime'],  # Use the correct column name
                    "Datum": date.strftime("%d.%m.%Y"),
                    "Razlog odsutnosti": row["Zadatak službenog puta"]
                })

    # Convert the list of rows into a DataFrame
    df_expanded = pd.DataFrame(expanded_rows)

    # Ensure 'Datum' in df_J_cleaned is in datetime format
    df_J_cleaned["Datum"] = pd.to_datetime(df_J_cleaned["Datum"], dayfirst=True)

    # Convert year and month to integers to avoid decimal points in the date string
    year = int(df_J_cleaned["Datum"].dt.year.iloc[0])
    month = int(df_J_cleaned["Datum"].dt.month.iloc[0])

    # Create a full date in melted_data by combining year, month, and "Day"
    melted_master["Full_Date"] = pd.to_datetime(melted_master["Day"].astype(str) + f"-{month}-{year}", format="%d-%m-%Y")

    # Merge both DataFrames on Employee Name and Date
    merged_df = melted_master.merge(df_J_cleaned, left_on=["PREZIME i IME", "Full_Date"], right_on=["Korisnik", "Datum"], how="left")

    # Convert both columns to datetime format
    df_expanded["Datum"] = pd.to_datetime(df_expanded["Datum"], dayfirst=True, errors="coerce")
    merged_df["Full_Date"] = pd.to_datetime(merged_df["Full_Date"], dayfirst=True, errors="coerce")

    # Now merge the dataframes
    merged_result = pd.merge(
        merged_df, df_expanded,
        left_on=["PREZIME i IME", "Full_Date"],
        right_on=["Prezime Ime", "Datum"],
        how="outer"
    )
    # Keep only the required columns
    merged_result = merged_result[["PREZIME i IME", "Full_Date", "Razlog odsutnosti", "Value", "Statistika"]]
        
    # Display the merged result
    st.write(merged_result)
        
    # Allow downloading the merged data
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        merged_result.to_excel(writer, index=False, sheet_name="Merged Report")
    output.seek(0)

    st.download_button(
        label="Preuzmi spojene tablice",
        data=output,
        file_name="merged_report.xlsx",
        mime="application/vnd.ms-excel"
    )   

    # First report (1. Odsutni prema Jantaru)

    # Make a copy of the merged_result to preserve the original DataFrame
    merged_result_copy = merged_result.copy()

    # Filter merged_result_copy to take the first non-null value for 'Statistika' for each person and date
    merged_result_copy['Statistika'] = merged_result_copy.groupby(['PREZIME i IME', 'Full_Date'])['Statistika'].transform(
        lambda x: x.dropna().iloc[0] if not x.dropna().empty else None
    )

    # Filter 'Statistika' to include only 'Odsutan' or no value (NaN)
    filtered_report_1 = merged_result_copy[
        (merged_result_copy['Statistika'] == 'Odsutan') | 
        (merged_result_copy['Statistika'].isna())
    ]

    # Convert 'Value' column to numeric, setting errors='coerce' to turn non-numeric values into NaN
    filtered_report_1['Value'] = pd.to_numeric(filtered_report_1['Value'], errors='coerce')
    # Filter 'Razlog odsutnosti' to be NaN
    filtered_report_1 = filtered_report_1[filtered_report_1['Razlog odsutnosti'].isna()]
    # Filter 'Value' to be numeric (remove NaN values)
    filtered_report_1 = filtered_report_1[filtered_report_1['Value'].notna()]
    # Get the last date in df_J_cleaned
    last_date_in_jantar = df_J_cleaned['Datum'].max()

    # Filter 'Full_Date' to be less than or equal to the max date from df_J_cleaned
    filtered_report_1 = filtered_report_1[filtered_report_1['Full_Date'] <= last_date_in_jantar]

    # Display filtered report 1
    if filtered_report_1.empty:
        st.write("⚠️ Filtered report is empty!")
    st.write(filtered_report_1)

    # Allow downloading the filtered report 1
    output_filtered_1 = BytesIO()
    with pd.ExcelWriter(output_filtered_1, engine='xlsxwriter') as writer:
        filtered_report_1.to_excel(writer, index=False, sheet_name="1. Odsutni prema Jantaru")
    output_filtered_1.seek(0)

    st.download_button(
        label="Preuzmi 1. Odsutni prema Jantaru",
        data=output_filtered_1,
        file_name="1_odsutni_prema_jantaru.xlsx",
        mime="application/vnd.ms-excel"
    )
    # Second report (1. Odsutni prema MasterTeam)

    # Filter merged_result to take the first non-null value for 'Statistika' for each person and date
    #merged_result['Statistika'] = merged_result.groupby(['PREZIME i IME', 'Full_Date'])['Statistika'].transform(
    #    lambda x: x.dropna().iloc[0] if not x.dropna().empty else None
    #)

    # Get the last date in df_J_cleaned (same as in the first report)
    last_date_in_jantar = df_J_cleaned['Datum'].max()

    filtered_report_2 = merged_result[
        (merged_result['Statistika'].notna()) &  # Statistika must not be NaN
        (merged_result['Statistika'] != 'Odsutan') &  # Statistika must not be 'Odsutan'
        (merged_result['Statistika'] != 'Vikend') # Statistika must not be Vikend
    ]

    # Keep only non-numeric values in 'Value'
    # We use `apply(pd.to_numeric, errors='coerce')` to attempt converting the 'Value' to numeric,
    # and keep the original non-numeric values by filtering rows where the converted value is NaN.
    filtered_report_2['Value_is_numeric'] = pd.to_numeric(filtered_report_2['Value'], errors='coerce').notna()

    # Filter rows where 'Value' is not numeric (the conversion resulted in NaN)
    filtered_report_2_non_numeric_value = filtered_report_2[~filtered_report_2['Value_is_numeric']]

    # Get the last date in df_J_cleaned
    last_date_in_jantar = df_J_cleaned['Datum'].max()

    # Filter 'Full_Date' to be less than or equal to the max date from df_J_cleaned
    filtered_report_2_non_numeric_value = filtered_report_2_non_numeric_value[
        filtered_report_2_non_numeric_value['Full_Date'] <= last_date_in_jantar
    ]

    # Display filtered report 2
    if filtered_report_2_non_numeric_value.empty:
        st.write("⚠️ Filtered report is empty!")
    st.write(filtered_report_2_non_numeric_value)

    # Allow downloading the filtered report 2
    output_filtered_2 = BytesIO()
    with pd.ExcelWriter(output_filtered_2, engine='xlsxwriter') as writer:
        filtered_report_2_non_numeric_value.to_excel(writer, index=False, sheet_name="1. Odsutni prema MasterTeam")
    output_filtered_2.seek(0)

    st.download_button(
        label="Preuzmi 1. Odsutni prema MasterTeam",
        data=output_filtered_2,
        file_name="1_odsutni_prema_masterteam.xlsx",
        mime="application/vnd.ms-excel"
    )