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
    df_master = pd.read_excel(uploaded_masterteam, header=3)

    # Process MasterTeam file (same logic as before)
    df_master = df_master.drop(columns=[df_master.columns[0]])  
    df_master = df_master.reset_index(drop=True)
    df_master['Rbr'] = pd.to_numeric(df_master['Rbr'], errors='coerce')
    df_master = df_master.dropna(subset=['Rbr'])
    df_master = df_master[df_master['Rbr'] <= 1000]
    df_master = df_master.loc[:, ~df_master.columns.str.contains('^Unnamed|^$', na=False)]
    df_master = df_master.reset_index(drop=True)

    # Extract columns
    personal_data_columns = ["Rbr", "PREZIME i IME"]
    day_columns = [col for col in df_master.columns if any(str(i) in col for i in range(1, 32))]
    
    # Reshape data
    melted_master = pd.melt(df_master, id_vars=personal_data_columns, value_vars=day_columns,
                            var_name="Day", value_name="Value")
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
if uploaded_masterteam is not None and uploaded_jantar is not None and uploaded_pn is not None and st.button('Merge Data and Generate Report'):
    df_master = pd.read_excel(uploaded_masterteam, header=3)

    # Process MasterTeam file (same logic as before)
    df_master = df_master.drop(columns=[df_master.columns[0]])  
    df_master = df_master.reset_index(drop=True)
    df_master['Rbr'] = pd.to_numeric(df_master['Rbr'], errors='coerce')
    df_master = df_master.dropna(subset=['Rbr'])
    df_master = df_master[df_master['Rbr'] <= 1000]
    df_master = df_master.loc[:, ~df_master.columns.str.contains('^Unnamed|^$', na=False)]
    df_master = df_master.reset_index(drop=True)

    # Extract columns
    personal_data_columns = ["Rbr", "PREZIME i IME"]
    day_columns = [col for col in df_master.columns if any(str(i) in col for i in range(1, 32))]
    
    # Reshape data
    melted_master = pd.melt(df_master, id_vars=personal_data_columns, value_vars=day_columns,
                            var_name="Day", value_name="Value")
    melted_master['Day'] = melted_master['Day'].str.extract('(\d+)')
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
        label="Download Merged Report",
        data=output,
        file_name="merged_report.xlsx",
        mime="application/vnd.ms-excel"
    )    