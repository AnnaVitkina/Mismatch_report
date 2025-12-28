import pandas as pd
import os
from pathlib import Path


def save_dataframe_to_excel(df, output_filename, folder_name="partly_df"):
    output_folder = Path(__file__).parent / folder_name
    output_folder.mkdir(exist_ok=True)
    df.to_excel(output_folder / output_filename, index=False, engine='openpyxl')


def process_etof_file(file_path):
    """
    Process an ETOF Excel file from the input folder.
    
    Args:
        file_path (str): Path to the file relative to the "input/" folder (e.g., "etof_file.xlsx")
    
    Returns:
        tuple: (dataframe, list of column names)
            - dataframe: Processed pandas DataFrame with specified columns removed
            - list: List of column names in the processed dataframe
    """
    # Construct full path from input folder
    input_folder = "input"
    full_path = os.path.join(input_folder, file_path)
    
    # Read Excel file (skip first row)
    df_etofs = pd.read_excel(full_path, skiprows=1)
    
    # Rename duplicate columns
    new_column_names = {
        'Country code': 'Origin Country',
        'Postal code': 'Origin postal code',
        'Airport': 'Origin airport',
        'City': 'Origin city',
        'Country code.1': 'Destination Country',
        'Postal code.1': 'Destination postal code',
        'Airport.1': 'Destination airport',
        'City.1': 'Destination city',
    }
    df_etofs = df_etofs.rename(columns=new_column_names, inplace=False)
    

    columns_to_remove = ['Match', 'Approve', 'Calculation', 'State', 'Issue',
                         'Currency', 'Value', 'Currency.1', 'Value.1', 'Currency.2', 'Value.2']
    # Remove specified columns
    # Only remove columns that actually exist in the dataframe
    columns_to_drop = [col for col in columns_to_remove if col in df_etofs.columns]
    if columns_to_drop:
        df_etofs = df_etofs.drop(columns=columns_to_drop)
    
    # Get list of column names
    column_names = df_etofs.columns.tolist()

    def extract_country_code(country_string):
        """Extract the two-letter country code from a country string."""
        if isinstance(country_string, str) and ' - ' in country_string:
            return country_string.split(' - ')[0]
        return country_string

    df_etofs['Origin Country'] = df_etofs['Origin Country'].apply(extract_country_code)
    df_etofs['Destination Country'] = df_etofs['Destination Country'].apply(extract_country_code)

    def extract_carrier_agreement(agreement_string):
        """Extract the carrier agreement number (e.g., 'RA20220420022') from the full string.
        Input: 'RA20220420022 (v.12) - Active'
        Output: 'RA20220420022'
        """
        if isinstance(agreement_string, str):
            # Split by space and take the first part (the RA number)
            return agreement_string.split(' ')[0]
        return agreement_string

    if 'Carrier agreement #' in df_etofs.columns:
        df_etofs['Carrier agreement #'] = df_etofs['Carrier agreement #'].apply(extract_carrier_agreement)

    return df_etofs, column_names

if __name__ == "__main__":
    etof_dataframe, etof_column_names = process_etof_file('etofs.xlsx')
    save_dataframe_to_excel(etof_dataframe, "etof_processed.xlsx")
    print(etof_dataframe.head())
