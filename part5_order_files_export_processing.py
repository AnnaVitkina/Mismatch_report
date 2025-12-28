import pandas as pd
import os


def process_order_files_export(file_path):
    """
    Process an Order Files Export Excel file from the input folder.
    Extracts only "Order file #" and "Order file name" columns.
    
    Args:
        file_path (str): Path to the file relative to the "input/" folder (e.g., "order_files_export.xlsx")
    
    Returns:
        dataframe: Processed pandas DataFrame with only "Order file #" and "Order file name" columns
    """
    # Construct full path from input folder
    input_folder = "input"
    full_path = os.path.join(input_folder, file_path)
    
    # Check if file exists
    if not os.path.exists(full_path):
        raise FileNotFoundError(f"File not found: {full_path}")
    
    # Read Excel file
    df = pd.read_excel(full_path)
    
    # Check if required columns exist
    required_cols = ['Order file #', 'Order file name']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        available_cols = list(df.columns)
        raise ValueError(
            f"Required columns not found in file: {missing_cols}\n"
            f"Available columns: {available_cols}"
        )
    
    # Select only the two required columns
    df_filtered = df[required_cols].copy()
    
    return df_filtered


# Example usage
#if __name__ == "__main__":
#    order_files_dataframe = process_order_files_export("Order_files_export.xls.xlsx")
#    print("DataFrame:")
#    print(order_files_dataframe)

