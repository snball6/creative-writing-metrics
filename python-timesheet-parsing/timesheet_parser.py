import pandas as pd

def parse_first_sheet_from_xlsx(file_path):
    """
    TODO DESCRIPTION

    Args:
        file_path (str): The path to the XLSX file.
    """
    try:

        # Read the first sheet from the Excel file into a pandas DataFrame
        df = pd.read_excel(file_path, sheet_name=0)

        # Optional: Print some basic information about the data
        print(f"Reading'{file_path}'")
        print("\nFirst 5 rows of the data:")
        # Use to_string() for better formatting in the console
        # print(df.head().to_string())

        mon_thru_fri_dates = df.iloc[3,6:11].tolist()
        print(mon_thru_fri_dates)
        # You can now iterate through the data
        # Example: iterate through rows and print a specific column value
        print("\nIterating through rows (example):")
        # for index, row in df.iterrows():
        #     print(row[6])

        return df

    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
    except ValueError as e:
        print(f"Error: Could not read the Excel file. Details: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# --- Example Usage ---
# Make sure you have an 'example.xlsx' file in the same directory, 
# or replace 'example.xlsx' with the correct path to your file.
if __name__ == "__main__":
    excel_file_path = "./sample_timesheets/SBAL241004.xlsx"
    dataframe = parse_first_sheet_from_xlsx(excel_file_path)
