import pandas as pd
from pathlib import Path

def parse_first_sheet_from_xlsx(folder_path):
    """
    TODO DESCRIPTION

    Args:
        folder_path (str): The path to the timesheet files.
    """
    try:      
        timesheet_data = pd.DataFrame(columns=["date","hours","week_ending","week_pto","category"])
        flagged_sheets = set()
        holidays = set()

        for file_path in folder_path.iterdir():
            if file_path.is_file():
                # timesheet data always on the first sheet
                df = pd.read_excel(file_path, sheet_name=0)
                print(f"Reading'{file_path}'")
                
                week_ending = df.iloc[1,2]
                mon_thru_fri_dates = df.iloc[3,6:11].tolist()
                mon_thru_fri_totals = []
                week_pto = ""
                for index, row in df.iterrows():
                    # template changes between 2023 and 2024 moved the daily totals text over a column due to a merged cell  
                    if(row.iloc[2]=="Daily totals:" or row.iloc[3]=="Daily totals:"):
                        mon_thru_fri_totals = row.iloc[6:11]

                    if(row.iloc[6]=="PTO:"):
                        week_pto = row.iloc[10]

                    # if("holiday" in row.iloc[2].lower()):
                    #     print('found')
                    #     pass
                        #check for column with 8 on this row, then get the corresponding date



                if len(mon_thru_fri_totals)<5:
                    flagged_sheets.add(file_path.name)
                    print("Failed to find totals")
                elif week_pto == "":
                    flagged_sheets.add(file_path.name)
                    print("Failed to find pto")
                else:
                    weekly_data = pd.DataFrame({"date": mon_thru_fri_dates, "hours": mon_thru_fri_totals, 
                                                "week_ending": week_ending, "week_pto": week_pto})
                    non_empty_dfs = [df for df in [weekly_data, timesheet_data] if not df.empty]
                    timesheet_data = pd.concat(non_empty_dfs, ignore_index=True)

        print(timesheet_data)
        print("unable to parse", flagged_sheets)

    except ValueError as e:
        print(f"Error: Could not read the Excel file. Details: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    folder_path = Path("./sample_timesheets")
    parse_first_sheet_from_xlsx(folder_path)
