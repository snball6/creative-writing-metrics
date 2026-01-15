import pandas as pd
from pathlib import Path

def parse_timesheets(folder_path, output_file_name):
    """
    Iterate through a folder of Menlo timesheets to extract hours worked and determine holidays, workdays, and pto

    Args:
        folder_path (Path): The path folder containing only timesheets
        output_file_name (str): Name for output file
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

                # template changes have dates in one of two rows
                if df.iloc[3,4] == "Sat":
                    mon_thru_fri_dates = df.iloc[4,4:11].tolist()
                else:
                    mon_thru_fri_dates = df.iloc[3,4:11].tolist()

                mon_thru_fri_totals = []
                week_pto = ""
                for index, row in df.iterrows():
                    # template changes between 2023 and 2024 moved the daily totals text over a column due to a merged cell  
                    if(row.iloc[2]=="Daily totals:" or row.iloc[3]=="Daily totals:"):
                        mon_thru_fri_totals = row.iloc[4:11].tolist()

                    if(row.iloc[6]=="PTO:"):
                        week_pto = row.iloc[10]

                    if("holiday" in str(row.iloc[2]).lower()):
                        mon_thru_fri_holiday = row.iloc[4:11].tolist()
                        for index, cell_value in enumerate(mon_thru_fri_holiday):
                            if cell_value > 0:
                                holidays.add(mon_thru_fri_dates[index])

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

        print("unable to parse", flagged_sheets)

        result = set_category(timesheet_data, holidays)
        result.to_csv(output_file_name, index=False)


    except ValueError as e:
        print(f"Error: Could not read the Excel file. Details: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def set_category(timesheet_data, holidays):
    timesheet_data['category'] = timesheet_data.apply(lambda row: category_by_time(row['hours'], row['date']), axis=1)
    timesheet_data.loc[timesheet_data['date'].isin(holidays), 'category'] = 'holiday'
    return timesheet_data


def category_by_time(hours, date):
    if hours >= 6:
        return "workday"
    elif 6 > hours >= 1:
        return "partial"
    elif 1 > hours:
        if date.weekday() in [4, 5, 6]:
            return "weekend"
        else:
            return "pto"



if __name__ == "__main__":
    folder_path = Path("./actual_timesheets")
    output_file_name = "output.csv"
    parse_timesheets(folder_path, output_file_name)
    print("Done!")