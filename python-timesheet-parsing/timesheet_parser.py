import pandas as pd
import calplot
import matplotlib.pyplot as plt
from pathlib import Path

def parse_timesheets(folder_path):
    try:      
        timesheet_data = pd.DataFrame(columns=["date","hours","week_ending","week_pto", "type_by_hours","category","notes"])
        flagged_sheets = set()
        holidays = set()
        issues_list = []

        for file_path in folder_path.iterdir():
            if file_path.is_file():
                try:
                    # timesheet data is always on the first sheet
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

                        # template changes moved the daily totals text over a column due to a merged cell  
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
                        issues_list.append(("File: " + file_path.name, "failed to find totals"))
                        flagged_sheets.add(file_path.name)
                    if week_pto == "":
                        issues_list.append(("File: " + file_path.name, "failed to find pto"))
                        flagged_sheets.add(file_path.name)

                    if file_path.name not in flagged_sheets:
                        weekly_data = pd.DataFrame({"date": mon_thru_fri_dates, "hours": mon_thru_fri_totals, 
                                                    "week_ending": week_ending, "week_pto": week_pto})
                        non_empty_dfs = [df for df in [weekly_data, timesheet_data] if not df.empty]
                        timesheet_data = pd.concat(non_empty_dfs, ignore_index=True)
                except:
                    issues_list.append(("File: " + file_path.name, "exception thrown while parsing"))
                    flagged_sheets.add(file_path.name)
        # Check for gaps in the date range
        full_range = pd.date_range(start=timesheet_data["date"].min(), end=timesheet_data["date"].max(), freq='D')
        actual_dates_set = set(timesheet_data['date'].dt.date)
        expected_dates_set = set(full_range.date)
        missing_dates = expected_dates_set - actual_dates_set

        dates_list = list(missing_dates)
        dates_list.sort()
        for date in dates_list:
            issues_list.append(("Date: " + date.strftime("%y%m%d"), "Missing"))

        print("-------------------\n\n\n")
        print("Unable to parse:")
        for flagged in flagged_sheets:
            print(flagged)
        print("Issues found: ")
        for issue in issues_list:
            print(issue)
        
        return (timesheet_data, holidays)

    except ValueError as e:
        print(f"Error: Could not read the Excel file. Details: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def set_category(timesheet_data, holidays, path_to_overrides):
    #deduce type by hours worked
    timesheet_data['type_by_hours'] = timesheet_data.apply(lambda row: type_by_hours(row['hours'], row['date']), axis=1)
    #set as starting value
    timesheet_data['category'] = timesheet_data['type_by_hours']
    #override with holidays and manual overrides
    timesheet_data.loc[timesheet_data["date"].isin(holidays), 'category'] = 'holiday'
    df = pd.read_csv(path_to_overrides)
    df['date'] = pd.to_datetime(df['date'])
    for index, row in df.iterrows():
        timesheet_data.loc[timesheet_data['date']==row['date'], 'category'] = row['category']
        timesheet_data.loc[timesheet_data['date']==row['date'], 'notes'] = row['source/reason']
    
    return timesheet_data


def type_by_hours(hours, date):
    if hours >= 6:
        return "workday"
    elif 6 > hours >= 1:
        return "partial workday"
    elif 1 > hours:
        if date.weekday() in [4, 5, 6]:
            return "weekend"
        else:
            return "pto"

def generate_hours_heatmap(timesheet_data):
    df = timesheet_data.set_index('date')

    calplot.calplot(df['hours'], edgecolor=None, cmap='YlGn')
    plt.show()

if __name__ == "__main__":
    folder_path = Path("./sample_timesheets")
    overrides = Path("sample_overrides.csv")
    output_file_name = "sample_schedule_output.csv"

    timesheet_data, holidays = parse_timesheets(folder_path)
    # generate_hours_heatmap(timesheet_data) # comment in to visually validate results

    result = set_category(timesheet_data, holidays, overrides)
    result.to_csv(output_file_name, index=False)
    print("Results saved to: " + output_file_name)