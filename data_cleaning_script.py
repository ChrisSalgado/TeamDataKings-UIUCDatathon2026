import datetime
import calendar
import pandas as pd


month_map = {name: i for i, name in enumerate(calendar.month_name) if name}


def output_dfs_dict(excel_chart_name: str):
    #Loads the DFs in one pass and returns a dictionary that indexes the DFs.
    call_center_labels = ['A', 'B', 'C', 'D']
    call_center_data_precisions = ['Daily', 'Interval']

    dfs = {}
    for current_call_center_name in call_center_labels:
        for call_center_data_precision in call_center_data_precisions:
            df = pd.read_excel(f"data_folder/{excel_chart_name}",
                                sheet_name=f"{current_call_center_name} - {call_center_data_precision}")
            dfs[f"{current_call_center_name} - {call_center_data_precision}"] = df
    return dfs

def insert_weekdays_and_sort_by_interval(dfs):
    # Identifies the weekdays of the week and writes it down.

    call_center_labels = ['A', 'B', 'C', 'D']
    call_center_data_precisions = ['Daily', 'Interval']
    with pd.ExcelWriter("data_folder/data_time_sorted.xlsx", mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        for current_call_center_name in call_center_labels:
            for call_center_data_precision in call_center_data_precisions:
                if call_center_data_precision == "Interval":
                    current_df = dfs[f"{current_call_center_name} - {call_center_data_precision}"]
                    current_df['Year'] = 2025
                    current_df['Month_English'] = current_df['Month']
                    current_df['Month'] = current_df['Month'].map(month_map)
                    current_df['Date'] = pd.to_datetime(current_df[['Year', 'Month', 'Day']])
                    current_df['WeekDay'] = current_df['Date'].dt.day_name()
                    current_df = current_df.sort_values(by=['Interval', 'Month', 'Day'], ascending=[True, True, True])
                    current_df['Date'] = current_df['Date'].dt.strftime("%d-%b-%Y")
                    current_df.to_excel(writer, sheet_name=f"{current_call_center_name} - {call_center_data_precision}", index=False)

def extrapolate_nans(dfs):
    # Identifies the weekdays of the week and writes it down.
    call_center_labels = ['A', 'B', 'C', 'D']
    call_center_data_precisions = ['Daily', 'Interval']
    with open('output.txt', 'w') as f:
        for current_call_center_name in call_center_labels:
            for call_center_data_precision in call_center_data_precisions:
                if call_center_data_precision == "Interval":
                    sheet_name = f"{current_call_center_name} - {call_center_data_precision}"
                    current_df = dfs[sheet_name]
                    for idx, row in current_df.iterrows():
                        nan_cols = row.index[row.isna()].tolist()
                        if nan_cols:
                            print(f"{sheet_name}, {idx}, {nan_cols}", file=f)

dfs = output_dfs_dict("data_time_sorted.xlsx")
extrapolate_nans(dfs)









                    
            





