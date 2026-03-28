import datetime as dt
import calendar
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import statsmodels.api as sm
import numpy as np
import statsmodels.api as sm
import statsmodels.formula.api as smf
import numpy as np


month_map = {name: i for i, name in enumerate(calendar.month_name) if name}


def output_dfs_dict(excel_chart_name: str):
    #Loads the DFs in one pass and returns a dictionary that indexes the DFs.

    call_center_labels = ['A', 'B', 'C', 'D']
    call_center_data_precisions = ['Daily', 'Interval']
    dicts = {}
    for current_call_center_name in call_center_labels:
        for call_center_data_precision in call_center_data_precisions:
            df = pd.read_excel(f"data_folder/{excel_chart_name}",
                                sheet_name=f"{current_call_center_name} - {call_center_data_precision}")
            dicts[f"{current_call_center_name} - {call_center_data_precision}"] = df
    return dicts

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


def plot_call_volume_against_interval_with_weekday(dfs):
    df = dfs["A - Interval"]
    df['Interval'] = pd.to_timedelta(df['Interval'], errors='coerce')
    print(df['Interval'])

    sns.scatterplot(
        data=df,
        x='Interval',
        y='Call Volume',
        hue='WeekDay'   # <-- color by this column
    )
    plt.show()








dfs = output_dfs_dict("data_time_sorted.xlsx")

df_a = dfs["A - Interval"]

# Prepare data
df = df_a[['Call Volume', 'Interval', 'WeekDay']].dropna()

# Convert Call Volume to int (counts)
df['Call Volume'] = df['Call Volume'].astype(int)

# Fit Poisson Regression model
# Using a formula: Call Volume as response, Interval and WeekDay as categorical predictors
model = smf.poisson("Q('Call Volume') ~ C(Interval) + C(WeekDay)", data=df).fit()

# Summary of results
print(model.summary())

# Check for overdispersion: Mean vs Variance
mean_val = df['Call Volume'].mean()
var_val = df['Call Volume'].var()
print(f"\nMean Call Volume: {mean_val:.2f}")
print(f"Variance Call Volume: {var_val:.2f}")

if var_val > mean_val:
    print("\nData shows signs of overdispersion (Variance > Mean). A Negative Binomial model might be better.")

new_data = {'Interval': '10:00:00', 'WeekDay': 'Monday'}
predicted_count = model.predict(new_data)


print(predicted_count)



# plot_call_volume_against_interval_with_weekday(dfs)



                    
            





