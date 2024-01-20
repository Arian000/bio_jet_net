import numpy as np
import pandas as pd
import joblib


def convert_xlsx_to_csv(xlsx_file, csv_folder, testname):
    
    xlsx = pd.ExcelFile(xlsx_file)

    # Get the list of sheet names in the Excel file
    sheet_names = xlsx.sheet_names

    # Initialize an empty DataFrame to hold the concatenated data
    all_data = pd.DataFrame()

    # Loop through each sheet, read data, and concatenate it
    for sheet_name in sheet_names:
        df = pd.read_excel(xlsx, sheet_name)
        df = df[df.columns[:2]]

        wavelength = df[df.columns[0]]
        value = df[df.columns[1]]
        
        wavelength = wavelength[5:]
        value = value[5:]

        temp_dict = {'wavelength' : wavelength, 'value' : value}
        cleaned_df = pd.DataFrame(temp_dict)
        all_data = pd.concat([all_data, cleaned_df], ignore_index=True)

    # Save the concatenated data as a single CSV file
    csv_file = f'{csv_folder}{testname}.csv'
    all_data.to_csv(csv_file, index=False)


def impactful_intervals(df):
    acc_impact = [0]
    imp = df['impact']
    for i in df.index:
        acc_impact.append(acc_impact[-1] + imp[i])

    intervals = []
    for i in range(len(df)):
        if not i % 100:
            print(i)
        for j in range(i + 1, min(i + 20, len(df))):
            impact = acc_impact[j+1] - acc_impact[i]
            intervals.append(((df.index[i], df.index[j]), impact))

    # Sorting the intervals by the second item in the tuple (the impact value)
    sorted_intervals = sorted(intervals, key=lambda x: x[1], reverse=True)  # Sorting in descending order

    return sorted_intervals


if __name__ == '__main__':
    df = pd.read_csv('interpreted_data.csv', index_col='feature_num')
    df = df.sort_index()

    intervals = impactful_intervals(df)
    print(len(intervals))
    impactfuls = intervals[ : 200] + intervals[-200 : ]

    # Uncomment the following line to save the intervals to a file
    joblib.dump(impactfuls, 'impactful_intervals.joblib')

    impactfuls = joblib.load('impactful_intervals.joblib')
    print(impactfuls)

    intervals = []
    impacts = []
    for tup in impactfuls:
        intervals.append(tup[0])
        impacts.append(tup[1])

    df = pd.DataFrame()
    df['interval'] = intervals
    df['impact'] = impacts
    df.to_csv('impactful_intervals.csv', index=False)