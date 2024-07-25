import pandas as pd

# List of file paths provided by the user
file_paths = [
    '2024-03-19.xlsx',
    '2024-03-20.xlsx',
    '2024-03-21.xlsx',
    '2024-03-22.xlsx',
    '2024-03-23.xlsx',
    '2024-03-26.xlsx'
]

# Initialize an empty DataFrame to aggregate data
aggregated_data = pd.DataFrame()

# Loop through each file and concatenate the data
for file_path in file_paths:
    data = pd.read_excel(file_path)
    data = data[['Symbol', 'Last']].copy()
    data['Date'] = pd.to_datetime(file_path.split('/')[-1].split('.')[0])
    aggregated_data = pd.concat([aggregated_data, data], ignore_index=True)

# Calculate count of occurrences for each symbol
symbol_counts = aggregated_data['Symbol'].value_counts().reset_index()
symbol_counts.columns = ['Symbol', 'Count']

# Get the first 'Last' price and its corresponding date for each symbol
first_last_prices_dates = aggregated_data.groupby('Symbol').first().reset_index()
first_last_prices_dates = first_last_prices_dates[['Symbol', 'Last', 'Date']]
first_last_prices_dates.columns = ['Symbol', 'First_Last', 'First_Last_Date']

# Get the last 'Last' price and its corresponding date for each symbol
last_last_prices_dates = aggregated_data.groupby('Symbol').last().reset_index()
last_last_prices_dates = last_last_prices_dates[['Symbol', 'Last', 'Date']]
last_last_prices_dates.columns = ['Symbol', 'Last_Last', 'Last_Last_Date']

# Merge the dataframes with dates
merged_data_with_dates = pd.merge(symbol_counts, first_last_prices_dates, on='Symbol')
merged_data_with_dates = pd.merge(merged_data_with_dates, last_last_prices_dates, on='Symbol')

# Calculate percentage gain
merged_data_with_dates['% Gain'] = ((merged_data_with_dates['Last_Last'] - merged_data_with_dates['First_Last']) / merged_data_with_dates['First_Last']) * 100

# Reorder columns for better clarity
final_summary = merged_data_with_dates[['Symbol', 'Count', 'First_Last', 'First_Last_Date', 'Last_Last', 'Last_Last_Date', '% Gain']]

# Display the resulting dataframe with dates
print(final_summary)

# Save the final summary dataframe to an Excel file
output_file_path = 'Market_Scanning_Summary.xlsx'
final_summary.to_excel(output_file_path, index=False)