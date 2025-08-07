import os
import pandas as pd

def count_timestamps_per_day(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path, skiprows=1, header=None)
    
    # Ensure the first column is in datetime format
    df[0] = pd.to_datetime(df[0], format='%Y-%m-%d %H:%M:%S')
    
    # Group by date and count the number of occurrences
    counts = df[0].dt.date.value_counts().sort_index()
    
    # Print the results
    print(f"Results for {file_path}:")
    for date, count in counts.items():
        print(f"{date}: {count} timestamps")
    print("\n")

if __name__ == "__main__":
    # Determine the path to the "averaged" folder
    current_directory = os.path.dirname(os.path.abspath(__file__))
    averaged_folder = os.path.join(current_directory, '..', 'averaged')

    # Get a list of all Excel files in the "averaged" folder
    file_paths = [os.path.join(averaged_folder, file) for file in os.listdir(averaged_folder) if file.endswith(('.xlsx', '.xls'))]
    
    if file_paths:
        for file_path in file_paths:
            count_timestamps_per_day(file_path)
    else:
        print("No Excel files found in the 'averaged' folder.")
