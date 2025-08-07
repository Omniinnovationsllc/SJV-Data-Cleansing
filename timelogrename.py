import os
import pandas as pd
import shutil

# Define the path to the "time corrected" folder (one level up)
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
time_corrected_dir = os.path.join(base_dir, "time corrected")
log_file = "timestamp_log.csv"

# Full path to the timestamp_log.csv file
log_file_path = os.path.join(time_corrected_dir, log_file)

# Check if the file exists
if os.path.exists(log_file_path):
    # Read the CSV file and get the value of cell A2 (row 1, column 0 for zero-indexing)
    df = pd.read_csv(log_file_path)
    
    # Ensure the A2 cell exists
    if df.shape[0] > 1 and df.shape[1] > 0:
        a2_value = df.iloc[1, 0]  # A2 cell is located at (1, 0) in zero-indexing
        
        # Get the first 3 characters of the A2 value
        prefix = a2_value[:4]

        # Define the new file name
        new_file_name = f"timestamp_log_{prefix}.csv"
        new_file_path = os.path.join(time_corrected_dir, new_file_name)

        # Rename the file
        shutil.move(log_file_path, new_file_path)
        print(f"File renamed to {new_file_name}")
    else:
        print("The file does not have enough rows/columns for A2.")
else:
    print(f"{log_file} not found in {time_corrected_dir}.")
