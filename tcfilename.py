import os
import shutil
import re

# Get the directory path one level up from where the script is being run
current_dir = os.getcwd()
parent_dir = os.path.abspath(os.path.join(current_dir, ".."))

# Define the target folder name and path
target_folder = "time corrected"
target_folder_path = os.path.join(parent_dir, target_folder)

# Define the subfolder names and the corresponding strings to search in filenames
folder_keywords = {
    "D12": "D12",
    "D15": "D15",
    "D17": "D17",
    "D31": "D31",
    "D32": "D32",
    "G06": "G06",
    "G08": "G08",
    "G10": "G10",
    "G14": "G14",
    "G15": "G15",
    "G24": "G24"
}

# Function to build a dynamic regex pattern for matching variations
def build_dynamic_pattern(keyword):
    pattern = re.sub(r'(\d+)', r'[-_ ]?\1', keyword)  # Allow for optional characters between letter and number
    return re.compile(pattern, re.IGNORECASE)  # Case-insensitive match

# Check if the "time corrected" folder exists
if os.path.exists(target_folder_path) and os.path.isdir(target_folder_path):
    
    # Loop through all the files in the "time corrected" folder
    files = os.listdir(target_folder_path)
    
    for file_name in files:
        
        # Only look for .xlsx or .csv files
        if file_name.lower().endswith((".xlsx", ".csv")):
            
            # Convert "O" (capital letter O) to "0" (zero) for matching
            normalized_file_name = file_name.replace("O", "0")
            
            # Check if the file name contains any of the keywords using the dynamic regex pattern
            for folder_name, keyword in folder_keywords.items():
                pattern = build_dynamic_pattern(keyword)
                if pattern.search(normalized_file_name):
                    
                    # Construct the source file path
                    source_file = os.path.join(target_folder_path, file_name)
                    
                    # Construct the destination folder path
                    destination_folder = os.path.join(target_folder_path, folder_name)
                    
                    # Move the file to the respective subfolder
                    if not os.path.exists(destination_folder):
                        os.makedirs(destination_folder)
                    
                    shutil.move(source_file, destination_folder)
                    break  # Move on to the next file after a match is found
