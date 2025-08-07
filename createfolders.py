import os

# Get the directory path one level up from where the script is being run
parent_dir = os.path.abspath(os.path.join(os.getcwd(), ".."))

# Define the target folder name
target_folder = "time corrected"

# Define the subfolders you want to create
subfolders = ["D12", "D15", "D17", "D31", "D32", "G06", "G08", "G10", "G14", "G15", "G24"]

# Construct the path to the "time corrected" folder
target_folder_path = os.path.join(parent_dir, target_folder)

# Check if the "time corrected" folder exists
if os.path.exists(target_folder_path) and os.path.isdir(target_folder_path):
    # Create the subfolders inside the "time corrected" folder
    for subfolder in subfolders:
        subfolder_path = os.path.join(target_folder_path, subfolder)
        os.makedirs(subfolder_path, exist_ok=True)
        print(f"Created: {subfolder_path}")
else:
    print(f"'{target_folder}' folder not found in {parent_dir}")
