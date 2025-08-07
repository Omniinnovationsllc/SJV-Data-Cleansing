#!/usr/bin/env python3
"""
This script scans all .xlsx files in the "original" folder (one level up from
where the script is run), looks for any of the specified folder identifiers
in each filename, and moves the file into the matching subfolder inside
"original". The subfolders are assumed to already exist.
"""

import os
import shutil
from pathlib import Path

def main():
    # Define the folder names to match against filenames
    folder_names = [
        "D12", "D15", "D17", "D31", "D32",
        "G06", "G10", "G14", "G15", "G24"
    ]
    
    # Determine the "original" directory path (one level up from script's cwd)
    script_dir = Path.cwd()
    original_dir = script_dir.parent / "original"
    
    # Verify that the "original" directory exists
    if not original_dir.is_dir():
        print(f"Error: '{original_dir}' does not exist or is not a directory.")
        return
    
    # Iterate over all .xlsx files directly inside the "original" folder
    for file_path in original_dir.glob("*.xlsx"):
        filename = file_path.name
        # Check each folder name for a match in the filename
        for name in folder_names:
            if name in filename:
                destination_dir = original_dir / name
                if not destination_dir.is_dir():
                    print(f"Warning: target folder '{destination_dir}' does not exist; skipping.")
                    break
                destination = destination_dir / filename
                try:
                    shutil.move(str(file_path), str(destination))
                    print(f"Moved '{filename}' to '{name}/'")
                except Exception as e:
                    print(f"Error moving '{filename}' to '{name}/': {e}")
                # Once moved (or attempted), stop checking other folder names
                break

if __name__ == "__main__":
    main()
