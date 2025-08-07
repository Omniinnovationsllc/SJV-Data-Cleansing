#!/usr/bin/env python3
"""
This script creates a set of subfolders inside an "averaged" folder located
one level up from where this script is run. It then scans all .xlsx files
in the "averaged" folder, and moves each file into the matching subfolder
if its filename (after normalization) contains one of the subfolder names.
Normalization steps:
  - Strip out any hyphens (“-”)
  - Convert letter “O” to digit “0”
  - Uppercase everything
"""

import os
import shutil
from pathlib import Path

def normalize_name(name: str) -> str:
    """
    Normalize a filename for matching:
      - Remove hyphens
      - Convert 'O' to '0'
      - Uppercase
    """
    return name.upper().replace('-', '').replace('O', '0')

def main():
    # Define the folder names to create and match against filenames
    folder_names = [
        "D12", "D15", "D17", "D31", "D32",
        "G06", "G10", "G14", "G15", "G24"
    ]
    
    # Determine the "averaged" directory path (one level up from script's cwd)
    script_dir = Path.cwd()
    averaged_dir = script_dir.parent / "averaged"
    
    # Create the "averaged" directory if it doesn't exist
    averaged_dir.mkdir(parents=True, exist_ok=True)
    
    # Create each of the subfolders inside "averaged"
    for name in folder_names:
        (averaged_dir / name).mkdir(exist_ok=True)
    
    # Iterate over all .xlsx files directly inside the "averaged" folder
    for file_path in averaged_dir.glob("*.xlsx"):
        filename = file_path.name
        normalized = normalize_name(filename)
        
        # Check each folder name for a match in the normalized filename
        for name in folder_names:
            if name in normalized:
                dest_folder = averaged_dir / name
                destination = dest_folder / filename
                try:
                    shutil.move(str(file_path), str(destination))
                    print(f"Moved '{filename}' to '{dest_folder.name}/'")
                except Exception as e:
                    print(f"Error moving '{filename}' to '{dest_folder.name}/': {e}")
                # Once moved, stop checking other folder names
                break

if __name__ == "__main__":
    main()
