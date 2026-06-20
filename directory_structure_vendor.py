import os
import json
from pathlib import Path

# Define your categories here
ORIENTO_FILES = "ORIENTO_FILES"
NATRANS_FILE = "NATRANS_FILE"
FML_FILE = "FML_FILE"
VANITO_FILE = "VANITO_FILE"

def ask_user_for_category(filename):
    """
    Prompts the user to manually select a category for a given file.
    """
    print(f"\n--- Assign Category for File ---")
    print(f"📄 File: {filename}")
    print(f"1. {ORIENTO_FILES}")
    print(f"2. {NATRANS_FILE}")
    print(f"3. {FML_FILE}")
    print(f"4. {VANITO_FILE}")
    print(f"5. Skip / No Category")
    
    while True:
        choice = input("Select a category (1-5): ").strip()
        if choice == '1':
            return ORIENTO_FILES
        elif choice == '2':
            return NATRANS_FILE
        elif choice == '3':
            return FML_FILE
        elif choice == '4':
            return VANITO_FILE
        elif choice == '5':
            return None
        else:
            print("❌ Invalid choice. Please enter a number between 1 and 5.")

def build_dir_tree(dir_path):
    """
    Recursively builds a dictionary representing the directory structure
    and prompts the user for file categories.
    """
    path = Path(dir_path)
    tree = {
        "name": path.name if path.name else str(path),
        "type": "directory",
        "path": str(path.resolve()),
        "children": []
    }

    try:
        # Iterate through everything in the current directory
        for entry in path.iterdir():
            if entry.is_dir():
                # Recursively add subdirectories
                tree["children"].append(build_dir_tree(entry))
            else:
                # Ask the user to manually link the file to a category
                chosen_category = ask_user_for_category(entry.name)
                
                file_data = {
                    "name": entry.name,
                    "type": "file",
                    "path": str(entry.resolve()),
                    "size_bytes": entry.stat().st_size
                }
                
                # Only append category if one was chosen
                if chosen_category:
                    file_data["category"] = chosen_category
                
                # Add file details
                tree["children"].append(file_data)
                
    except PermissionError:
        # Handle folders that the script doesn't have permission to read
        tree["error"] = "Permission Denied"
        
    return tree

def export_structure_to_json(target_directory, output_json_file):
    """
    Generates the directory tree and saves it to a JSON file.
    """
    # Clean up user input (handles wrapped quotes from dragging/dropping folders)
    target_directory = target_directory.strip("'\"")
    
    if not target_directory:
        target_directory = "." # Default to current directory if empty

    path_obj = Path(target_directory)
    if not path_obj.exists():
        print(f"\n❌ Error: The directory '{target_directory}' does not exist. Please check the path and try again.")
        return
    
    if not path_obj.is_dir():
        print(f"\n❌ Error: '{target_directory}' is a file, not a directory.")
        return

    print(f"\n🔍 Scanning '{path_obj.resolve()}'... Preparing to prompt for file categories.")
    tree_data = build_dir_tree(path_obj)

    # Write the dictionary to a nicely formatted JSON file
    with open(output_json_file, 'w', encoding='utf-8') as f:
        json.dump(tree_data, f, indent=4, ensure_ascii=False)
        
    print(f"\n✨ Successfully saved directory structure to: {os.path.abspath(output_json_file)}")

# --- Interactive Execution ---
if __name__ == "__main__":
    print("=" * 50)
    print("  DIRECTORY TO JSON GENERATOR WITH CATEGORIES")
    print("=" * 50)
    
    # Ask the user for the directory path
    user_path = input("Enter the directory path to scan (or press Enter for current directory): ").strip()
    output_file = "directory_structure.json"

    export_structure_to_json(user_path, output_file)