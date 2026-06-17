import os
import json
from pathlib import Path

def build_dir_tree(dir_path):
    """
    Recursively builds a dictionary representing the directory structure.
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
                # Add file details
                tree["children"].append({
                    "name": entry.name,
                    "type": "file",
                    "path": str(entry.resolve()),
                    "size_bytes": entry.stat().st_size
                })
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

    print(f"\n🔍 Scanning '{path_obj.resolve()}'... Please wait.")
    tree_data = build_dir_tree(path_obj)

    # Write the dictionary to a nicely formatted JSON file
    with open(output_json_file, 'w', encoding='utf-8') as f:
        json.dump(tree_data, f, indent=4, ensure_ascii=False)
        
    print(f"✨ Successfully saved directory structure to: {os.path.abspath(output_json_file)}")

# --- Interactive Execution ---
if __name__ == "__main__":
    print("=" * 50)
    print("  DIRECTORY TO JSON GENERATOR")
    print("=" * 50)
    
    # Ask the user for the directory path
    user_path = input("Enter the directory path to scan (or press Enter for current directory): ").strip()
    output_file = "directory_structure.json"

    export_structure_to_json(user_path, output_file)