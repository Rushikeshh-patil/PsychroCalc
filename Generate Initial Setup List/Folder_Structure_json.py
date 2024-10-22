import os
import json

def get_folder_structure(folder_path):
    """
    Recursively builds the folder structure as a nested dictionary.
    """
    folder_dict = {"name": os.path.basename(folder_path), "type": "folder", "contents": []}
    
    for root, dirs, files in os.walk(folder_path):
        # Calculate the relative path from the root folder to the current folder
        rel_path = os.path.relpath(root, folder_path)
        parent_folder = folder_dict
        
        # Traverse the nested structure to find the correct parent
        if rel_path != ".":
            for part in rel_path.split(os.sep):
                parent_folder = next(
                    item for item in parent_folder["contents"] 
                    if item["type"] == "folder" and item["name"] == part
                )
        
        # Add directories
        for dir_name in dirs:
            parent_folder["contents"].append({
                "name": dir_name,
                "type": "folder",
                "contents": []
            })
        
        # Add files
        for file_name in files:
            parent_folder["contents"].append({
                "name": file_name,
                "type": "file"
            })
    
    return folder_dict

def main():
    # Get the folder path from the user
    folder_path = input("Enter the folder path: ").strip()

    # Verify the path exists
    if not os.path.exists(folder_path):
        print(f"Error: The folder path '{folder_path}' does not exist.")
        return

    # Get the folder structure as a dictionary
    folder_structure = get_folder_structure(folder_path)

    # Save the dictionary as a JSON file in the same directory as the script
    output_path = os.path.join(os.path.dirname(__file__), "folder_structure.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(folder_structure, f, indent=4)

    print(f"Folder structure saved successfully to: {output_path}")

if __name__ == "__main__":
    main()
