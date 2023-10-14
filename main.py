from sharepointManager import SharePointFileManager
from config import site_url, library_name, folder_name
from config import client_id, client_secret
# or
from config import username, password

# Usage example in main.py:
if __name__ == "__main__":
    
    sharepoint = SharePointFileManager(site_url, username=username, password=password)
    # # or
    # sharepoint = SharePointManager(site_url, tenant_id, client_id, client_secret)

    # # get folder context
    sharepoint.set_folder_ctx(library_name, folder_name)

    # # Use as per the requirement

    # # Get files in the folder
    files = sharepoint.get_files()
    print("Files in folder:")
    for file in files:
        print(file.properties["Name"])
    
    # # Get folders in the folder
    folders = sharepoint.get_folders()
    print("\nFolders in folder:")
    for folder in folders:
        print(folder.properties["Name"])

    # # Create an empty folder in sharepoint
    # local_folder_name='folder'
    # sharepoint.create_sharepoint_folder(local_folder_name)
    
    # # Upload a file
    # local_file_path = "folder//sample1.txt"
    # sharepoint.upload_file(local_file_path)
    # print("File Uploaded ...")

    # # Multiple file uploader 
    # file_paths_to_upload = ["folder//sample1.txt", "folder//sample2.txt", "folder//sample1.txt"]
    # sharepoint.upload_multiple_files(file_paths_to_upload)

    # # Upload file without over ride
    # local_file_path = "folder//sample1.txt"
    # sharepoint.upload_file_without_override(local_file_path)

    # # Upload multiple files without override
    # file_paths_to_upload = ["folder//sample1.txt", "folder//sample2.txt", "folder//sample1.txt"]
    # sharepoint.upload_multiple_files_without_override(file_paths_to_upload)

    # # Delete a file
    # file_to_delete = "folder//sample1.txt"
    # sharepoint.delete_file(library_name)
    # print("Deleted Successfully...")

    # # Delete multiple files
    # file_names = ["folder//sample1.txt", "folder//sample2.txt"]
    # sharepoint.delete_multiple_files(file_names)

    # # Delete all files and folders in a context
    # sharepoint.delete_all_files_and_folders()

    # # Delete folder
    # folder = 'folder'
    # sharepoint.delete_entire_folder(folder)
    
    # # Get the details of recently deleted items
    # recently_deleted_items = sharepoint.get_recently_deleted_items(max_items=2)
    # for item in recently_deleted_items:
    #     print(item.properties["Title"])
    
    # # Restore the deleted files
    # sharepoint.recover_data(num_to_restore=1)