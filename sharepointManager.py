from sharepointConnector import SharePointConnection
import os

class SharePointFileManager(SharePointConnection):
    """
    SharePointFileManager allows you to manage files and folders in a SharePoint document library.

    This class extends SharePointConnection to provide a high-level interface for working with SharePoint
    document libraries, including operations such as uploading, deleting, and restoring files and folders.

    Author:
    Pavan Kumar Gattupalli
    LinkedIn: https://www.linkedin.com/in/g-pavan-kumar/
    GitHub: https://github.com/g-pavan

    Parameters:
    - site_url (str): The URL of the SharePoint site.
    - username (str, optional): The username for authentication. If not provided, the connection will use client credentials.
    - password (str, optional): The password for authentication.
    - client_id (str, optional): The client ID for OAuth-based authentication.
    - client_secret (str, optional): The client secret for OAuth-based authentication.

    Attributes:
    - connection: An instance of SharePointConnection to handle the low-level SharePoint API operations.
    - existing_files (set): A set containing the names of existing SharePoint files in the current library and folder context.
    - library_name (str): The name of the SharePoint library you are working with.
    - folder_name (str): The name of the subfolder within the library.
    - folder: The SharePoint folder object representing the current context.

    Example Usage:
    connection = SharePointConnection("your_site_url", username="your_username", password="your_password")
    file_manager = SharePointFileManager(connection)
    file_manager.set_folder_ctx("Shared Documents", "Subfolder")
    files = file_manager.get_files()
    """
    def __init__(self, site_url, relative_url, username=None, password=None, client_id=None, client_secret=None):
        self.connection = SharePointConnection(site_url, username, password, client_id, client_secret)
        self.relative_url = relative_url
        self.existing_files = set()
        self.library_name = None
        self.folder_name = None
        self.folder = None
    
    def set_folder_ctx(self, library_name, folder_name=''):
        """
        Set the current SharePoint library and folder context for file operations.

        :param library_name: str
            The name of the SharePoint library to work with. For example, "Shared Documents."

        :param folder_name: str, optional
            The name of the SharePoint subfolder within the library. This can be an empty string
            to indicate the root of the library. (Default: Empty string)

        :raises Exception:
            If an error occurs while setting the folder context, an exception will be raised
            with details about the error.
        """
        try:
            self.library_name = library_name
            self.folder_name = folder_name
            self.folder = self.ctx.web.get_folder_by_server_relative_url(f'{self.relative_url}/{library_name}/{folder_name}')
        except Exception as e:
            print(f"Error setting folder context: {str(e)}")
             

    def get_files(self):
        """
        Retrieve a list of SharePoint file objects within the current library and folder context.

        :return: list of SharePoint file objects
            A list containing SharePoint file objects within the current context.

        :raises Exception:
            If an error occurs while retrieving files, an exception will be raised
            with details about the error.
        """
        try:
            files = self.folder.files
            self.ctx.load(files)
            self.ctx.execute_query()
            return [file for file in files]
        except Exception as e:
            print(f"Error getting files: {str(e)}")


    def get_folders(self):
        """
        Retrieve a list of SharePoint folder objects within the current library and folder context.

        :return: list of SharePoint folder objects
            A list containing SharePoint folder objects within the current context.

        :raises Exception:
            If an error occurs while retrieving folders, an exception will be raised
            with details about the error.
        """
        try:
            folders = self.folder.folders
            self.ctx.load(folders)
            self.ctx.execute_query()
            return [folder for folder in folders]
        except Exception as e:
            print(f"Error getting folders: {str(e)}")

    def _get_existing_files(self):
        """
        Retrieve a set of existing SharePoint file names within the current library and folder context.

        :return: set of file names
            A set containing the names of existing SharePoint files within the current context.

        :raises Exception:
            If an error occurs while retrieving existing files, an exception will be raised
            with details about the error.
        """
        try:
            files = self.get_files()
            self.existing_files = set(map(lambda file: file['Name'], files))
        except Exception as e:
            print(f"Error getting existing files: {str(e)}")    

    def _reset_existing_files(self):
        """
        Clear the set of existing SharePoint file names within the current library and folder context.
        """
        self.existing_files.clear()


    def create_sharepoint_folder(self, parent_folder_url):
        """
        This method allows you to create a new folder within the current library and folder context.
        
        :param parent_folder_url: The relative URL of the new folder within the current context.
        :type parent_folder_url: str

        :raises:
            Exception: If an error occurs during folder creation.
        """
        try:
            self.folder.add(parent_folder_url)
            self.ctx.execute_query()
            print(f"Created SharePoint folder: {parent_folder_url}")
        except Exception as e:
            print(f"Error creating SharePoint folder: {str(e)}") 

    def get_file_name(self, file_name):
        """
        Get a unique file name within the current SharePoint folder.

        This method ensures that the given file name is unique within the current folder context. 
        If a file with the same name already exists, a unique name is generated by appending an incrementing suffix.

        :param file_name: The original file name.
        :type file_name: str

        :return: A unique file name.
        :rtype: str

        :raises:
            Exception: If an error occurs while generating a unique file name.
        """
        try:
            base_name, ext = os.path.splitext(file_name)
            unique_suffix = 1

            while file_name in self.existing_files:
                file_name = f"{base_name}_{unique_suffix}{ext}"
                unique_suffix += 1

            return file_name
        except Exception as e:
            print(f"Error getting file name: {str(e)}")

             

    def upload_file(self, local_path):
        """
        Upload a file to the current SharePoint folder, overwriting an existing file with the same name if it exists.

        This method allows you to upload a file to the current SharePoint folder. If a file with the same name already exists
        in the folder, it will be overwritten by the new file. The method explicitly overwrites the existing file, ensuring
        that the new file takes precedence.

        :param local_path: The local file path to upload, replacing any existing file with the same name.
        :type local_path: str

        :raises:
            Exception: If an error occurs during the upload process.
        """
        try:
            file_name = os.path.basename(local_path)

            with open(local_path, 'rb') as content_file:
                file_content = content_file.read()
                self.folder.upload_file(file_name=file_name, content=file_content)

            self.ctx.execute_query()
            self.existing_files.add(file_name)
            print(f"Uploaded file with override: {file_name}")
        except Exception as e:
            print(f"Error uploading file with override: {str(e)}")

    
    def upload_file_without_override(self, local_path, file_name=None):
        """
        Upload a file to the current SharePoint folder, avoiding overwriting an existing file with the same name.

        This method allows you to upload a file to the current SharePoint folder without overwriting an existing file with
        the same name. If a file with the specified name already exists in the folder, a new filename will be generated
        to avoid conflicts. The method ensures that no existing files are overwritten.

        :param local_path: The local file path to upload.
        :type local_path: str

        :param file_name: (Optional) The desired name for the file in SharePoint. If not provided, the original filename will be used.
        :type file_name: str

        :raises:
            Exception: If an error occurs during the upload process.
        """
        try:
            if not self.existing_files:
                self._get_existing_files()  # Load existing files if the set is empty

            if file_name is None: 
                file_name = os.path.basename(local_path)
            
            file_name = self.get_file_name(file_name)

            with open(local_path, 'rb') as content_file:
                file_content = content_file.read()
                self.folder.upload_file(file_name=file_name, content=file_content)

            self.ctx.execute_query()
            self.existing_files.add(file_name)
            print(f"Uploaded file without override: {file_name}")
        except Exception as e:
            print(f"Error uploading file without override: {str(e)}")

       
    def upload_multiple_files(self, local_file_paths):
        """
        Upload multiple files to the current SharePoint folder.

        This method allows you to upload multiple files to the current SharePoint folder. It iterates through a list of local
        file paths, uploads each file. If files with the same names already exist in the folder, the new files will overwrite
        the existing ones.

        :param local_file_paths: A list of local file paths to upload.
        :type local_file_paths: list of str

        :raises:
            Exception: If an error occurs during the upload process.
        """
        try:
            self._get_existing_files()
            for local_path in local_file_paths:
                self.upload_file(local_path)
            self._reset_existing_files()
        except Exception as e:
            print(f"Error uploading multiple files: {str(e)}")

   
    def upload_multiple_files_without_override(self, local_file_paths):
        """
        Upload multiple files to the current SharePoint folder without overwriting.

        This method allows you to upload multiple files to the current SharePoint folder without overwriting existing files.
        It iterates through a list of local file paths, uploads each file. If files with the same names already exist in the
        folder, the new files will be assigned unique names to avoid overwriting.

        :param local_file_paths: A list of local file paths to upload.
        :type local_file_paths: list of str

        :raises:
            Exception: If an error occurs during the upload process.
        """
        try:
            self._get_existing_files()
            for local_path in local_file_paths:
                self.upload_file_without_override(local_path)
            self._reset_existing_files()
        except Exception as e:
            print(f"Error uploading multiple files without override: {str(e)}")


    def delete_file(self, file_name):
        """
        This method allows you to delete a specific file by its name from the current SharePoint folder context.

        :param file_name: The name of the file to be deleted.
        :type file_name: str

        :raises:
            Exception: If an error occurs during the file deletion process.
        """
        try:
            file = self.folder.files.get_by_url(file_name)
            file.delete_object()
            self.ctx.execute_query()
            print(f"Deleted file: {file_name}")
        except Exception as e:
            print(f"Error deleting file: {str(e)}")

    
    def delete_multiple_files(self, file_names):
        """
        This method allows you to delete multiple files by providing a list of file names to be deleted from the current SharePoint folder.

        :param file_names: A list of file names to be deleted.
        :type file_names: List[str]
        """
        for file_name in file_names:
            self.delete_file(file_name)

    
    def delete_all_files_and_folders(self):
        """
        Delete all files and folders from the current SharePoint folder.

        This method allows you to delete all files and folders from the current SharePoint folder. It iterates through the files and folders in the current folder and deletes them.

        :raises:
            Exception: If an error occurs during the deletion process.
        """
        try:
            for file in self.get_files():
                file.delete_object()

            for folder in self.get_folders():
                folder.delete_object()

            print("All files and folders are deleted")
        except Exception as e:
            print(f"Error deleting all files and folders: {str(e)}")

    
    def delete_entire_folder(self):
        """
        Delete the entire SharePoint folder.

        This method allows you to delete the entire SharePoint folder represented by the current context. 

        :raises:
            Exception: If an error occurs during the deletion process.
        """
        try:
            self.folder.delete_object()
            self.ctx.execute_query()
            print("SharePoint folder and its contents are deleted")
        except Exception as e:
            print(f"Error deleting the SharePoint folder: {str(e)}")

    def get_recently_deleted_items(self, max_items=5):
        """
        Retrieve recently deleted items from the SharePoint recycle bin.

        This method allows you to retrieve recently deleted items from the SharePoint recycle bin. You can specify the maximum number of items to retrieve (default is 5).

        :param max_items: The maximum number of recently deleted items to retrieve. Default is 10.
        :type max_items: int

        :return: A list of SharePoint items recently deleted from the recycle bin.
        :rtype: list

        :raises:
            Exception: If an error occurs during the retrieval process.
        """
        try:
            recycle_bin = self.ctx.site.recycle_bin
            self.ctx.load(recycle_bin)
            self.ctx.execute_query()

            # Get all items from the recycle bin
            items = recycle_bin.get()
            self.ctx.load(items)
            self.ctx.execute_query()

            # Sort the items by DeletedDate in descending order
            items = sorted(items, key=lambda item: item.deleted_date, reverse=True)

            # Take the latest 'max_items' items or all if 'max_items' is not specified
            items = items[:max_items] if max_items is not None else items

            return items
        except Exception as e:
            print(f"Error getting recently deleted items: {str(e)}")

    def recover_data(self, num_to_restore=1):
        """
        Recover recently deleted items from the SharePoint recycle bin.

        This method allows you to recover a specified number of recently deleted items from the SharePoint recycle bin.
        It reuses the `get_recently_deleted_items` method to retrieve the recently deleted items and restores them (default is 1).
        After the recovery process, it prints information about the items restored.

        :param num_to_restore: The number of recently deleted items to recover.
        :type num_to_restore: int

        :return: None

        :raises:
            Exception: If an error occurs during the recovery process.
        """
        try:
            # Reuse the get_recently_deleted_items method to retrieve recently deleted items
            recently_deleted_items = self.get_recently_deleted_items(num_to_restore)

            # Restore the specified number of files
            restored_count = 0
            for item in recently_deleted_items:
                if restored_count < num_to_restore:
                    try:
                        # Restore the file
                        item.restore()
                    except:
                        # Restore the folder
                        self.restore_folder(item)
                    self.ctx.execute_query()
                    print(f"Item recovered to {item.properties['DirName']} file name is {item.properties['Title']}")
                    restored_count += 1
            print(f"Restored {restored_count} items")
        except Exception as e:
            print(f"Error restoring files: {str(e)}")