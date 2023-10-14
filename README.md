# SharePoint Manager

![GitHub](https://img.shields.io/github/license/g-pavan/SharePoint-Connector) ![GitHub stars](https://img.shields.io/github/stars/g-pavan/SharePoint-Connector) ![GitHub forks](https://img.shields.io/github/forks/g-pavan/SharePoint-Connector) ![GitHub issues](https://img.shields.io/github/issues/g-pavan/SharePoint-Connector)

Manage SharePoint files and folders with ease. Simplify file uploads, deletions, and more using this Python script.

## Table of Contents
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
  - [Authentication](#authentication)
  - [SharePoint Context (ctx)](#sharepoint-contex)
  - [Setting SharePoint Folder Context](#setting-sharepoint-folder-context)
- [Managing SharePoint Files](#managing-sharepoint-files)
  - [Listing Files](#listing-files)
  - [Uploading Files](#uploading-files)
  - [Deleting Files](#deleting-files)
- [Managing SharePoint Folders](#managing-sharepoint-folders)
  - [Listing Folders](#listing-folders)
  - [Creating Folders](#creating-folders)
  - [Deleting Folders](#deleting-folders)
- [Advanced Features](#advanced-features)
  - [Handling Recently Deleted Items](#handling-recently-deleted-items)
  - [Restoring Deleted Items](#restoring-deleted-items)
- [Conclusion](#conclusion)
- [Contributing](#contributing)
- [License](#license)
- [About the Author](#about-the-author)

## Getting Started

### Prerequisites

- Python 3.x
- [Required Python Libraries](#mention-specific-libraries)
- A SharePoint account with the necessary permissions.

### Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/YourUsername/SharePointManager.git
   ```

2. Install the required Python libraries:

    ```bash
    pip install -r requirements.txt
    ```

### Authentication

To authenticate with SharePoint, you can choose one of the following methods:

1. **User Credentials:**
   - This method is suitable when you want to use your SharePoint username and password for authentication.
   - [Create a SharePoint App Password](https://support.microsoft.com/en-us/account-billing/manage-app-passwords-for-two-step-verification-d6dc8c6d-4bf7-4851-ad95-6d07799387e9)

2. **Client Credentials:**
   - This method is suitable for application-level authentication using a client ID and client secret.
   - [Register an application with SharePoint](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs)

You can click on the links to create the necessary credentials for your SharePoint authentication.


To authenticate with SharePoint, you can choose one of the following methods:

### User Credentials

Provide your SharePoint username and password in your script:

```python
# Authenticate using user credentials
sharepoint = SharePointManager(site_url, username="your_username", password="your_password")
```
### Client Credentials
Use your client ID and client secret:

```python
# Authenticate using client credentials
sharepoint = SharePointManager(site_url, client_id="your_client_id", client_secret="your_client_secret")
```
*Please replace `your_username`, `your_password`, `your_client_id`, and `your_client_secret` with your actual SharePoint credentials*

### SharePoint Context (ctx)
The SharePoint context, often denoted as ctx, is a crucial component when working with SharePoint using this Python script. It provides the necessary connection to your SharePoint site, enabling you to perform various operations on files and folders within SharePoint libraries. The context is established using your authentication credentials and SharePoint site URL.

*When you call the SharePointManager() with relavent credentials this method will automatically sets an SharePoint Context (ctx)*

### Setting SharePoint Folder Context
Before you can work with files and folders within a specific SharePoint library or folder, you need to set the appropriate folder context. This context allows you to target the right location for your operations. The set_folder_ctx method within the script handles this task. You can change folder context any time in code flow.

```python
# Set the current SharePoint library and folder context for file operations
library_name = "Shared Documents"
folder_name = "Subfolder"
sharepoint.set_folder_ctx(library_name, folder_name)
```
In the example above, `library_name` refers to the name of the SharePoint library, and `folder_name` refers to the name of a subfolder within the library. You can choose to omit the `folder_name` or provide an empty string to work within the root of the library.

The **set_folder_ctx** method helps you focus your SharePoint operations within a specific library and folder, ensuring you work in the correct location.

## Managing SharePoint Files

### Listing Files

To list files in a SharePoint folder, you can use the following method:

```python
# Example Python code to list files in a SharePoint folder
files = sharepoint.get_files()
for file in files:
    print(file.properties["Name"])
```
### Uploading Files

To upload files to SharePoint, you can use the following methods, which offer different options for handling file uploads. Whether you want to overwrite existing files with the same name or avoid overwrites, these methods have you covered:

**upload_file(local_path)**

This method allows you to upload a file to the current SharePoint folder. If a file with the same name already exists in the folder, it will be overwritten by the new file. The method explicitly overwrites the existing file, ensuring that the new file takes precedence.

```python
local_file_path = "example.csv"
sharepoint.upload_file(local_file_path)
```
**upload_file_without_override(local_path, file_name=None)**

Use this method to upload a file to the current SharePoint folder without overwriting an existing file with the same name. If a file with the specified name already exists in the folder, a new filename will be generated to avoid conflicts. The method ensures that no existing files are overwritten.

```python
local_file_path = "example.csv"
sharepoint.upload_file_without_override(local_file_path)
```

**upload_multiple_files(local_file_paths)**

With this method, you can upload multiple files to the current SharePoint folder. It iterates through a list of local file paths, and if files with the same names already exist in the folder, the new files will overwrite the existing ones.

```python
file_paths_to_upload = ['file1.csv', 'file2.docx', 'file3.txt']
sharepoint.upload_multiple_files(file_paths_to_upload)
```

**upload_multiple_files_without_override(local_file_paths)**

Use this method to upload multiple files to the current SharePoint folder without overwriting existing files. It iterates through a list of local file paths, and if files with the same names already exist in the folder, the new files will be assigned unique names to avoid overwrites.

```python
file_paths_to_upload = ['file1.docx', 'file2.docx', 'file3.docx']
sharepoint.upload_multiple_files_without_override(file_paths_to_upload)
```

### Deleting Files

To remove files from your SharePoint library, you can utilize the following methods:

**delete_file(file_name)**

This method allows you to delete a specific file by its name from the current SharePoint folder context.

```python
file_to_delete = "example.csv"
sharepoint.delete_file(file_to_delete)
```

**delete_multiple_files(file_names)**

Use this method to delete multiple files by providing a list of file names to be removed from the current SharePoint folder.

```python
file_names = ["file1.csv", "file2.docx", "file3.txt"]
sharepoint.delete_multiple_files(file_names)
```

## Listing Folders
To retrieve a list of folders within your SharePoint library and folder context, you can use the following method:

**get_folders()**

This method returns a list of SharePoint folder objects found within the current library and folder context.

```python
folders = sharepoint.get_folders()
print("Folders in the current directory:")
for folder in folders:
    print(folder.properties["Name"])
```
### Creating Folders

To create a new folder within your SharePoint library and folder context, use the following method:

**create_sharepoint_folder(parent_folder_url)**

This method allows you to create a new folder by providing the relative URL of the folder within the current context.

```python
local_folder_name = "NewFolder"
sharepoint.create_sharepoint_folder(local_folder_name)
```
### Deleting Folders

To delete an entire folder from your SharePoint context, you can use the following method:

**delete_all_files_and_folders()**

This method allows you to delete all files and folders within the current SharePoint folder. It iterates through the files and folders in the current context and deletes them.

```python
sharepoint.set_folder_ctx(library_name, folder_name)  # Set the folder context
sharepoint.delete_all_files_and_folders()  # Delete all files and folders
```
### Delete Entire Folder

To delete the entire SharePoint folder represented by the current context, you can use the following method:

**delete_entire_folder()**

This method allows you to delete the entire SharePoint folder, including its contents, represented by the current context.

```python
folder_to_delete = "FolderToDelete"
sharepoint.set_folder_ctx(library_name, folder_to_delete)  # Set the folder context
sharepoint.delete_entire_folder()  # Delete the entire folder and its contents
```
## Advanced Features

### Handling Recently Deleted Items

To manage recently deleted items from the SharePoint recycle bin, use the following methods:

**get_recently_deleted_items(max_items)**

Retrieve recently deleted items from the SharePoint recycle bin. You can specify the maximum number of items to retrieve.

```python
recently_deleted_items = sharepoint.get_recently_deleted_items(max_items=5)
for item in recently_deleted_items:
    print(item.properties["Title"])
```
### Restoring Deleted Items

To restore deleted items, use the following method:

**recover_data(num_to_restore)** 

Recover a specified number of recently deleted items from the SharePoint recycle bin.

```python
sharepoint.recover_data(num_to_restore=3)
```
## Conclusion

In conclusion, SharePoint Manager is an invaluable tool for simplifying SharePoint file and folder management. By streamlining routine tasks and offering versatile features, this script enables you to work more efficiently and effectively within your SharePoint environment.

As you've seen, SharePoint Manager provides various functionalities, from managing connections and handling recently deleted items to creating, listing, and deleting files and folders. The possibilities for improving your SharePoint connections are endless.

I believe that technology thrives through collaboration and shared knowledge. I invite you to contribute to this project by:

- **Enhancing Functionality:** If you have ideas for new features or improvements, please share them. Your contributions can make this script even more robust.

- **Reporting Issues:** If you encounter any bugs or issues, let us know through our GitHub repository. Your feedback helps us maintain a reliable tool.

- **Documentation:** Clear and comprehensive documentation is crucial for project success. If you're passionate about technical writing, your expertise would be highly valuable in making connection management seamless.

- **Spread the Word:** If you've found SharePoint Manager helpful, share your experience with others. Word of mouth can help this project reach a broader audience.

Together, we can enhance SharePoint Manager and make it a go-to tool for managing your SharePoint connections worldwide. We encourage you to explore the project on GitHub, participate in discussions, and contribute to its growth.

## Contributing
We welcome contributions from the community to make SharePoint Manager even better. Whether you're a developer, a documentation enthusiast, or just someone with valuable feedback, your contributions are highly appreciated.

Here's how you can contribute:

**1. Feature Requests**
If you have ideas for new features, enhancements, or improvements, please open an issue on our GitHub repository. Describe the feature you'd like to see and how it would benefit the project.

**2. Reporting Issues**
If you come across any bugs, problems, or issues while using SharePoint Manager, please report them on our GitHub repository. Be sure to provide detailed information about the problem and any steps to reproduce it.

**3. Documentation**
Good documentation is essential for a successful project. If you have expertise in technical writing or simply want to help improve our documentation, your contributions are invaluable.

**4. Code Contributions**
If you're a developer and want to contribute code to the project, you can fork the repository, make your changes, and submit a pull request. We'll review your contributions and work together to integrate them into the project.

**5. Spread the Word**
Sharing your experiences with SharePoint Manager helps us reach a broader audience. Tell others about the project, write blog posts, or mention it in your professional network. Your word-of-mouth support is valuable.

**6. Collaborate**
If you have unique skills or ideas that can help advance the project, we're open to collaboration. Reach out to us on GitHub, and let's explore how we can work together to improve SharePoint Manager.

Together, we can make SharePoint Manager a more powerful and versatile tool for SharePoint connection management

## License

SharePoint Manager is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute this software as per the terms of the MIT License.

For more details, please refer to the [LICENSE](LICENSE) file.

## About the Author

**```Pavan Kumar Gattupalli```**

Connect with me on:

- [LinkedIn](https://www.linkedin.com/in/g-pavan-kumar/)
- [GitHub](https://github.com/g-pavan)

Feel free to reach out if you have any questions, suggestions, or just want to chat about SharePoint or technology in general. I'm always eager to connect with fellow enthusiasts!
