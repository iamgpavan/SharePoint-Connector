from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

class SharePointConnection:
    """
    SharePointConnection manages the connection to a SharePoint site using client or user credentials.

    This class provides a singleton instance for connecting to a SharePoint site based on the provided authentication method.
    You can connect using either user credentials (username and password) or client credentials (client ID and client secret).

    Author:
    Pavan Kumar Gattupalli
    LinkedIn: https://www.linkedin.com/in/g-pavan-kumar/
    GitHub: https://github.com/g-pavan

    Example Usage:
    connection = SharePointConnection("your_site_url", username="your_username", password="your_password")
    ctx = connection.ctx  # Get the SharePoint ClientContext for API interactions

    Parameters:
    - site_url (str): The URL of the SharePoint site.
    - username (str, optional): The username for user credentials. (Default: None)
    - password (str, optional): The password for user credentials. (Default: None)
    - client_id (str, optional): The client ID for client credentials. (Default: None)
    - client_secret (str, optional): The client secret for client credentials. (Default: None)

    Attributes:
    - ctx (ClientContext): The SharePoint ClientContext instance for making API requests.

    Note:
    - This class implements the Singleton design pattern to ensure there is only one active connection instance.
    """

    _instance = None

    def __new__(cls, site_url, username=None, password=None, client_id=None, client_secret=None):
        if cls._instance is None:
            cls._instance = super(SharePointConnection, cls).__new__(cls)
            cls._instance.site_url = site_url
            cls._instance.username = username
            cls._instance.password = password
            cls._instance.client_id = client_id
            cls._instance.client_secret = client_secret
            cls._instance.ctx = cls._instance._connect_to_sharepoint()
        return cls._instance

    def _connect_to_sharepoint(self):
        """
        Create a SharePoint ClientContext based on the provided authentication method.

        Returns:
        - ctx (ClientContext): The SharePoint ClientContext instance.

        Raises:
        - ValueError: If neither username and password nor client ID and secret are provided.
        """
        if self.username and self.password:
            user_credentials = UserCredential(self.username, self.password)
            ctx = ClientContext(self.site_url).with_credentials(user_credentials)
        elif self.client_id and self.client_secret:
            client_credentials = ClientCredential(self.client_id, self.client_secret)
            ctx = ClientContext(self.site_url).with_credentials(client_credentials)
        else:
            raise ValueError("Provide either username and password or client ID and secret.")
        return ctx