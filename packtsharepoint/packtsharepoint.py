from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File

class packt_sharepoint:
    '''
    A class used to connect to sharepoint.
    
    Attributes
    ----------
    client_id: str
        the client id used for authentication
    client_secret: str
        the client secret used for authentication
    site_url: str
        the url of the sharepoint site
    
    Methods
    --------
    connect
        Connects to sharepoint returns client
    get_file
        Returns a file from a sharepoint url        
    '''
    
    def __init__(self, client_id, client_secret, site_url):
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_url = site_url
        self.url = 'https://packtservices.sharepoint.com/'
        self.ctx = None
        self.connect()

    def connect(self):
        self.ctx = AuthenticationContext(self.url)
        self.ctx.acquire_token_for_app(self.client_id, self.client_secret)
        self.ctx = ClientContext(self.site_url, self.ctx)
        return self.ctx

    def get_file(self, file_url):
        file = File.open_binary(self.ctx, file_url)
        self.ctx.execute_query()
        return file