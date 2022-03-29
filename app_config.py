CLIENT_ID = "8a3cbd3a-b215-454d-868c-90cd977165a6"  # Application (client) ID of app registration

CLIENT_SECRET = "gkg7Q~NexPhfUwgTnWI1IJCd.tbndrKyGFuPD"
# In a production app, we recommend you use a more secure method of storing your secret,
# like Azure Key Vault. Or, use an environment variable as described in Flask's documentation:
# https://flask.palletsprojects.com/en/1.1.x/config/#configuring-from-environment-variables
# CLIENT_SECRET = os.getenv("CLIENT_SECRET")
# if not CLIENT_SECRET:
#     raise ValueError("Need to define CLIENT_SECRET environment variable")

TENANT_ID = "fb0f750f-9428-4166-a99d-8fc795c1a99e"

AUTHORITY = "https://login.microsoftonline.com/" + TENANT_ID  # For multi-tenant app
# AUTHORITY = "https://login.microsoftonline.com/Enter_the_Tenant_Name_Here"

REDIRECT_PATH = "/getAToken"  # Used for forming an absolute URL to your redirect URI.
# The absolute URL must match the redirect URI you set
# in the app's registration in the Azure portal.

GRAPH_API_CALLS = {"get_all_sites": "https://graph.microsoft.com/v1.0/sites?search=*",
                   "get_subsites": "https://graph.microsoft.com/v1.0/sites/{site_id}/sites",
                   "get_lists": "https://graph.microsoft.com/v1.0/sites/{site_id}/lists",
                   "get_list_items": "https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items",
                   "get_drive_item": "https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items/{item_id}/driveItem"
                   }

# You can find more Microsoft Graph API endpoints from Graph Explorer
# https://developer.microsoft.com/en-us/graph/graph-explorer
ENDPOINT = 'https://graph.microsoft.com/v1.0/users'  # This resource requires no admin consent
ENDPOINT = "https://graph.microsoft.com/v1.0/me/drive/recent"

# You can find the proper permission names from this document
# https://docs.microsoft.com/en-us/graph/permissions-reference
SCOPE = ['https://graph.microsoft.com/.default']

SESSION_TYPE = "filesystem"  # Specifies the token cache should be stored in server-side session
SHAREPOINT_FILE_DIR = 'SharePoint_files'
SHAREPOINT_CONTENT_JSON_FILE = "sharepoint_content.json"
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'
TIKA_SERVICE_URL = "http://localhost:5007"
