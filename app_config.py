import os

CLIENT_SECRET = "[CLIENT_SECRET]"

AUTHORITY = "https://login.microsoftonline.com/common"  # For multi-tenant app
# AUTHORITY = "https://login.microsoftonline.com/Enter_the_Tenant_Name_Here"

CLIENT_ID = "[CLIENT_ID]"

REDIRECT_URL = "[Hosted Domain]/login/authorized"  # It will be used to form an absolute URL

ENDPOINT = 'https://graph.microsoft.com/v1.0/users'

SCOPE = [
"User.Read",
"Contacts.Read",
"Mail.Read",
"Mail.Send",
"Notes.Read.All",
"Mailboxsettings.ReadWrite",
"Files.ReadWrite.All",
"User.ReadBasic.All",
"Sites.Read.All",
"Sites.ReadWrite.All"
]


SESSION_TYPE = "filesystem"  # So token cache will be stored in server-side session
SESSION_PERMANENT = False
SECRET_KEY = os.urandom(24)  # Ensure session data security