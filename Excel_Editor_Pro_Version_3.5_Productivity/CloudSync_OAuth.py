"""
Enhanced Cloud Synchronization Module with Real OAuth Authentication
This replaces the authenticate_oauth method with actual browser-based OAuth
"""

import webbrowser
import http.server
import socketserver
import urllib.parse
import threading
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import QTimer

# OAuth Configuration - YOU NEED TO REPLACE THESE WITH YOUR ACTUAL CREDENTIALS
OAUTH_CONFIGS = {
    'Google Drive': {
        'client_id': 'YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com',
        'client_secret': 'YOUR_GOOGLE_CLIENT_SECRET',
        'auth_url': 'https://accounts.google.com/o/oauth2/v2/auth',
        'token_url': 'https://oauth2.googleapis.com/token',
        'redirect_uri': 'http://localhost:8080/callback',
        'scope': 'https://www.googleapis.com/auth/drive.file',
    },
    'OneDrive': {
        'client_id': 'YOUR_ONEDRIVE_CLIENT_ID',
        'client_secret': 'YOUR_ONEDRIVE_CLIENT_SECRET',
        'auth_url': 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
        'token_url': 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        'redirect_uri': 'http://localhost:8080/callback',
        'scope': 'Files.ReadWrite offline_access',
    },
    'Dropbox': {
        'client_id': 'YOUR_DROPBOX_APP_KEY',
        'client_secret': 'YOUR_DROPBOX_APP_SECRET',
        'auth_url': 'https://www.dropbox.com/oauth2/authorize',
        'token_url': 'https://api.dropboxapi.com/oauth2/token',
        'redirect_uri': 'http://localhost:8080/callback',
        'scope': 'files.content.write files.content.read',
    }
}


class OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    """HTTP handler to receive OAuth callback"""
    
    auth_code = None
    
    def do_GET(self):
        """Handle the OAuth callback"""
        # Parse the callback URL
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        
        if 'code' in params:
            # Store the authorization code
            OAuthCallbackHandler.auth_code = params['code'][0]
            
            # Send success response
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            
            success_html = """
            <!DOCTYPE html>
            <html>
            <head>
                <title>Authentication Successful</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 100vh;
                        margin: 0;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    }
                    .container {
                        background: white;
                        padding: 40px;
                        border-radius: 10px;
                        box-shadow: 0 10px 40px rgba(0,0,0,0.2);
                        text-align: center;
                    }
                    h1 {
                        color: #4CAF50;
                        margin-bottom: 20px;
                    }
                    p {
                        color: #555;
                        font-size: 18px;
                    }
                    .checkmark {
                        font-size: 72px;
                        color: #4CAF50;
                        margin-bottom: 20px;
                    }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="checkmark">✓</div>
                    <h1>Authentication Successful!</h1>
                    <p>You can now close this window and return to Excel Editor Pro.</p>
                </div>
            </body>
            </html>
            """
            self.wfile.write(success_html.encode())
        elif 'error' in params:
            # Handle error
            self.send_response(400)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            
            error_html = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Authentication Failed</title>
                <style>
                    body {{
                        font-family: Arial, sans-serif;
                        display: flex;
                        justify-content: center;
                        align-items: center;
                        height: 100vh;
                        margin: 0;
                        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                    }}
                    .container {{
                        background: white;
                        padding: 40px;
                        border-radius: 10px;
                        box-shadow: 0 10px 40px rgba(0,0,0,0.2);
                        text-align: center;
                    }}
                    h1 {{
                        color: #f5576c;
                        margin-bottom: 20px;
                    }}
                    p {{
                        color: #555;
                        font-size: 18px;
                    }}
                    .error-icon {{
                        font-size: 72px;
                        color: #f5576c;
                        margin-bottom: 20px;
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="error-icon">✗</div>
                    <h1>Authentication Failed</h1>
                    <p>Error: {params['error'][0]}</p>
                    <p>Please close this window and try again.</p>
                </div>
            </body>
            </html>
            """
            self.wfile.write(error_html.encode())
    
    def log_message(self, format, *args):
        """Suppress logging"""
        pass


def start_oauth_server(port=8080, timeout=300):
    """Start a local server to handle OAuth callback"""
    OAuthCallbackHandler.auth_code = None
    
    server = socketserver.TCPServer(("", port), OAuthCallbackHandler)
    server.timeout = timeout
    
    # Run server in a separate thread
    server_thread = threading.Thread(target=server.handle_request)
    server_thread.daemon = True
    server_thread.start()
    
    return server


def authenticate_oauth_real(service_name, parent_widget=None):
    """
    Real OAuth authentication that opens browser and handles callback
    
    Args:
        service_name: Name of the cloud service (Google Drive, OneDrive, Dropbox)
        parent_widget: Parent widget for displaying messages
    
    Returns:
        tuple: (success: bool, credentials: dict or None)
    """
    
    if service_name not in OAUTH_CONFIGS:
        if parent_widget:
            QMessageBox.warning(parent_widget, "Error", 
                f"No OAuth configuration found for {service_name}")
        return False, None
    
    config = OAUTH_CONFIGS[service_name]
    
    # Check if credentials are configured
    if config['client_id'].startswith('YOUR_'):
        if parent_widget:
            QMessageBox.warning(parent_widget, "Configuration Required", 
                f"OAuth credentials not configured for {service_name}.\n\n"
                "To enable OAuth authentication:\n\n"
                "1. Get OAuth credentials from the cloud provider's developer console:\n"
                "   • Google Drive: https://console.cloud.google.com/\n"
                "   • OneDrive: https://portal.azure.com/\n"
                "   • Dropbox: https://www.dropbox.com/developers/apps\n\n"
                "2. Update the OAuth configuration in CloudSync_OAuth.py\n\n"
                "3. Set the redirect URI to: http://localhost:8080/callback\n\n"
                "For now, you can use manual credential authentication instead.")
        return False, None
    
    try:
        # Start local server to receive callback
        print(f"Starting OAuth server for {service_name}...")
        server = start_oauth_server(8080)
        
        # Build authorization URL
        auth_params = {
            'client_id': config['client_id'],
            'redirect_uri': config['redirect_uri'],
            'response_type': 'code',
            'scope': config['scope'],
        }
        
        # Add service-specific parameters
        if service_name == 'Google Drive':
            auth_params['access_type'] = 'offline'
            auth_params['prompt'] = 'consent'
        
        auth_url = f"{config['auth_url']}?{urllib.parse.urlencode(auth_params)}"
        
        # Open browser for authentication
        print(f"Opening browser for {service_name} authentication...")
        if parent_widget:
            QMessageBox.information(parent_widget, "Browser Opening", 
                f"Your browser will open for {service_name} authentication.\n\n"
                "Please sign in and grant permissions.\n\n"
                "The application will continue automatically after authentication.")
        
        webbrowser.open(auth_url)
        
        # Wait for callback (with timeout)
        print("Waiting for authentication callback...")
        timeout_counter = 0
        max_timeout = 300  # 5 minutes
        
        while OAuthCallbackHandler.auth_code is None and timeout_counter < max_timeout:
            if parent_widget:
                # Process Qt events to keep UI responsive
                from PyQt5.QtWidgets import QApplication
                QApplication.processEvents()
            
            import time
            time.sleep(0.5)
            timeout_counter += 0.5
        
        if OAuthCallbackHandler.auth_code is None:
            if parent_widget:
                QMessageBox.warning(parent_widget, "Timeout", 
                    "Authentication timed out. Please try again.")
            return False, None
        
        # Exchange authorization code for access token
        print("Exchanging authorization code for access token...")
        import requests
        
        token_data = {
            'client_id': config['client_id'],
            'client_secret': config['client_secret'],
            'code': OAuthCallbackHandler.auth_code,
            'redirect_uri': config['redirect_uri'],
            'grant_type': 'authorization_code',
        }
        
        response = requests.post(config['token_url'], data=token_data)
        
        if response.status_code == 200:
            credentials = response.json()
            print("Authentication successful!")
            
            if parent_widget:
                QMessageBox.information(parent_widget, "Success", 
                    f"Successfully authenticated with {service_name}!")
            
            return True, credentials
        else:
            print(f"Token exchange failed: {response.text}")
            if parent_widget:
                QMessageBox.warning(parent_widget, "Authentication Failed", 
                    f"Failed to obtain access token: {response.text}")
            return False, None
            
    except Exception as e:
        print(f"OAuth authentication error: {e}")
        if parent_widget:
            QMessageBox.critical(parent_widget, "Error", 
                f"Authentication error: {str(e)}\n\n"
                "Please check your internet connection and try again.")
        return False, None


# INSTRUCTIONS TO INTEGRATE THIS INTO YOUR EXISTING CODE:
# ========================================================
# 
# In your CloudSync_.py file, find the authenticate_oauth method
# (around line 185) and replace it with:
#
# def authenticate_oauth(self):
#     """Authenticate using OAuth 2.0"""
#     from CloudSync_OAuth import authenticate_oauth_real
#     
#     success, credentials = authenticate_oauth_real(self.service_name, self)
#     
#     if success and credentials:
#         self.authenticated = True
#         self.credentials = credentials
#         self.accept()
#     else:
#         self.authenticated = False
