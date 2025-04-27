import os
import json
import requests
from functools import wraps
from flask import redirect, request, session, url_for

class GoogleAuth:
    GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
    GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token"
    GOOGLE_USERINFO_URL = "https://www.googleapis.com/oauth2/v1/userinfo"

    def __init__(self, app=None):
        self.client_id = os.environ.get('GOOGLE_OAUTH_CLIENT_ID')
        self.client_secret = os.environ.get('GOOGLE_OAUTH_CLIENT_SECRET')
        if app:
            self.init_app(app)

    def init_app(self, app):
        self.app = app
        # Ensure secret key is set for session management
        if not app.secret_key:
            app.secret_key = os.urandom(24)

    def get_auth_url(self):
        """Generate Google OAuth2 authorization URL"""
        params = {
            'client_id': self.client_id,
            'redirect_uri': self._get_redirect_uri(),
            'scope': 'email profile',
            'response_type': 'code',
            'access_type': 'offline',
            'prompt': 'consent'
        }
        
        # Convert params to URL query string
        query_string = '&'.join([f"{key}={value}" for key, value in params.items()])
        return f"{self.GOOGLE_AUTH_URL}?{query_string}"

    def get_token(self, code):
        """Exchange authorization code for tokens"""
        data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'code': code,
            'redirect_uri': self._get_redirect_uri(),
            'grant_type': 'authorization_code'
        }
        response = requests.post(self.GOOGLE_TOKEN_URL, data=data)
        return response.json()

    def get_user_info(self, access_token):
        """Get user info using access token"""
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get(self.GOOGLE_USERINFO_URL, headers=headers)
        return response.json()

    def _get_redirect_uri(self):
        """Get the appropriate redirect URI based on environment"""
        # Check if we're running on Render by looking at the request host
        is_production = 'onrender.com' in request.host
        
        if is_production:
            # Production URL on Render
            return "https://decklyst.onrender.com/login/google/authorized"
        else:
            # Local development URL
            return request.url_root.rstrip('/') + "/login/google/authorized"

def login_required(f):
    """Decorator to require login for routes"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function
