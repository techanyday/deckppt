import os
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import logging
from flask import url_for, session

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/presentations']

class GoogleSlidesGenerator:
    def __init__(self, client_secrets_file=None):
        self.service = None
        
        # For production (Render), load from environment variable
        if os.environ.get('GOOGLE_SLIDES_CREDENTIALS'):
            self.client_secrets = json.loads(os.environ.get('GOOGLE_SLIDES_CREDENTIALS'))
        # For local development, load from file
        elif client_secrets_file and os.path.exists(client_secrets_file):
            with open(client_secrets_file, 'r') as f:
                self.client_secrets = json.load(f)
        else:
            raise ValueError("No credentials found. Please provide either a client_secrets_file or set GOOGLE_SLIDES_CREDENTIALS environment variable.")
        
    def get_authorization_url(self, state=None):
        """Get the authorization URL for OAuth2 flow."""
        flow = Flow.from_client_config(
            self.client_secrets,
            scopes=SCOPES,
            state=state
        )
        # The URI must match one of the authorized redirect URIs
        # configured in the OAuth2 client settings
        flow.redirect_uri = url_for('oauth2callback', _external=True)
        
        authorization_url, state = flow.authorization_url(
            access_type='offline',
            include_granted_scopes='true'
        )
        
        return authorization_url, state
        
    def get_credentials_from_code(self, code, state=None):
        """Get credentials from OAuth2 callback code."""
        flow = Flow.from_client_config(
            self.client_secrets,
            scopes=SCOPES,
            state=state
        )
        flow.redirect_uri = url_for('oauth2callback', _external=True)
        
        # Get credentials from the callback code
        flow.fetch_token(code=code)
        credentials = flow.credentials
        
        return credentials
        
    def init_service(self, credentials):
        """Initialize the Slides service with credentials."""
        try:
            self.service = build('slides', 'v1', credentials=credentials)
            return True
        except Exception as e:
            logger.error(f'Error initializing service: {e}')
            return False
            
    def create_presentation(self, title, num_slides=5, theme="MODERN_BLUE"):
        """Create a new presentation with the specified title and theme."""
        if not self.service:
            raise ValueError("Service not initialized. Call init_service first.")
            
        try:
            # Create new presentation
            presentation = {
                'title': title
            }
            presentation = self.service.presentations().create(body=presentation).execute()
            presentation_id = presentation.get('presentationId')
            logger.info(f'Created presentation with ID: {presentation_id}')
            
            # Apply theme (modern blue gradient)
            if theme == "MODERN_BLUE":
                self._apply_modern_blue_theme(presentation_id)
            
            # Create title slide
            self._create_title_slide(presentation_id, title)
            
            # Create overview slide
            self._create_overview_slide(presentation_id, title)
            
            # Generate and create content slides
            content = self._generate_content_sections(title, (num_slides - 2) * 3)
            self._create_content_slides(presentation_id, content)
            
            return presentation_id
            
        except HttpError as error:
            logger.error(f'An error occurred: {error}')
            return None
            
    def _apply_modern_blue_theme(self, presentation_id):
        """Apply modern blue theme to the presentation."""
        requests = [{
            'updatePageProperties': {
                'objectId': 'p',
                'pageProperties': {
                    'pageBackgroundFill': {
                        'solidFill': {
                            'color': {
                                'rgbColor': {
                                    'red': 1.0,
                                    'green': 1.0,
                                    'blue': 1.0
                                }
                            }
                        }
                    }
                },
                'fields': 'pageBackgroundFill'
            }
        }]
        
        self.service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': requests}
        ).execute()
        
    def _create_title_slide(self, presentation_id, title):
        """Create the title slide with modern design."""
        requests = [
            # Add title box with gradient background
            {
                'createShape': {
                    'objectId': 'titleBox',
                    'shapeType': 'RECTANGLE',
                    'elementProperties': {
                        'pageObjectId': 'p',
                        'size': {
                            'width': {'magnitude': 720, 'unit': 'PT'},
                            'height': {'magnitude': 100, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 50,
                            'translateY': 200,
                            'unit': 'PT'
                        }
                    }
                }
            },
            # Add title text
            {
                'insertText': {
                    'objectId': 'titleBox',
                    'text': title
                }
            },
            # Style the title text
            {
                'updateTextStyle': {
                    'objectId': 'titleBox',
                    'style': {
                        'fontFamily': 'Roboto',
                        'fontSize': {'magnitude': 40, 'unit': 'PT'},
                        'foregroundColor': {
                            'opaqueColor': {
                                'rgbColor': {
                                    'red': 0,
                                    'green': 0.4,
                                    'blue': 1.0
                                }
                            }
                        }
                    },
                    'fields': 'fontFamily,fontSize,foregroundColor'
                }
            }
        ]
        
        self.service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': requests}
        ).execute()
        
    def _create_overview_slide(self, presentation_id, topic):
        """Create an overview slide."""
        # Implementation similar to title slide but with different layout
        pass
        
    def _create_content_slides(self, presentation_id, content_sections):
        """Create content slides with insights."""
        # Implementation for content slides with modern design
        pass
        
    def _generate_content_sections(self, topic, num_sections):
        """Generate content sections using existing OpenAI integration."""
        # Reuse existing content generation logic
        pass
