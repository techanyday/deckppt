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
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
    'openid'
]

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
        try:
            flow = Flow.from_client_config(
                self.client_secrets,
                scopes=SCOPES,
                state=state
            )
            flow.redirect_uri = url_for('oauth2callback', _external=True)
            
            # Get credentials from the callback code
            flow.fetch_token(code=code)
            credentials = flow.credentials
            
            logger.info("Successfully obtained credentials from OAuth flow")
            return credentials
            
        except Exception as e:
            logger.error(f"Error getting credentials from code: {str(e)}")
            raise
            
    def init_service(self, credentials_dict):
        """Initialize the Slides service with credentials."""
        try:
            # Convert dictionary back to Credentials object
            credentials = Credentials(
                token=credentials_dict['token'],
                refresh_token=credentials_dict['refresh_token'],
                token_uri=credentials_dict['token_uri'],
                client_id=credentials_dict['client_id'],
                client_secret=credentials_dict['client_secret'],
                scopes=credentials_dict['scopes']
            )
            
            # Build the service
            self.service = build('slides', 'v1', credentials=credentials)
            logger.info("Successfully initialized Slides service")
            return self.service
            
        except Exception as e:
            logger.error(f"Error initializing service: {str(e)}")
            raise
            
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
        
    def _generate_content_sections(self, topic, num_sections):
        """Generate content sections using OpenAI."""
        try:
            from openai_client import OpenAIClient
            client = OpenAIClient()
            
            prompt = f"""Create {num_sections} detailed sections for a presentation about "{topic}".
            For each section:
            1. Create a clear, concise title (2-4 words)
            2. Provide EXACTLY 5 bullet points that are:
               - Concise (one short phrase or sentence)
               - Focused on key insights
               - Clear and informative
               - Not overly wordy
            
            Format your response exactly like this example (including the quotes):
            [
                {{
                    "title": "Market Overview",
                    "points": [
                        "Global market reached $500B in 2024",
                        "35% annual growth rate expected",
                        "North America leads adoption at 40%",
                        "Five major players control 60%",
                        "Emerging markets show 50% potential"
                    ]
                }}
            ]

            Requirements:
            - Each section MUST have exactly 5 bullet points
            - Each bullet point should be 5-12 words maximum
            - Keep content professional and data-driven
            - Use proper JSON formatting with double quotes
            - Include exactly {num_sections} sections"""
            
            response = client.generate(prompt)
            
            # Clean up the response
            response = response.strip()
            if response.startswith('```python'):
                response = response.replace('```python', '').replace('```', '').strip()
            if response.startswith('```json'):
                response = response.replace('```json', '').replace('```', '').strip()
                
            # Handle single quotes
            response = response.replace("'", '"')
            
            # Safely evaluate the response
            import json
            try:
                sections = json.loads(response)
            except json.JSONDecodeError as e:
                logger.error(f"Error parsing JSON: {str(e)}")
                logger.error(f"Response was: {response}")
                # Use default content as fallback
                sections = [
                    {
                        'title': 'Overview',
                        'points': [
                            f'Introduction to {topic} and its importance',
                            'Current market size and growth potential',
                            'Key industry players and market share',
                            'Major technological developments',
                            'Future growth projections and trends'
                        ]
                    },
                    {
                        'title': 'Key Benefits',
                        'points': [
                            'Increased operational efficiency and productivity',
                            'Significant cost reduction and ROI',
                            'Enhanced customer satisfaction and retention',
                            'Improved market competitiveness',
                            'Sustainable long-term growth potential'
                        ]
                    },
                    {
                        'title': 'Implementation',
                        'points': [
                            'Step-by-step deployment strategy',
                            'Required resources and timeline',
                            'Key success factors and metrics',
                            'Risk mitigation approaches',
                            'Best practices and guidelines'
                        ]
                    }
                ][:num_sections]
            
            return sections
            
        except Exception as e:
            logger.error(f"Error generating content: {str(e)}")
            return [
                {
                    'title': 'Overview',
                    'points': [
                        f'Introduction to {topic} and its importance',
                        'Current market size and growth potential',
                        'Key industry players and market share',
                        'Major technological developments',
                        'Future growth projections and trends'
                    ]
                },
                {
                    'title': 'Key Benefits',
                    'points': [
                        'Increased operational efficiency and productivity',
                        'Significant cost reduction and ROI',
                        'Enhanced customer satisfaction and retention',
                        'Improved market competitiveness',
                        'Sustainable long-term growth potential'
                    ]
                },
                {
                    'title': 'Implementation',
                    'points': [
                        'Step-by-step deployment strategy',
                        'Required resources and timeline',
                        'Key success factors and metrics',
                        'Risk mitigation approaches',
                        'Best practices and guidelines'
                    ]
                }
            ][:num_sections]
            
    def _create_content_slides(self, presentation_id, content_sections):
        """Create content slides with insights."""
        try:
            # Color schemes for rotation
            color_schemes = [
                {
                    'background': {'red': 0.95, 'green': 0.97, 'blue': 1.0},  # Light blue
                    'title': {'red': 0.0, 'green': 0.4, 'blue': 0.8},  # Deep blue
                    'text': {'red': 0.2, 'green': 0.2, 'blue': 0.2}  # Dark gray
                },
                {
                    'background': {'red': 0.96, 'green': 1.0, 'blue': 0.96},  # Light green
                    'title': {'red': 0.0, 'green': 0.6, 'blue': 0.4},  # Teal
                    'text': {'red': 0.2, 'green': 0.2, 'blue': 0.2}  # Dark gray
                },
                {
                    'background': {'red': 0.98, 'green': 0.95, 'blue': 1.0},  # Soft lavender
                    'title': {'red': 0.5, 'green': 0.2, 'blue': 0.8},  # Purple
                    'text': {'red': 0.2, 'green': 0.2, 'blue': 0.2}  # Dark gray
                }
            ]
            
            for i, section in enumerate(content_sections):
                # Get color scheme for this slide
                color_scheme = color_schemes[i % len(color_schemes)]
                
                # Create a new slide
                slide = {
                    'createSlide': {
                        'objectId': f'slide_{i+1}',
                        'slideLayoutReference': {
                            'predefinedLayout': 'BLANK'
                        }
                    }
                }
                
                self.service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': [slide]}
                ).execute()
                
                slide_id = f'slide_{i+1}'
                title = section.get('title', '')
                points = section.get('points', [])
                
                # Create requests for slide elements
                requests = [
                    # Add background shape
                    {
                        'createShape': {
                            'objectId': f'background_{i+1}',
                            'shapeType': 'RECTANGLE',
                            'elementProperties': {
                                'pageObjectId': slide_id,
                                'size': {
                                    'width': {'magnitude': 720, 'unit': 'PT'},
                                    'height': {'magnitude': 405, 'unit': 'PT'}
                                },
                                'transform': {
                                    'scaleX': 1,
                                    'scaleY': 1,
                                    'translateX': 0,
                                    'translateY': 0,
                                    'unit': 'PT'
                                }
                            }
                        }
                    },
                    # Style background
                    {
                        'updateShapeProperties': {
                            'objectId': f'background_{i+1}',
                            'shapeProperties': {
                                'solidFill': {'color': color_scheme['background']}
                            },
                            'fields': 'solidFill'
                        }
                    },
                    # Add title box
                    {
                        'createShape': {
                            'objectId': f'title_{i+1}',
                            'shapeType': 'RECTANGLE',
                            'elementProperties': {
                                'pageObjectId': slide_id,
                                'size': {
                                    'width': {'magnitude': 650, 'unit': 'PT'},
                                    'height': {'magnitude': 50, 'unit': 'PT'}
                                },
                                'transform': {
                                    'scaleX': 1,
                                    'scaleY': 1,
                                    'translateX': 35,
                                    'translateY': 20,
                                    'unit': 'PT'
                                }
                            }
                        }
                    },
                    # Add title text
                    {
                        'insertText': {
                            'objectId': f'title_{i+1}',
                            'text': title
                        }
                    },
                    # Style title
                    {
                        'updateTextStyle': {
                            'objectId': f'title_{i+1}',
                            'style': {
                                'fontFamily': 'Roboto',
                                'fontSize': {'magnitude': 32, 'unit': 'PT'},
                                'foregroundColor': {'opaqueColor': {'rgbColor': color_scheme['title']}},
                                'bold': True
                            },
                            'fields': 'fontFamily,fontSize,foregroundColor,bold'
                        }
                    },
                    # Add content box
                    {
                        'createShape': {
                            'objectId': f'content_{i+1}',
                            'shapeType': 'RECTANGLE',
                            'elementProperties': {
                                'pageObjectId': slide_id,
                                'size': {
                                    'width': {'magnitude': 650, 'unit': 'PT'},
                                    'height': {'magnitude': 280, 'unit': 'PT'}
                                },
                                'transform': {
                                    'scaleX': 1,
                                    'scaleY': 1,
                                    'translateX': 35,
                                    'translateY': 90,
                                    'unit': 'PT'
                                }
                            }
                        }
                    }
                ]
                
                # Add bullet points with proper spacing
                bullet_text = '\n'.join(f'• {point}' for point in points)
                requests.extend([
                    # Add bullet points
                    {
                        'insertText': {
                            'objectId': f'content_{i+1}',
                            'text': bullet_text
                        }
                    },
                    # Style bullet points
                    {
                        'updateTextStyle': {
                            'objectId': f'content_{i+1}',
                            'style': {
                                'fontFamily': 'Roboto',
                                'fontSize': {'magnitude': 20, 'unit': 'PT'},
                                'foregroundColor': {'opaqueColor': {'rgbColor': color_scheme['text']}},
                                'spacingMode': 'SPACING_MODE_FIXED',
                                'lineSpacing': 150  # 1.5 line spacing
                            },
                            'fields': 'fontFamily,fontSize,foregroundColor,spacingMode,lineSpacing'
                        }
                    },
                    # Add paragraph spacing
                    {
                        'updateParagraphStyle': {
                            'objectId': f'content_{i+1}',
                            'style': {
                                'lineSpacing': 150,
                                'spaceAbove': {'magnitude': 10, 'unit': 'PT'},
                                'spaceBelow': {'magnitude': 10, 'unit': 'PT'}
                            },
                            'fields': 'lineSpacing,spaceAbove,spaceBelow'
                        }
                    }
                ])
                
                # Execute the requests
                self.service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
                
            logger.info(f"Created {len(content_sections)} content slides")
            
        except Exception as e:
            logger.error(f"Error creating content slides: {str(e)}")
            raise
