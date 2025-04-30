import os
import json
import logging
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from flask import url_for, session
import openai
from openai import OpenAI

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configure OpenAI
openai_client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

# If modifying these scopes, delete the file token.json.
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
    'openid'
]

class GoogleSlidesGenerator:
    def __init__(self):
        """Initialize the Google Slides generator."""
        self.service = None
        
    def init_service(self, credentials_dict=None):
        """Initialize the Google Slides service with credentials."""
        try:
            if not credentials_dict:
                raise ValueError("No credentials provided")
                
            # Create credentials from dictionary
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
            
        except Exception as e:
            logger.error(f"Error initializing service: {str(e)}")
            raise
            
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
            
    def create_presentation(self, title, topic, num_slides=5):
        """Create a new presentation with the given title and content."""
        try:
            if not self.service:
                raise ValueError("Slides service not initialized. Call init_service first.")
                
            # Create new presentation
            presentation = {
                'title': title
            }
            
            logger.info(f"Creating new presentation: {title}")
            presentation = self.service.presentations().create(body=presentation).execute()
            presentation_id = presentation.get('presentationId')
            
            if not presentation_id:
                raise ValueError("Failed to get presentation ID from Google Slides API")
                
            logger.info(f"Created presentation with ID: {presentation_id}")
            
            # Generate content using OpenAI
            logger.info(f"Generating content for topic: {topic}")
            content_sections = self._generate_content(topic, num_slides)
            
            if not content_sections:
                raise ValueError("Failed to generate content sections")
                
            # Create title slide
            logger.info("Creating title slide")
            title_requests = self._create_title_slide(title)
            
            # Create content slides
            logger.info("Creating content slides")
            content_requests = self._create_content_slides(content_sections)
            
            # Combine all requests
            requests = title_requests + content_requests
            
            # Execute the requests
            logger.info("Executing slide creation requests")
            body = {
                'requests': requests
            }
            
            response = self.service.presentations().batchUpdate(
                presentationId=presentation_id,
                body=body
            ).execute()
            
            logger.info(f"Successfully updated presentation: {response}")
            return presentation_id
            
        except Exception as e:
            logger.error(f"Error creating presentation: {str(e)}")
            if isinstance(e, HttpError):
                logger.error(f"Google API error response: {e.resp.status} - {e.content}")
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
        
    def _create_title_slide(self, title):
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
        
        return requests
        
    def _create_overview_slide(self, presentation_id, topic):
        """Create an overview slide."""
        # Implementation similar to title slide but with different layout
        pass
        
    def _generate_content(self, topic, num_slides):
        """Generate content for the presentation using OpenAI."""
        try:
            logger.info(f"Generating content for {num_slides} slides")
            
            prompt = f"""Create an outline for a {num_slides}-slide presentation about {topic}.
            For each section, provide 5 concise bullet points.
            Format as JSON with this structure:
            {{
                "sections": [
                    {{
                        "title": "Section Title",
                        "points": ["Point 1", "Point 2", "Point 3", "Point 4", "Point 5"]
                    }}
                ]
            }}
            """
            
            # Make API call using the client
            response = openai_client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful presentation content creator."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7
            )
            
            # Parse response
            content = response.choices[0].message.content
            try:
                data = json.loads(content)
                sections = data.get('sections', [])
                
                if not sections:
                    raise ValueError("No sections found in OpenAI response")
                    
                if len(sections) != num_slides:
                    raise ValueError(f"Got {len(sections)} sections, expected {num_slides}")
                    
                # Validate each section has exactly 5 points
                for section in sections:
                    if len(section.get('points', [])) != 5:
                        raise ValueError("Each section must have exactly 5 points")
                    
                return sections
                
            except json.JSONDecodeError as e:
                logger.error(f"Failed to parse OpenAI response: {str(e)}")
                raise ValueError("Invalid JSON response from OpenAI")
                
        except Exception as e:
            logger.error(f"Error generating content: {str(e)}")
            raise
            
    def _create_content_slides(self, content_sections):
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
            
            requests = []
            
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
                
                requests.append(slide)
                
                slide_id = f'slide_{i+1}'
                title = section.get('title', '')
                points = section.get('points', [])
                
                # Create requests for slide elements
                requests.extend([
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
                            'fields': 'shapeBackgroundFill.solidFill.color',
                            'shapeProperties': {
                                'shapeBackgroundFill': {
                                    'solidFill': {
                                        'color': {'rgbColor': color_scheme['background']}
                                    }
                                }
                            }
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
                    # Style title box
                    {
                        'updateShapeProperties': {
                            'objectId': f'title_{i+1}',
                            'fields': 'outline.propertyState,shapeBackgroundFill.propertyState',
                            'shapeProperties': {
                                'outline': {'propertyState': 'NOT_RENDERED'},
                                'shapeBackgroundFill': {'propertyState': 'NOT_RENDERED'}
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
                            'fields': 'fontFamily,fontSize,foregroundColor,bold',
                            'style': {
                                'fontFamily': 'Roboto',
                                'fontSize': {'magnitude': 32, 'unit': 'PT'},
                                'foregroundColor': {'opaqueColor': {'rgbColor': color_scheme['title']}},
                                'bold': True
                            }
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
                    },
                    # Style content box
                    {
                        'updateShapeProperties': {
                            'objectId': f'content_{i+1}',
                            'fields': 'outline.propertyState,shapeBackgroundFill.propertyState',
                            'shapeProperties': {
                                'outline': {'propertyState': 'NOT_RENDERED'},
                                'shapeBackgroundFill': {'propertyState': 'NOT_RENDERED'}
                            }
                        }
                    }
                ])
                
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
                            'fields': 'fontFamily,fontSize,foregroundColor',
                            'style': {
                                'fontFamily': 'Roboto',
                                'fontSize': {'magnitude': 20, 'unit': 'PT'},
                                'foregroundColor': {'opaqueColor': {'rgbColor': color_scheme['text']}}
                            }
                        }
                    },
                    # Add paragraph spacing
                    {
                        'updateParagraphStyle': {
                            'objectId': f'content_{i+1}',
                            'fields': 'spaceAbove,spaceBelow,indentStart,indentFirstLine,alignment,direction',
                            'style': {
                                'spaceAbove': {'magnitude': 12, 'unit': 'PT'},
                                'spaceBelow': {'magnitude': 12, 'unit': 'PT'},
                                'indentStart': {'magnitude': 20, 'unit': 'PT'},
                                'indentFirstLine': {'magnitude': -20, 'unit': 'PT'},
                                'alignment': 'START',
                                'direction': 'LEFT_TO_RIGHT'
                            }
                        }
                    }
                ])
                
            logger.info(f"Created {len(content_sections)} content slides")
            return requests
            
        except Exception as e:
            logger.error(f"Error creating content slides: {str(e)}")
            raise
