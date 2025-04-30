import os
import json
import logging
import random
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

class SlideLayout:
    """Predefined slide layouts."""
    TITLE = 'TITLE'
    TITLE_AND_BODY = 'TITLE_AND_BODY'
    TITLE_AND_TWO_COLUMNS = 'TITLE_AND_TWO_COLUMNS'
    SECTION = 'SECTION_HEADER'
    TITLE_ONLY = 'TITLE_ONLY'
    BLANK = 'BLANK'

class PlaceholderType:
    """Predefined placeholder types."""
    TITLE = 'TITLE'
    BODY = 'BODY'
    CENTERED_TITLE = 'CENTERED_TITLE'
    SUBTITLE = 'SUBTITLE'
    SLIDE_NUMBER = 'SLIDE_NUMBER'

class SlideTheme:
    """Predefined color themes."""
    THEMES = [
        {
            'name': 'Ocean Breeze',
            'colors': {
                'primary': {'solid': {'color': {'rgbColor': {'red': 0.0, 'green': 0.4, 'blue': 0.8}}}},
                'secondary': {'solid': {'color': {'rgbColor': {'red': 0.8, 'green': 0.9, 'blue': 1.0}}}},
                'accent': {'solid': {'color': {'rgbColor': {'red': 0.0, 'green': 0.6, 'blue': 0.4}}}},
                'background': {'solid': {'color': {'rgbColor': {'red': 0.95, 'green': 0.98, 'blue': 1.0}}}}
            }
        },
        {
            'name': 'Lavender Dream',
            'colors': {
                'primary': {'solid': {'color': {'rgbColor': {'red': 0.4, 'green': 0.2, 'blue': 0.8}}}},
                'secondary': {'solid': {'color': {'rgbColor': {'red': 0.9, 'green': 0.85, 'blue': 1.0}}}},
                'accent': {'solid': {'color': {'rgbColor': {'red': 0.6, 'green': 0.2, 'blue': 0.8}}}},
                'background': {'solid': {'color': {'rgbColor': {'red': 0.98, 'green': 0.95, 'blue': 1.0}}}}
            }
        },
        {
            'name': 'Forest Fresh',
            'colors': {
                'primary': {'solid': {'color': {'rgbColor': {'red': 0.0, 'green': 0.6, 'blue': 0.4}}}},
                'secondary': {'solid': {'color': {'rgbColor': {'red': 0.85, 'green': 0.95, 'blue': 0.9}}}},
                'accent': {'solid': {'color': {'rgbColor': {'red': 0.2, 'green': 0.8, 'blue': 0.4}}}},
                'background': {'solid': {'color': {'rgbColor': {'red': 0.95, 'green': 1.0, 'blue': 0.98}}}}
            }
        }
    ]

    @classmethod
    def get_random_theme(cls):
        """Get a random color theme."""
        return random.choice(cls.THEMES)

class GoogleSlidesGenerator:
    def __init__(self):
        """Initialize the Google Slides generator."""
        self.service = None
        self.theme = None
        self.current_layout_index = 0
        self.current_slide_index = 0
        self.layouts = [
            SlideLayout.TITLE_AND_BODY,
            SlideLayout.TITLE_AND_TWO_COLUMNS,
            SlideLayout.TITLE_AND_BODY,
            SlideLayout.SECTION
        ]
        
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
            
    def get_next_layout(self):
        """Get the next layout in rotation."""
        layout = self.layouts[self.current_layout_index]
        self.current_layout_index = (self.current_layout_index + 1) % len(self.layouts)
        return layout

    def get_unique_id(self, base_id):
        """Generate a unique ID for a slide element."""
        unique_id = f"{base_id}_{self.current_slide_index}"
        return unique_id

    def create_presentation(self, title, topic, num_slides=5):
        """Create a new presentation with the specified title and theme."""
        try:
            logger.info("Creating new presentation")
            
            # Reset slide index
            self.current_slide_index = 0
            
            # Create presentation
            presentation = {
                'title': title
            }
            presentation = self.service.presentations().create(body=presentation).execute()
            presentation_id = presentation.get('presentationId')
            
            if not presentation_id:
                raise ValueError("Failed to create presentation - no ID returned")
            
            # Get random theme
            self.theme = SlideTheme.get_random_theme()
            
            # Create title slide
            self._create_title_slide(presentation_id, title)
            
            # Generate content
            sections = self._generate_content(topic, num_slides)
            
            # Create content slides
            for section in sections:
                layout = self.get_next_layout()
                if layout == SlideLayout.SECTION:
                    self._create_section_slide(presentation_id, section['title'])
                else:
                    self._create_content_slide(presentation_id, section['title'], section['points'], layout)
            
            # Get the presentation metadata to verify it exists
            presentation = self.service.presentations().get(presentationId=presentation_id).execute()
            
            # Return presentation URL
            return {
                'id': presentation_id,
                'url': f"https://docs.google.com/presentation/d/{presentation_id}/edit",
                'title': presentation.get('title', title)
            }
            
        except Exception as e:
            logger.error(f"Error creating presentation: {str(e)}")
            raise

    def _create_title_slide(self, presentation_id, title):
        """Create the title slide with theme colors."""
        # Convert theme colors to API format
        background_fill = {
            'solidFill': {
                'color': self.theme['colors']['background']['solid']['color']
            }
        }
        text_color = {
            'opaqueColor': self.theme['colors']['primary']['solid']['color']
        }

        # Generate unique IDs
        title_id = self.get_unique_id('title')
        background_id = self.get_unique_id('background')
        page_id = self.get_unique_id('page')

        requests = [{
            'createSlide': {
                'objectId': page_id,
                'slideLayoutReference': {'predefinedLayout': SlideLayout.TITLE},
                'placeholderIdMappings': [
                    {
                        'layoutPlaceholder': {'type': PlaceholderType.CENTERED_TITLE},
                        'objectId': title_id
                    }
                ]
            }
        }, {
            'createShape': {
                'objectId': background_id,
                'shapeType': 'RECTANGLE',
                'elementProperties': {
                    'pageObjectId': page_id,
                    'size': {'width': {'magnitude': 720, 'unit': 'PT'},
                            'height': {'magnitude': 405, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1,
                                'translateX': 0, 'translateY': 0, 'unit': 'PT'}
                }
            }
        }, {
            'updateShapeProperties': {
                'objectId': background_id,
                'shapeProperties': {
                    'shapeBackgroundFill': background_fill
                },
                'fields': 'shapeBackgroundFill'
            }
        }, {
            'insertText': {
                'objectId': title_id,
                'text': title
            }
        }, {
            'updateTextStyle': {
                'objectId': title_id,
                'style': {
                    'foregroundColor': text_color,
                    'fontFamily': 'Montserrat',
                    'fontSize': {'magnitude': 40, 'unit': 'PT'},
                    'bold': True
                },
                'fields': 'foregroundColor,fontFamily,fontSize,bold'
            }
        }]
        
        self.service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': requests}
        ).execute()
        self.current_slide_index += 1

    def _create_section_slide(self, presentation_id, title):
        """Create a section break slide."""
        # Convert theme colors to API format
        background_fill = {
            'solidFill': {
                'color': self.theme['colors']['secondary']['solid']['color']
            }
        }
        text_color = {
            'opaqueColor': self.theme['colors']['primary']['solid']['color']
        }

        # Generate unique IDs
        title_id = self.get_unique_id('title')
        background_id = self.get_unique_id('background')
        page_id = self.get_unique_id('page')

        requests = [{
            'createSlide': {
                'objectId': page_id,
                'slideLayoutReference': {'predefinedLayout': SlideLayout.SECTION},
                'placeholderIdMappings': [
                    {
                        'layoutPlaceholder': {'type': PlaceholderType.TITLE},
                        'objectId': title_id
                    }
                ]
            }
        }, {
            'createShape': {
                'objectId': background_id,
                'shapeType': 'RECTANGLE',
                'elementProperties': {
                    'pageObjectId': page_id,
                    'size': {'width': {'magnitude': 720, 'unit': 'PT'},
                            'height': {'magnitude': 405, 'unit': 'PT'}},
                    'transform': {'scaleX': 1, 'scaleY': 1,
                                'translateX': 0, 'translateY': 0, 'unit': 'PT'}
                }
            }
        }, {
            'updateShapeProperties': {
                'objectId': background_id,
                'shapeProperties': {
                    'shapeBackgroundFill': background_fill
                },
                'fields': 'shapeBackgroundFill'
            }
        }, {
            'insertText': {
                'objectId': title_id,
                'text': title
            }
        }, {
            'updateTextStyle': {
                'objectId': title_id,
                'style': {
                    'foregroundColor': text_color,
                    'fontFamily': 'Montserrat',
                    'fontSize': {'magnitude': 36, 'unit': 'PT'},
                    'bold': True
                },
                'fields': 'foregroundColor,fontFamily,fontSize,bold'
            }
        }]
        
        self.service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': requests}
        ).execute()
        self.current_slide_index += 1

    def _create_content_slide(self, presentation_id, title, points, layout):
        """Create a content slide with the specified layout."""
        # Convert theme colors to API format
        background_fill = {
            'solidFill': {
                'color': self.theme['colors']['background']['solid']['color']
            }
        }
        text_color = {
            'opaqueColor': self.theme['colors']['primary']['solid']['color']
        }
        body_color = {
            'opaqueColor': {'rgbColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2}}
        }

        # Generate unique IDs
        title_id = self.get_unique_id('title')
        background_id = self.get_unique_id('background')
        body_id = self.get_unique_id('body')
        page_id = self.get_unique_id('page')

        requests = []
        
        # Create slide with layout
        slide_request = {
            'createSlide': {
                'objectId': page_id,
                'slideLayoutReference': {'predefinedLayout': layout},
                'placeholderIdMappings': [
                    {
                        'layoutPlaceholder': {'type': PlaceholderType.TITLE},
                        'objectId': title_id
                    },
                    {
                        'layoutPlaceholder': {'type': PlaceholderType.BODY},
                        'objectId': body_id
                    }
                ]
            }
        }
        
        requests.append(slide_request)
        
        # Add background
        requests.extend([
            {
                'createShape': {
                    'objectId': background_id,
                    'shapeType': 'RECTANGLE',
                    'elementProperties': {
                        'pageObjectId': page_id,
                        'size': {'width': {'magnitude': 720, 'unit': 'PT'},
                                'height': {'magnitude': 405, 'unit': 'PT'}},
                        'transform': {'scaleX': 1, 'scaleY': 1,
                                    'translateX': 0, 'translateY': 0, 'unit': 'PT'}
                    }
                }
            },
            {
                'updateShapeProperties': {
                    'objectId': background_id,
                    'shapeProperties': {
                        'shapeBackgroundFill': background_fill
                    },
                    'fields': 'shapeBackgroundFill'
                }
            }
        ])
        
        # Add title
        requests.extend([
            {
                'insertText': {
                    'objectId': title_id,
                    'text': title
                }
            },
            {
                'updateTextStyle': {
                    'objectId': title_id,
                    'style': {
                        'foregroundColor': text_color,
                        'fontFamily': 'Montserrat',
                        'fontSize': {'magnitude': 24, 'unit': 'PT'},
                        'bold': True
                    },
                    'fields': 'foregroundColor,fontFamily,fontSize,bold'
                }
            }
        ])
        
        if layout == SlideLayout.TITLE_AND_TWO_COLUMNS:
            # Split points into two columns
            mid = len(points) // 2
            left_points = points[:mid]
            right_points = points[mid:]
            
            # Create left column shape
            left_column_id = self.get_unique_id('leftColumn')
            requests.extend([
                {
                    'createShape': {
                        'objectId': left_column_id,
                        'shapeType': 'TEXT_BOX',
                        'elementProperties': {
                            'pageObjectId': page_id,
                            'size': {'width': {'magnitude': 320, 'unit': 'PT'},
                                    'height': {'magnitude': 300, 'unit': 'PT'}},
                            'transform': {'scaleX': 1, 'scaleY': 1,
                                        'translateX': 40, 'translateY': 100,
                                        'unit': 'PT'}
                        }
                    }
                },
                {
                    'insertText': {
                        'objectId': left_column_id,
                        'text': '\n'.join(f'• {point}' for point in left_points)
                    }
                },
                {
                    'updateTextStyle': {
                        'objectId': left_column_id,
                        'style': {
                            'foregroundColor': body_color,
                            'fontFamily': 'Roboto',
                            'fontSize': {'magnitude': 18, 'unit': 'PT'}
                        },
                        'fields': 'foregroundColor,fontFamily,fontSize'
                    }
                }
            ])
            
            # Create right column shape
            right_column_id = self.get_unique_id('rightColumn')
            requests.extend([
                {
                    'createShape': {
                        'objectId': right_column_id,
                        'shapeType': 'TEXT_BOX',
                        'elementProperties': {
                            'pageObjectId': page_id,
                            'size': {'width': {'magnitude': 320, 'unit': 'PT'},
                                    'height': {'magnitude': 300, 'unit': 'PT'}},
                            'transform': {'scaleX': 1, 'scaleY': 1,
                                        'translateX': 380, 'translateY': 100,
                                        'unit': 'PT'}
                        }
                    }
                },
                {
                    'insertText': {
                        'objectId': right_column_id,
                        'text': '\n'.join(f'• {point}' for point in right_points)
                    }
                },
                {
                    'updateTextStyle': {
                        'objectId': right_column_id,
                        'style': {
                            'foregroundColor': body_color,
                            'fontFamily': 'Roboto',
                            'fontSize': {'magnitude': 18, 'unit': 'PT'}
                        },
                        'fields': 'foregroundColor,fontFamily,fontSize'
                    }
                }
            ])
        else:
            # Add content to single column
            requests.extend([
                {
                    'insertText': {
                        'objectId': body_id,
                        'text': '\n'.join(f'• {point}' for point in points)
                    }
                },
                {
                    'updateTextStyle': {
                        'objectId': body_id,
                        'style': {
                            'foregroundColor': body_color,
                            'fontFamily': 'Roboto',
                            'fontSize': {'magnitude': 18, 'unit': 'PT'}
                        },
                        'fields': 'foregroundColor,fontFamily,fontSize'
                    }
                }
            ])
        
        self.service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': requests}
        ).execute()
        self.current_slide_index += 1

    def _generate_content(self, topic, num_slides):
        """Generate content for the presentation using OpenAI."""
        try:
            logger.info(f"Generating content for {num_slides} slides")
            
            prompt = f"""Create an outline for a {num_slides}-slide presentation about {topic}.
            Return only a JSON object with this structure:
            {{
                "sections": [
                    {{
                        "title": "Section Title",
                        "points": ["Point 1", "Point 2", "Point 3", "Point 4", "Point 5"]
                    }}
                ]
            }}
            Each section must have exactly 5 points.
            The total number of sections must be exactly {num_slides}.
            Keep points concise and impactful.
            Do not include any markdown formatting, code blocks, or extra text.
            """
            
            # Make API call using the client
            response = openai_client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful presentation content creator. Always return valid JSON without any markdown formatting or extra text."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7
            )
            
            # Get response content and clean it
            content = response.choices[0].message.content.strip()
            
            # Remove any markdown formatting
            if content.startswith('```'):
                content = content[content.find('{'):content.rfind('}')+1]
            
            logger.info(f"Got OpenAI response: {content[:100]}...")
            
            try:
                # Parse JSON response
                data = json.loads(content)
                sections = data.get('sections', [])
                
                # Validate sections
                if not sections:
                    raise ValueError("No sections found in OpenAI response")
                    
                if len(sections) != num_slides:
                    raise ValueError(f"Got {len(sections)} sections, expected {num_slides}")
                    
                # Validate each section
                for section in sections:
                    if not isinstance(section, dict):
                        raise ValueError("Invalid section format")
                        
                    if 'title' not in section or 'points' not in section:
                        raise ValueError("Section missing title or points")
                        
                    if not isinstance(section['points'], list):
                        raise ValueError("Points must be a list")
                        
                    if len(section['points']) != 5:
                        raise ValueError(f"Section '{section['title']}' has {len(section['points'])} points, expected 5")
                        
                    # Clean up points
                    section['points'] = [
                        point.strip().strip('•').strip('-').strip()
                        for point in section['points']
                    ]
                    
                return sections
                
            except json.JSONDecodeError as e:
                logger.error(f"Failed to parse OpenAI response: {str(e)}")
                raise ValueError("Invalid JSON response from OpenAI")
                
        except Exception as e:
            logger.error(f"Error generating content: {str(e)}")
            raise
