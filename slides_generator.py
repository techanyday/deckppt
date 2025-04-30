import os
import json
import logging
import random
import uuid
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from flask import url_for, session
import openai

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configure OpenAI
openai_client = openai.OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

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
    SECTION_HEADER = 'SECTION_HEADER'
    TITLE_ONLY = 'TITLE_ONLY'
    BLANK = 'BLANK'

class PlaceholderType:
    """Slide placeholder types."""
    TITLE = 'TITLE'
    BODY = 'BODY'
    CENTERED_TITLE = 'CENTERED_TITLE'
    SUBTITLE = 'SUBTITLE'
    SLIDE_NUMBER = 'SLIDE_NUMBER'

class SlideTheme:
    """Theme colors and styles."""
    THEMES = [
        {
            'name': 'Modern Blue',
            'colors': {
                'primary': {'solid': {'color': {'rgbColor': {'red': 0.27, 'green': 0.49, 'blue': 0.75}}}},
                'text': {'solid': {'color': {'rgbColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2}}}}
            }
        },
        {
            'name': 'Forest Green',
            'colors': {
                'primary': {'solid': {'color': {'rgbColor': {'red': 0.22, 'green': 0.60, 'blue': 0.40}}}},
                'text': {'solid': {'color': {'rgbColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2}}}}
            }
        },
        {
            'name': 'Ocean Breeze',
            'colors': {
                'primary': {'solid': {'color': {'rgbColor': {'red': 0.0, 'green': 0.4, 'blue': 0.8}}}},
                'text': {'solid': {'color': {'rgbColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2}}}}
            }
        },
        {
            'name': 'Lavender Dream',
            'colors': {
                'primary': {'solid': {'color': {'rgbColor': {'red': 0.4, 'green': 0.2, 'blue': 0.8}}}},
                'text': {'solid': {'color': {'rgbColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2}}}}
            }
        },
        {
            'name': 'Forest Fresh',
            'colors': {
                'primary': {'solid': {'color': {'rgbColor': {'red': 0.0, 'green': 0.6, 'blue': 0.4}}}},
                'text': {'solid': {'color': {'rgbColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2}}}}
            }
        }
    ]
    
    @classmethod
    def get_random_theme(cls):
        """Get a random theme."""
        return random.choice(cls.THEMES)

class GoogleSlidesGenerator:
    def __init__(self):
        """Initialize the Google Slides generator."""
        self.service = None
        self.theme = None
        self.current_slide_index = 0
        self.layouts = [
            SlideLayout.TITLE_AND_BODY,
            SlideLayout.TITLE_AND_TWO_COLUMNS,
            SlideLayout.TITLE_AND_BODY,
            SlideLayout.SECTION_HEADER
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
        layout = self.layouts[self.current_slide_index]
        self.current_slide_index = (self.current_slide_index + 1) % len(self.layouts)
        return layout

    def get_unique_id(self, base_id):
        """Generate a unique ID for a slide element."""
        unique_id = f"{base_id}_{uuid.uuid4().hex[:8]}"
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
            self._create_content_slides(presentation_id, sections)
            
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
                'color': self.theme['colors']['primary']['solid']['color']
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

    def _create_content_slides(self, presentation_id, sections):
        """Create content slides with proper error handling and validation."""
        requests = []
        
        for index, section in enumerate(sections):
            title = section.get('title', '').strip()
            points = [p.strip() for p in section.get('points', []) if p.strip()]
            
            # Skip invalid sections
            if not title or len(points) < 3:
                logger.warning(f"Skipping slide #{index + 1} due to insufficient content")
                continue
                
            # Generate unique IDs
            slide_id = self.get_unique_id('slide')
            title_id = self.get_unique_id('title')
            body_id = self.get_unique_id('body')
            
            # Create slide
            requests.append({
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {'predefinedLayout': SlideLayout.TITLE_AND_BODY},
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
            })
            
            # Insert title
            requests.append({
                'insertText': {
                    'objectId': title_id,
                    'text': title
                }
            })
            
            # Style title
            requests.append({
                'updateTextStyle': {
                    'objectId': title_id,
                    'style': {
                        'foregroundColor': {
                            'opaqueColor': {
                                'rgbColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2}
                            }
                        },
                        'fontSize': {'magnitude': 24, 'unit': 'PT'},
                        'bold': True
                    },
                    'fields': 'foregroundColor,fontSize,bold'
                }
            })
            
            # Format bullet points
            bullet_text = ''
            for point in points:
                point = point.strip()
                if not point.endswith(('.', '!', '?')):
                    point += '.'
                bullet_text += f"{point}\n"
            
            # Insert bullet points
            requests.append({
                'insertText': {
                    'objectId': body_id,
                    'text': bullet_text.strip()
                }
            })
            
            # Style bullet points
            requests.append({
                'updateTextStyle': {
                    'objectId': body_id,
                    'style': {
                        'foregroundColor': {
                            'opaqueColor': {
                                'rgbColor': {'red': 0.3, 'green': 0.3, 'blue': 0.3}
                            }
                        },
                        'fontSize': {'magnitude': 18, 'unit': 'PT'}
                    },
                    'fields': 'foregroundColor,fontSize'
                }
            })
            
            # Add bullets
            requests.append({
                'createParagraphBullets': {
                    'objectId': body_id,
                    'textRange': {'type': 'ALL'},
                    'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                }
            })
            
            self.current_slide_index += 1
        
        # Execute requests if any valid slides were created
        if requests:
            try:
                self.service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
            except Exception as e:
                logger.error(f"Error creating slides: {str(e)}")
                raise
        else:
            logger.error("No valid slides to create")
            raise ValueError("Failed to create any valid slides")

    def _generate_content(self, topic, num_slides):
        """Generate content for the presentation using OpenAI."""
        try:
            logger.info(f"Generating content for {num_slides} slides")
            
            # Create a detailed prompt
            prompt = f"""Create a detailed presentation outline on "{topic}" with {num_slides} sections.
            For each section, provide:
            1. A clear, engaging title (2-5 words)
            2. 4-6 detailed bullet points that:
               - Are complete thoughts (10-15 words each)
               - Include specific examples, data, or insights
               - Flow logically from one point to the next
               - Avoid vague statements
            
            Format as JSON:
            {{
                "sections": [
                    {{
                        "title": "Section Title",
                        "points": [
                            "Detailed point 1 with specific example",
                            "Detailed point 2 with data or insight",
                            ...
                        ]
                    }},
                    ...
                ]
            }}
            
            Make sure:
            - First section introduces the topic
            - Middle sections develop key ideas
            - Final section concludes with takeaways
            - Each point is substantive and informative
            - No placeholder or generic content"""

            # Get completion from OpenAI using new client interface
            client = openai.OpenAI()
            response = client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
                max_tokens=2000
            )

            # Parse response using new response format
            content = response.choices[0].message.content.strip()
            
            try:
                # Parse JSON response
                data = json.loads(content)
                
                # Validate response format
                if not isinstance(data, dict) or 'sections' not in data:
                    raise ValueError("Invalid response format")
                
                sections = data['sections']
                if not sections or len(sections) < num_slides:
                    raise ValueError(f"Not enough sections generated (got {len(sections)}, need {num_slides})")
                
                # Validate each section
                for section in sections:
                    if not isinstance(section, dict):
                        raise ValueError("Invalid section format")
                    
                    if 'title' not in section or not section['title'].strip():
                        raise ValueError("Missing or empty section title")
                    
                    if 'points' not in section or not section['points']:
                        raise ValueError("Missing or empty section points")
                    
                    if len(section['points']) < 3:
                        raise ValueError(f"Not enough points in section '{section['title']}'")
                
                return sections[:num_slides]
                
            except json.JSONDecodeError:
                raise ValueError("Failed to parse OpenAI response as JSON")
                
        except Exception as e:
            logger.error(f"Error generating content: {str(e)}")
            raise ValueError("Failed to generate presentation content")
