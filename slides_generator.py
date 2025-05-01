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
    BIG_NUMBER = 'BIG_NUMBER'

class PlaceholderType:
    """Slide placeholder types."""
    TITLE = 'TITLE'
    BODY = 'BODY'
    CENTERED_TITLE = 'CENTERED_TITLE'
    SUBTITLE = 'SUBTITLE'
    SLIDE_NUMBER = 'SLIDE_NUMBER'

class GoogleSlidesGenerator:
    def __init__(self, credentials_path=None):
        self.service = self._create_slides_service(credentials_path)
        # Modern color palette
        self.theme = {
            'primary': {'r': 0.27, 'g': 0.36, 'b': 0.87},  # Royal Blue
            'secondary': {'r': 0.95, 'g': 0.49, 'b': 0.33},  # Coral
            'accent': {'r': 0.33, 'g': 0.78, 'b': 0.69},  # Teal
            'background': {'r': 0.98, 'g': 0.98, 'b': 0.98},  # Light Gray
            'text': {'r': 0.13, 'g': 0.13, 'b': 0.13}  # Dark Gray
        }

    def _create_slides_service(self, credentials_path=None):
        """Initialize the Google Slides service with credentials."""
        try:
            # Try to get credentials from environment variable first
            creds_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
            if creds_json:
                try:
                    creds_dict = json.loads(creds_json)
                    credentials = Credentials.from_service_account_info(
                        creds_dict,
                        scopes=SCOPES
                    )
                    logger.info("Using credentials from environment variable")
                except Exception as e:
                    logger.error(f"Error parsing credentials JSON from env: {str(e)}")
                    raise ValueError("Invalid credentials JSON in environment")
            
            # Fall back to file if provided
            elif credentials_path and os.path.exists(credentials_path):
                credentials = Credentials.from_service_account_file(
                    credentials_path,
                    scopes=SCOPES
                )
                logger.info(f"Using credentials from file: {credentials_path}")
            
            # Try default credentials as last resort
            elif os.path.exists('token.json'):
                credentials = Credentials.from_authorized_user_file(
                    'token.json', SCOPES)
                logger.info("Using default credentials from token.json")
            
            else:
                logger.error("No valid credentials found")
                raise ValueError(
                    "No valid credentials found. Please set GOOGLE_CREDENTIALS_JSON environment variable "
                    "or provide a valid credentials file."
                )
                
            # Build the service
            service = build('slides', 'v1', credentials=credentials)
            logger.info("Successfully initialized Slides service")
            return service
            
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
            
    def get_unique_id(self, base_id):
        """Generate a unique ID for a slide element."""
        unique_id = f"{base_id}_{uuid.uuid4().hex[:8]}"
        return unique_id

    def _create_title_slide(self, presentation_id, title, subtitle="Generated by DeckSky"):
        """Create an attractive title slide."""
        slide_id = f"title_{uuid.uuid4().hex[:8]}"
        requests = [
            # Create title slide
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {'predefinedLayout': 'TITLE'},
                    'placeholderIdMappings': [
                        {
                            'layoutPlaceholder': {'type': 'TITLE', 'index': 0},
                            'objectId': f"{slide_id}_title"
                        },
                        {
                            'layoutPlaceholder': {'type': 'SUBTITLE', 'index': 1},
                            'objectId': f"{slide_id}_subtitle"
                        }
                    ]
                }
            },
            # Set slide background
            {
                'updatePageProperties': {
                    'objectId': slide_id,
                    'pageProperties': {
                        'pageBackgroundFill': {
                            'solidFill': {
                                'color': {
                                    'rgbColor': self.theme['background']
                                }
                            }
                        }
                    },
                    'fields': 'pageBackgroundFill'
                }
            },
            # Insert title
            {
                'insertText': {
                    'objectId': f"{slide_id}_title",
                    'text': title
                }
            },
            # Style title
            {
                'updateTextStyle': {
                    'objectId': f"{slide_id}_title",
                    'style': {
                        'foregroundColor': {
                            'opaqueColor': {'rgbColor': self.theme['primary']}
                        },
                        'fontSize': {'magnitude': 40, 'unit': 'PT'},
                        'bold': True,
                        'fontFamily': 'Google Sans'
                    },
                    'textRange': {'type': 'ALL'},
                    'fields': 'foregroundColor,fontSize,bold,fontFamily'
                }
            },
            # Insert subtitle
            {
                'insertText': {
                    'objectId': f"{slide_id}_subtitle",
                    'text': subtitle
                }
            },
            # Style subtitle
            {
                'updateTextStyle': {
                    'objectId': f"{slide_id}_subtitle",
                    'style': {
                        'foregroundColor': {
                            'opaqueColor': {'rgbColor': self.theme['secondary']}
                        },
                        'fontSize': {'magnitude': 24, 'unit': 'PT'},
                        'fontFamily': 'Google Sans'
                    },
                    'textRange': {'type': 'ALL'},
                    'fields': 'foregroundColor,fontSize,fontFamily'
                }
            }
        ]
        
        return requests

    def _get_slide_layout(self, section_index, total_sections, title):
        """Determine appropriate slide layout based on content and position."""
        title_lower = title.lower()
        
        if section_index == 0:
            return 'SECTION_HEADER'  # Introduction
        elif section_index == total_sections - 1:
            return 'BIG_NUMBER'  # Conclusion/Summary
        elif any(word in title_lower for word in ['overview', 'summary', 'key points']):
            return 'TITLE_AND_TWO_COLUMNS'
        elif len(title_lower.split()) <= 3:  # Short titles often indicate section breaks
            return 'SECTION_HEADER'
        else:
            return 'TITLE_AND_BODY'  # Default layout

    def _create_content_slide(self, presentation_id, title, points, section_index, total_sections):
        """Create a content slide with dynamic layout and styling."""
        if not points or len(points) < 2:  # Skip if insufficient content
            return []

        slide_id = f"slide_{uuid.uuid4().hex[:8]}"
        layout = self._get_slide_layout(section_index, total_sections, title)
        
        requests = [
            # Create slide with dynamic layout
            {
                'createSlide': {
                    'objectId': slide_id,
                    'slideLayoutReference': {'predefinedLayout': layout},
                    'placeholderIdMappings': [
                        {
                            'layoutPlaceholder': {'type': 'TITLE', 'index': 0},
                            'objectId': f"{slide_id}_title"
                        },
                        {
                            'layoutPlaceholder': {'type': 'BODY', 'index': 0},
                            'objectId': f"{slide_id}_body"
                        }
                    ]
                }
            },
            # Set slide background with subtle gradient
            {
                'updatePageProperties': {
                    'objectId': slide_id,
                    'pageProperties': {
                        'pageBackgroundFill': {
                            'solidFill': {
                                'color': {
                                    'rgbColor': self.theme['background']
                                },
                                'alpha': 0.95
                            }
                        }
                    },
                    'fields': 'pageBackgroundFill'
                }
            }
        ]

        # Add title
        requests.extend([
            {
                'insertText': {
                    'objectId': f"{slide_id}_title",
                    'text': title
                }
            },
            {
                'updateTextStyle': {
                    'objectId': f"{slide_id}_title",
                    'style': {
                        'foregroundColor': {
                            'opaqueColor': {'rgbColor': self.theme['primary']}
                        },
                        'fontSize': {'magnitude': 28, 'unit': 'PT'},
                        'bold': True,
                        'fontFamily': 'Google Sans'
                    },
                    'textRange': {'type': 'ALL'},
                    'fields': 'foregroundColor,fontSize,bold,fontFamily'
                }
            }
        ])

        # Format bullet points
        bullet_text = ''
        for point in points:
            point = point.strip()
            if not point.endswith(('.', '!', '?')):
                point += '.'
            bullet_text += f"{point}\n"

        # Add content
        requests.extend([
            {
                'insertText': {
                    'objectId': f"{slide_id}_body",
                    'text': bullet_text
                }
            },
            {
                'updateTextStyle': {
                    'objectId': f"{slide_id}_body",
                    'style': {
                        'foregroundColor': {
                            'opaqueColor': {'rgbColor': self.theme['text']}
                        },
                        'fontSize': {'magnitude': 18, 'unit': 'PT'},
                        'fontFamily': 'Google Sans'
                    },
                    'textRange': {'type': 'ALL'},
                    'fields': 'foregroundColor,fontSize,fontFamily'
                }
            },
            {
                'createParagraphBullets': {
                    'objectId': f"{slide_id}_body",
                    'textRange': {'type': 'ALL'},
                    'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                }
            }
        ])

        return requests

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

    def create_presentation(self, title, topic, num_slides=5):
        """Create a presentation with consistent styling and layout."""
        try:
            # Create new presentation
            presentation = {'title': title}
            presentation = self.service.presentations().create(body=presentation).execute()
            presentation_id = presentation.get('presentationId')

            # Generate content
            sections = self._generate_content(topic, num_slides)
            if not sections:
                raise ValueError("No content generated")

            # Start with requests for title slide
            requests = self._create_title_slide(presentation_id, title)

            # Add content slides
            for i, section in enumerate(sections):
                if not section.get('points'):  # Skip sections without content
                    continue
                slide_requests = self._create_content_slide(
                    presentation_id,
                    section['title'],
                    section['points'],
                    i,
                    len(sections)
                )
                requests.extend(slide_requests)

            # Execute all requests
            body = {'requests': requests}
            response = self.service.presentations().batchUpdate(
                presentationId=presentation_id, body=body).execute()

            return presentation_id

        except Exception as e:
            logger.error(f"Error creating presentation: {str(e)}")
            raise ValueError("Failed to create presentation")
