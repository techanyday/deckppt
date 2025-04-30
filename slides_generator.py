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
            # Ensure num_slides is an integer and within bounds
            try:
                num_slides = int(num_slides)
                num_slides = max(1, min(num_slides, 10))  # Limit between 1 and 10 slides
            except (TypeError, ValueError):
                num_slides = 5  # Default to 5 if invalid
                
            logger.info(f"Creating presentation with {num_slides} slides")
            
            # Create new presentation
            presentation = self.service.presentations().create(body={'title': title}).execute()
            presentation_id = presentation.get('presentationId')
            logger.info(f"Created presentation with ID: {presentation_id}")
            
            # Generate content sections (with exact slide count)
            content_sections = self._generate_content_sections(topic, num_slides)
            
            # Create content slides
            self._create_content_slides(presentation_id, content_sections)
            
            return presentation_id
            
        except Exception as e:
            logger.error(f"An error occurred: {str(e)}")
            raise
            
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
            # Get OpenAI client
            from openai import OpenAI
            client = OpenAI()
            
            system_prompt = f"""You are a presentation content generator. Create informative and engaging content for a presentation.

            CRITICAL INSTRUCTION:
            You MUST generate EXACTLY {num_sections} slides, no more and no less.
            
            Each slide section MUST have:
            1. A clear, concise title (2-5 words)
            2. Exactly 5 bullet points (5-12 words each)
            3. Each bullet point should be a complete thought
            4. Content should be factual and professional
            
            Format the response as a JSON array with EXACTLY {num_sections} objects.
            Each object must have 'title' and 'points' fields.
            
            Example format (showing 1 slide):
            [
                {{
                    "title": "Market Overview",
                    "points": [
                        "Global market expected to reach $500B by 2025",
                        "North America leads with 35% market share",
                        "Asia Pacific showing fastest growth rate of 12%",
                        "Top three players control 45% of market",
                        "New entrants focus on innovative technologies"
                    ]
                }}
            ]

            FINAL CHECK:
            - Your response MUST contain exactly {num_sections} slide objects
            - Each slide MUST have exactly 5 bullet points
            - DO NOT generate more than {num_sections} slides"""

            messages = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"Generate a presentation with EXACTLY {num_sections} slides about: {topic}"}
            ]

            max_retries = 3
            retry_count = 0
            
            while retry_count < max_retries:
                try:
                    response = client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=messages,
                        temperature=0.7,
                        max_tokens=1000,
                        top_p=0.9,
                        frequency_penalty=0.2,
                        presence_penalty=0.2
                    )
                    
                    content = response.choices[0].message.content
                    
                    # Validate JSON response
                    if not self._validate_json_response(content):
                        logger.warning(f"Invalid JSON response on attempt {retry_count + 1}")
                        retry_count += 1
                        continue
                    
                    # Parse the response
                    content_sections = json.loads(content)
                    
                    # Force exact section count (failsafe)
                    if len(content_sections) > num_sections:
                        logger.warning(f"Truncating from {len(content_sections)} to {num_sections} sections")
                        content_sections = content_sections[:num_sections]
                    elif len(content_sections) < num_sections:
                        logger.warning(f"Too few sections generated (got {len(content_sections)}, expected {num_sections})")
                        retry_count += 1
                        continue
                        
                    # Validate each section has exactly 5 points
                    valid_sections = True
                    for section in content_sections:
                        if len(section.get('points', [])) != 5:
                            valid_sections = False
                            break
                            
                    if not valid_sections:
                        logger.warning("Some sections don't have exactly 5 points")
                        retry_count += 1
                        continue
                    
                    return content_sections
                    
                except Exception as e:
                    logger.error(f"Error in OpenAI request: {str(e)}")
                    retry_count += 1
                    if retry_count == max_retries:
                        raise Exception("Failed to generate valid content after multiple attempts")
                    
            raise Exception("Failed to generate valid content after multiple attempts")
            
        except Exception as e:
            logger.error(f"Error generating content: {str(e)}")
            raise

    def _validate_json_response(self, response_text):
        """Validate if the JSON response is complete and well-formed."""
        try:
            # Check if response ends with proper JSON closure
            if not response_text.strip().endswith(']'):
                logger.error("Incomplete JSON response")
                return False
            
            # Try parsing the JSON
            json.loads(response_text)
            return True
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON: {str(e)}")
            return False

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
                
                # Execute the requests
                self.service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
                
            logger.info(f"Created {len(content_sections)} content slides")
            
        except Exception as e:
            logger.error(f"Error creating content slides: {str(e)}")
            raise
