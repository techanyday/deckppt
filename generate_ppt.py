import os
import logging
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from apis.openai_api import OpenAIClient
import re
import tempfile

def get_theme_layout_ids(theme):
    """Get layout IDs and theme colors for different themes"""
    themes = {
        "professional": {
            "layouts": {
                "title": 0,
                "content": 1,
                "section": 2,
                "two_content": 3,
                "comparison": 4
            },
            "colors": {
                "background": "2B579A",  # Dark blue
                "title": "FFFFFF",       # White
                "accent": "E6F0FF"       # Light blue
            }
        },
        "modern": {
            "layouts": {
                "title": 0,
                "content": 1,
                "section": 2,
                "two_content": 3,
                "comparison": 4
            },
            "colors": {
                "background": "292929",  # Dark gray
                "title": "FFFFFF",       # White
                "accent": "00BFA5"       # Teal
            }
        },
        "minimal": {
            "layouts": {
                "title": 0,
                "content": 1,
                "section": 2,
                "two_content": 3,
                "comparison": 4
            },
            "colors": {
                "background": "F0F0F0",  # Light gray
                "title": "1A1A1A",       # Almost black
                "accent": "404040"       # Dark gray
            }
        },
        "creative": {
            "layouts": {
                "title": 0,
                "content": 1,
                "section": 2,
                "two_content": 3,
                "comparison": 4
            },
            "colors": {
                "background": "6200EA",  # Deep purple
                "title": "FFFFFF",       # White
                "accent": "B388FF"       # Light purple
            }
        },
        "corporate": {
            "layouts": {
                "title": 0,
                "content": 1,
                "section": 2,
                "two_content": 3,
                "comparison": 4
            },
            "colors": {
                "background": "01579B",  # Dark blue
                "title": "FFFFFF",       # White
                "accent": "81D4FA"       # Light blue
            }
        }
    }
    return themes.get(theme, themes["professional"])

def apply_theme_color(shape, color_hex, is_fill=False):
    """Apply theme color to shape"""
    try:
        rgb = RGBColor.from_string(color_hex)
        if is_fill:
            shape.fill.solid()
            shape.fill.fore_color.rgb = rgb
        else:
            if hasattr(shape, 'font'):
                shape.font.color.rgb = rgb
            elif hasattr(shape, 'text_frame'):
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.font.color.rgb = rgb
    except Exception as e:
        logging.error(f"Error applying theme color: {e}")

def apply_slide_background(slide, color_hex):
    """Apply background color to slide"""
    try:
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(color_hex)
    except Exception as e:
        logging.error(f"Error applying background color: {e}")

def create_title_slide(ppt, title, theme="professional"):
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["layouts"]["title"]]
    slide = ppt.slides.add_slide(layout)
    theme_colors = get_theme_layout_ids(theme)["colors"]
    
    # Apply background color
    apply_slide_background(slide, theme_colors["background"])
    
    # Find title and subtitle placeholders
    title_placeholder = None
    subtitle_placeholder = None
    
    for shape in slide.placeholders:
        if shape.placeholder_format.type == 1:  # Title
            title_placeholder = shape
        elif shape.placeholder_format.type == 2:  # Subtitle
            subtitle_placeholder = shape
    
    # Add title
    if title_placeholder:
        title_placeholder.text = title
        title_placeholder.text_frame.paragraphs[0].font.size = Pt(44)
        title_placeholder.text_frame.paragraphs[0].font.name = 'Calibri'
        apply_theme_color(title_placeholder, theme_colors["title"])
    else:
        # If no title placeholder, create a text box for title
        left = Pt(36)
        top = Pt(180)
        width = Pt(648)
        height = Pt(72)
        title_box = slide.shapes.add_textbox(left, top, width, height)
        tf = title_box.text_frame
        tf.text = title
        tf.paragraphs[0].font.size = Pt(44)
        tf.paragraphs[0].font.name = 'Calibri'
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        apply_theme_color(title_box, theme_colors["title"])
    
    # Add subtitle
    if subtitle_placeholder:
        subtitle_placeholder.text = "Generated with AI"
        subtitle_placeholder.text_frame.paragraphs[0].font.size = Pt(24)
        subtitle_placeholder.text_frame.paragraphs[0].font.name = 'Calibri'
        apply_theme_color(subtitle_placeholder, theme_colors["accent"])
    else:
        # If no subtitle placeholder, create a text box for subtitle
        left = Pt(36)
        top = Pt(288)
        width = Pt(648)
        height = Pt(36)
        subtitle_box = slide.shapes.add_textbox(left, top, width, height)
        tf = subtitle_box.text_frame
        tf.text = "Generated with AI"
        tf.paragraphs[0].font.size = Pt(24)
        tf.paragraphs[0].font.name = 'Calibri'
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        apply_theme_color(subtitle_box, theme_colors["accent"])
    
    return slide

def create_content_slide(ppt, title, content, theme="professional"):
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["layouts"]["content"]]
    slide = ppt.slides.add_slide(layout)
    theme_colors = get_theme_layout_ids(theme)["colors"]
    
    # Apply background color first
    apply_slide_background(slide, theme_colors["background"])
    
    # Find title and content placeholders
    title_placeholder = None
    content_placeholder = None
    
    for shape in slide.placeholders:
        if shape.placeholder_format.type == 1:  # Title
            title_placeholder = shape
        elif shape.placeholder_format.type == 7:  # Content
            content_placeholder = shape
    
    # Add title if placeholder exists
    if title_placeholder:
        title_placeholder.text = title
        apply_theme_color(title_placeholder, theme_colors["title"])
    else:
        # If no title placeholder, create a text box for title
        left = Pt(36)
        top = Pt(36)
        width = Pt(648)
        height = Pt(50)
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_box.text_frame.text = title
        title_box.text_frame.paragraphs[0].font.size = Pt(32)
        title_box.text_frame.paragraphs[0].font.bold = True
        apply_theme_color(title_box, theme_colors["title"])
    
    # Add content if placeholder exists
    if content_placeholder:
        tf = content_placeholder.text_frame
    else:
        # If no content placeholder, create a text box for content
        left = Pt(36)
        top = Pt(108)
        width = Pt(648)
        height = Pt(432)
        content_box = slide.shapes.add_textbox(left, top, width, height)
        tf = content_box.text_frame
    
    # Format content
    tf.text = ""
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    # Add bullet points with proper formatting
    lines = content.strip().split('\n')
    for line in lines:
        if line.strip():
            p = tf.add_paragraph()
            p.text = line.strip()
            p.font.size = Pt(18)
            p.font.name = 'Calibri'
            p.level = 0
            apply_theme_color(p, theme_colors["accent"])
            p.bullet.enabled = True
    
    return slide

def create_section_slide(ppt, title, theme="professional"):
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["layouts"]["section"]]
    slide = ppt.slides.add_slide(layout)
    
    # Find title placeholder
    title_placeholder = None
    for shape in slide.placeholders:
        if shape.placeholder_format.type == 1:  # Title
            title_placeholder = shape
            break
    
    # Add title
    if title_placeholder:
        title_placeholder.text = title
        title_placeholder.text_frame.paragraphs[0].font.size = Pt(44)
        title_placeholder.text_frame.paragraphs[0].font.name = 'Calibri'
        apply_theme_color(title_placeholder.text_frame.paragraphs[0], get_theme_layout_ids(theme)["colors"]["title"])
    else:
        # If no title placeholder, create a text box for title
        left = Pt(36)  # 0.5 inch from left
        top = Pt(216)  # 3 inches from top (centered vertically)
        width = Pt(648)  # 9 inches
        height = Pt(72)  # 1 inch
        title_box = slide.shapes.add_textbox(left, top, width, height)
        tf = title_box.text_frame
        tf.text = title
        tf.paragraphs[0].font.size = Pt(44)
        tf.paragraphs[0].font.name = 'Calibri'
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        apply_theme_color(tf.paragraphs[0], get_theme_layout_ids(theme)["colors"]["title"])
    
    return slide

def generate_ppt(topic, num_slides=5, theme="professional"):
    # Clean the topic for file naming
    clean_topic = re.sub(r'[^\w\s-]', '', topic.replace('/', '_'))
    clean_topic = re.sub(r'[-\s]+', '_', clean_topic)
    
    # Create a unique filename
    filename = f"{clean_topic}_{os.urandom(3).hex()}.pptx"
    save_dir = "generated_presentations"
    os.makedirs(save_dir, exist_ok=True)
    
    ppt = Presentation()
    
    try:
        # Initialize the API client with GPT-3.5 Turbo
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable is not set")
        
        client = OpenAIClient(api_key, "gpt-3.5-turbo")
        
        # Generate outline with unique headings
        outline_prompt = f"""Create a presentation outline about {topic} with {num_slides} unique sections.
        Each section should have a distinct focus and heading (NO slide numbers).
        Format the content to be concise with 3-4 short bullet points per section.
        
        Example format for 'AI Impact on Business':
        - Digital Transformation Strategies
        - Customer Experience Revolution
        - Operational Efficiency Gains
        - Risk Management Evolution
        - Future Business Landscape
        
        Format as JSON with 'slides' array containing 'heading' and 'content' for each section.
        Keep bullet points under 60 characters each to prevent overflow."""
        
        outline_response = client.generate(outline_prompt)
        
        # Create title slide
        create_title_slide(ppt, topic, theme)
        
        # Create content slides based on the outline
        for i in range(num_slides):
            slide_prompt = f"""Create content for a presentation section about {topic}.
            Requirements:
            - Unique heading (NO slide numbers, NO topic repetition)
            - 3-4 bullet points
            - Each bullet point MUST be under 60 characters
            - Use active voice and concise language
            - Focus on specific aspects, not general overview
            
            Example format:
            Market Transformation Strategies
            - AI drives 40% efficiency gain in operations
            - Smart automation reduces costs by 25%
            - Customer satisfaction increased to 95%"""
            
            slide_content = client.generate(slide_prompt)
            create_content_slide(ppt, slide_content.split('\n')[0], '\n'.join(slide_content.split('\n')[1:]), theme)
        
        # Save the presentation
        output_path = os.path.join(save_dir, filename)
        ppt.save(output_path)
        logging.info(f"Presentation saved as {filename}")
        
        return filename
        
    except Exception as e:
        logging.error(f"Error generating presentation: {str(e)}")
        raise
