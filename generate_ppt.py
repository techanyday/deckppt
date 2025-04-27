import os
import logging
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from apis.openai_api import OpenAIClient
import re
import tempfile

def get_theme_layout_ids(theme):
    """Get layout IDs and theme colors for different themes"""
    themes = {
        "professional": {
            "layouts": {
                "title": 0,
                "content": 1,  # Using Title and Content layout
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
                "content": 1,  # Using Title and Content layout
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
                "content": 1,  # Using Title and Content layout
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
                "content": 1,  # Using Title and Content layout
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
                "content": 1,  # Using Title and Content layout
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
    """Create the title slide with background color and title text."""
    # Get the title slide layout (first layout)
    layout = ppt.slide_layouts[0]
    slide = ppt.slides.add_slide(layout)
    theme_colors = get_theme_layout_ids(theme)["colors"]
    
    # Apply background color first
    apply_slide_background(slide, theme_colors["background"])
    
    # Get the title placeholder
    title_placeholder = slide.shapes.title
    
    # Set title text and properties
    title_placeholder.text = title
    title_frame = title_placeholder.text_frame
    title_frame.paragraphs[0].font.size = Pt(44)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    title_frame.word_wrap = True
    title_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    apply_theme_color(title_frame.paragraphs[0], theme_colors["title"])
    
    # Remove the subtitle placeholder if it exists
    for shape in slide.placeholders:
        if shape.placeholder_format.type == 2:  # Subtitle placeholder
            sp = shape._element
            sp.getparent().remove(sp)
    
    return slide

def create_content_slide(ppt, title, bullet_points, theme="professional", image_data=None):
    """Create a content slide with title, bullet points and optional image."""
    # Limit bullet points per slide
    MAX_POINTS_PER_SLIDE = 4
    slides_needed = (len(bullet_points) + MAX_POINTS_PER_SLIDE - 1) // MAX_POINTS_PER_SLIDE
    all_slides = []

    for slide_num in range(slides_needed):
        # Get points for this slide
        start_idx = slide_num * MAX_POINTS_PER_SLIDE
        end_idx = min(start_idx + MAX_POINTS_PER_SLIDE, len(bullet_points))
        current_points = bullet_points[start_idx:end_idx]

        # Create slide with current set of points
        layout = ppt.slide_layouts[get_theme_layout_ids(theme)["layouts"]["content"]]
        slide = ppt.slides.add_slide(layout)
        all_slides.append(slide)
        
        # Apply background color
        apply_slide_background(slide, get_theme_layout_ids(theme)["colors"]["background"])
        
        # Add title (add part number if multiple slides)
        title_placeholder = slide.shapes.title
        if title_placeholder:
            slide_title = title if slides_needed == 1 else f"{title} ({slide_num + 1}/{slides_needed})"
            title_placeholder.text = slide_title
            title_placeholder.text_frame.paragraphs[0].font.size = Pt(40)
            title_placeholder.text_frame.paragraphs[0].font.name = 'Calibri'
            title_placeholder.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            title_placeholder.text_frame.word_wrap = True
            title_placeholder.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
            apply_theme_color(title_placeholder.text_frame.paragraphs[0], get_theme_layout_ids(theme)["colors"]["title"])

        # Add content
        content_placeholder = None
        for shape in slide.placeholders:
            if shape.placeholder_format.type == 1:  # Content placeholder
                content_placeholder = shape
                break
        
        if not content_placeholder:
            # If no content placeholder, create a text box
            left = Pt(36)  # 0.5 inch from left
            top = Pt(144)  # 2 inches from top
            width = Pt(648)  # 9 inches
            height = Pt(324)  # 4.5 inches
            content_placeholder = slide.shapes.add_textbox(left, top, width, height)

        # Configure text frame for bullet points
        tf = content_placeholder.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        tf.clear()  # Clear existing content

        # Add bullet points with proper formatting
        for point in current_points:
            p = tf.add_paragraph()
            p.text = point
            p.font.size = Pt(24)
            p.font.name = 'Calibri'
            p.alignment = PP_ALIGN.LEFT
            p.level = 0  # Top level bullet
            apply_theme_color(p, get_theme_layout_ids(theme)["colors"]["title"])
            
            # Add spacing between bullet points
            p.space_after = Pt(12)
            p.space_before = Pt(6)

        # Add image only to the first slide if provided
        if image_data and slide_num == 0:
            try:
                # Create a temporary file for the image
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                    tmp.write(image_data)
                    tmp.flush()
                    
                    # Calculate image position and size
                    # Make image smaller and position it on the right
                    img_left = Pt(400)  # Move image to the right
                    img_top = Pt(144)
                    img_width = Pt(284)  # Make image narrower
                    img_height = Pt(213)  # Maintain aspect ratio
                    
                    slide.shapes.add_picture(tmp.name, img_left, img_top, img_width, img_height)
                    
                    # Adjust content width to make room for image
                    if hasattr(content_placeholder, 'width'):
                        content_placeholder.width = Pt(360)  # Make text area narrower
                    
                os.unlink(tmp.name)  # Clean up temp file
                
            except Exception as e:
                logging.error(f"Error adding image to slide: {e}")
    
    return all_slides[0]  # Return the first slide for compatibility

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
        # Initialize the API client
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable is not set")
        
        client = OpenAIClient(api_key, "gpt-3.5-turbo")
        
        # Create title slide
        create_title_slide(ppt, topic, theme)
        
        # Create content slides based on the outline
        for i in range(num_slides):
            # Generate slide content
            slide_prompt = f"""Create content for a presentation section about {topic}.
            Requirements:
            - Unique heading (NO slide numbers, NO topic repetition)
            - 3-4 bullet points ONLY
            - Each bullet point MUST be under 60 characters
            - Do NOT use any bullet symbols or dashes
            - Each bullet point should be on a new line
            - Use active voice and concise language
            - Focus on specific aspects, not general overview
            
            Example format:
            Market Transformation Strategies
            Automation reduces operational costs significantly
            Customer satisfaction reaches record levels
            Digital solutions drive efficiency improvements"""
            
            slide_content = client.generate(slide_prompt)
            
            # Split content into title and bullet points
            content_lines = [line.strip() for line in slide_content.split('\n') if line.strip()]
            if not content_lines:
                continue
                
            title = content_lines[0]
            bullet_points = content_lines[1:]  # Get bullet points without any symbols
            
            # Generate image for the slide
            image_prompt = f"""Create a professional presentation image for the topic: {title}
            Requirements:
            - Modern and minimalist style
            - Professional business context
            - No text or words in the image
            - Suitable for corporate presentations
            - Clean and simple design
            - High contrast and visibility"""
            
            try:
                image_data = client.generate_image(image_prompt)
                create_content_slide(ppt, title, bullet_points, theme, image_data)
            except Exception as e:
                logging.error(f"Error generating image: {e}")
                create_content_slide(ppt, title, bullet_points, theme)
        
        # Save the presentation
        output_path = os.path.join(save_dir, filename)
        ppt.save(output_path)
        logging.info(f"Presentation saved as {filename}")
        
        return filename
        
    except Exception as e:
        logging.error(f"Error generating presentation: {str(e)}")
        raise
