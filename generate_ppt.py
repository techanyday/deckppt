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
from io import BytesIO
from datetime import datetime

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
    """Create a content slide with title, image, and bullet points in that order."""
    # Limit bullet points per slide
    MAX_POINTS_PER_SLIDE = 5
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
        
        # Calculate consistent width for content
        content_width = Pt(576)  # 8 inches wide
        content_left = Pt(72)  # 1 inch from left margin
        
        # 1. Add title at the top
        title_placeholder = slide.shapes.title
        if title_placeholder:
            # Remove slide numbers from title
            slide_title = title.split(" (")[0] if " (" in title else title
            title_placeholder.text = slide_title
            title_placeholder.text_frame.paragraphs[0].font.size = Pt(40)
            title_placeholder.text_frame.paragraphs[0].font.name = 'Calibri'
            title_placeholder.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            title_placeholder.text_frame.word_wrap = True
            title_placeholder.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
            # Position title
            title_placeholder.top = Pt(24)  # 0.33 inch from top
            title_placeholder.left = content_left
            title_placeholder.width = content_width
            title_placeholder.height = Pt(60)
            apply_theme_color(title_placeholder.text_frame.paragraphs[0], get_theme_layout_ids(theme)["colors"]["title"])

        # 2. Add image below title
        image_height = Pt(216)  # 3 inches
        if image_data and slide_num == 0:
            try:
                # Create a temporary file for the image
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp:
                    if isinstance(image_data, BytesIO):
                        tmp.write(image_data.getvalue())
                    else:
                        tmp.write(image_data)
                    tmp.flush()
                    
                    # Use same width as content area
                    img_width = content_width
                    img_height = image_height
                    img_left = content_left  # Align with bullet points
                    img_top = Pt(96)  # Below title
                    
                    slide.shapes.add_picture(tmp.name, img_left, img_top, img_width, img_height)
                    
                os.unlink(tmp.name)  # Clean up temp file
                
            except Exception as e:
                logging.error(f"Error adding image to slide: {e}")
                image_height = Pt(0)  # If image fails, don't reserve space for it

        # 3. Add bullet points below image
        content_top = Pt(96) + image_height + Pt(24)  # Below image
        content_height = Pt(396) - image_height  # Remaining space

        # Create content text box
        content_placeholder = slide.shapes.add_textbox(
            content_left, content_top, content_width, content_height
        )

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

def generate_intro_slide(ppt, presentation_title, theme="professional"):
    """Generate an introduction slide with overview of the topic."""
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["layouts"]["title"]]
    slide = ppt.slides.add_slide(layout)
    
    # Apply background color
    apply_slide_background(slide, get_theme_layout_ids(theme)["colors"]["background"])
    
    # Add title
    title_placeholder = slide.shapes.title
    if title_placeholder:
        overview_title = f"{presentation_title} Overview"
        title_placeholder.text = overview_title
        title_placeholder.text_frame.paragraphs[0].font.size = Pt(44)
        title_placeholder.text_frame.paragraphs[0].font.name = 'Calibri'
        title_placeholder.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        apply_theme_color(title_placeholder.text_frame.paragraphs[0], get_theme_layout_ids(theme)["colors"]["title"])
    
    # Add overview text
    content_left = Pt(72)  # 1 inch from left
    content_top = Pt(216)  # 3 inches from top
    content_width = Pt(576)  # 8 inches wide
    content_height = Pt(216)  # 3 inches tall
    
    content = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
    tf = content.text_frame
    tf.word_wrap = True
    
    # Generate overview text using OpenAI
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError("OPENAI_API_KEY environment variable is not set")
        
    client = OpenAIClient(api_key)
    prompt = f"""Generate a brief 2-3 sentence overview for a presentation titled '{presentation_title}'.
    The overview should be professional, engaging, and set up the context for the presentation.
    Do not use phrases like 'In this presentation' or 'We will discuss'.
    Instead, make direct statements about the topic."""
    
    try:
        response = client.generate(prompt)
        overview_text = response.strip()
    except Exception as e:
        logging.error(f"Error generating overview: {e}")
        overview_text = f"Discover key insights and strategies about {presentation_title.lower()}. This presentation explores proven approaches and practical solutions for success in this domain."
    
    p = tf.add_paragraph()
    p.text = overview_text
    p.font.size = Pt(28)
    p.font.name = 'Calibri'
    p.alignment = PP_ALIGN.CENTER
    apply_theme_color(p, get_theme_layout_ids(theme)["colors"]["title"])
    
    return slide

def generate_conclusion_slide(ppt, key_points, theme="professional"):
    """Generate a conclusion slide with key takeaways."""
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["layouts"]["content"]]
    slide = ppt.slides.add_slide(layout)
    
    # Apply background color
    apply_slide_background(slide, get_theme_layout_ids(theme)["colors"]["background"])
    
    # Add title
    title_placeholder = slide.shapes.title
    if title_placeholder:
        title_placeholder.text = "Key Takeaways"
        title_placeholder.text_frame.paragraphs[0].font.size = Pt(44)
        title_placeholder.text_frame.paragraphs[0].font.name = 'Calibri'
        title_placeholder.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        apply_theme_color(title_placeholder.text_frame.paragraphs[0], get_theme_layout_ids(theme)["colors"]["title"])
    
    # Add key takeaways
    content_left = Pt(72)  # 1 inch from left
    content_top = Pt(144)  # 2 inches from top
    content_width = Pt(576)  # 8 inches wide
    content_height = Pt(360)  # 5 inches tall
    
    content = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
    tf = content.text_frame
    tf.word_wrap = True
    
    # Add takeaway points
    for point in key_points[:5]:  # Limit to 5 points
        p = tf.add_paragraph()
        p.text = point
        p.font.size = Pt(28)
        p.font.name = 'Calibri'
        p.alignment = PP_ALIGN.LEFT
        p.level = 0  # Top level bullet
        apply_theme_color(p, get_theme_layout_ids(theme)["colors"]["title"])
        p.space_after = Pt(20)
        p.space_before = Pt(8)
    
    return slide

def generate_slide_overview(title):
    """Generate a brief overview of the presentation topic."""
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError("OPENAI_API_KEY environment variable is not set")
        
    client = OpenAIClient(api_key)
    prompt = f"""Generate a brief 2-3 sentence overview for a presentation titled '{title}'.
    The overview should be professional, engaging, and set up the context for the presentation.
    Do not use phrases like 'In this presentation' or 'We will discuss'.
    Instead, make direct statements about the topic."""
    
    try:
        response = client.generate(prompt)
        return response.strip()
    except Exception as e:
        logging.error(f"Error generating overview: {e}")
        return f"Discover key insights and strategies about {title.lower()}. This presentation explores proven approaches and practical solutions for success in this domain."

def generate_slide_title(content, base_title):
    """Generate a unique and relevant title based on slide content."""
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError("OPENAI_API_KEY environment variable is not set")
        
    client = OpenAIClient(api_key)
    prompt = f"""Generate a short, professional slide title (4-6 words) based on this content:
    Base title: {base_title}
    Content points: {content}
    The title should be specific to the content but not too long.
    Do not use generic words like 'Overview' or 'Introduction'.
    Do not use punctuation in the title."""
    
    try:
        response = client.generate(prompt)
        return response.strip()
    except Exception as e:
        logging.error(f"Error generating slide title: {e}")
        return base_title

def generate_ppt(topic, num_slides=5, theme="professional"):
    # Clean the topic for file naming
    clean_topic = re.sub(r'[^\w\s-]', '', topic.replace('/', '_'))
    clean_topic = re.sub(r'\s+', '_', clean_topic.strip())
    
    # Initialize OpenAI client
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError("OPENAI_API_KEY environment variable is not set")
        
    client = OpenAIClient(api_key)
    
    try:
        # Generate content for all slides
        sections = []
        remaining_slides = num_slides - 2  # Account for intro and conclusion slides
        
        # Generate main content points
        content_prompt = f"""Generate {remaining_slides} distinct sections for a presentation about '{topic}'.
        For each section:
        1. Focus on a unique aspect or subtopic
        2. Include 3-4 clear, concise bullet points
        3. Each bullet point should be a complete, actionable statement
        4. Avoid repetition between sections
        5. Ensure natural progression of ideas
        
        Format each section's bullet points as a list."""
        
        content_response = client.generate(content_prompt)
        content_sections = content_response.strip().split("\n\n")
        
        # Process each section
        for section_text in content_sections[:remaining_slides]:
            # Extract bullet points
            points = [p.strip().strip('*-').strip() for p in section_text.split("\n") 
                     if p.strip() and not p.strip().isdigit()]
            
            # Generate image prompt for this section
            image_prompt = f"""Based on these points about {topic}:
            {points}
            Generate a creative prompt for an image that would complement this content.
            The image should be professional and business-appropriate.
            Focus on abstract concepts, data visualization, or business scenarios."""
            
            try:
                image_prompt_response = client.generate(image_prompt)
                image_data = client.generate_image(image_prompt_response)
            except Exception as e:
                logging.error(f"Error generating image: {e}")
                image_data = None
            
            sections.append({
                "points": points,
                "image_data": image_data
            })
        
        # Create the complete presentation
        ppt = create_presentation(topic, sections)
        
        # Save the presentation
        timestamp = datetime.now().strftime("%H%M%S")
        filename = f"{clean_topic}_{timestamp[:6]}.pptx"
        ppt.save(filename)
        
        return filename
        
    except Exception as e:
        logging.error(f"Error generating presentation: {e}")
        raise

def create_presentation(title, sections):
    """Create a complete presentation with proper structure."""
    ppt = Presentation()
    
    # 1. Generate intro slide
    generate_intro_slide(ppt, title)
    
    # 2. Generate content slides with unique titles
    all_points = []  # Collect points for conclusion
    for section in sections:
        section_points = section.get("points", [])
        all_points.extend(section_points)
        
        # Generate unique title for this section
        section_title = generate_slide_title(section_points, title)
        
        # Create content slides
        create_content_slide(ppt, section_title, section_points, 
                           image_data=section.get("image_data"))
    
    # 3. Generate conclusion slide
    # Extract key takeaways from all points
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError("OPENAI_API_KEY environment variable is not set")
        
    client = OpenAIClient(api_key)
    takeaways_prompt = f"""Based on these presentation points, generate 3-5 key takeaways:
    {all_points}
    Each takeaway should be a complete, actionable statement.
    Focus on the most important insights and practical implications."""
    
    try:
        takeaways = client.generate(takeaways_prompt).strip().split("\n")
    except Exception as e:
        logging.error(f"Error generating takeaways: {e}")
        # Create basic takeaways from the first point of each section
        takeaways = [p[0] for p in [s.get("points", []) for s in sections] if p][:5]
    
    generate_conclusion_slide(ppt, takeaways)
    
    return ppt
