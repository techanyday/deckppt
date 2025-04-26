import os
import logging
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from apis.openai_api import OpenAIClient
import re
import tempfile

def get_theme_layout_ids(theme):
    """Get layout IDs for different themes"""
    themes = {
        "professional": {
            "title": 0,  # Title Slide
            "content": 1,  # Title and Content
            "section": 2,  # Section Header
            "two_content": 3,  # Two Content
            "comparison": 4  # Comparison
        },
        "modern": {
            "title": 5,  # Modern Title Slide
            "content": 6,  # Modern Content
            "section": 7,  # Modern Section
            "two_content": 8,  # Modern Two Content
            "comparison": 9  # Modern Comparison
        },
        "minimal": {
            "title": 0,  # Simple Title
            "content": 1,  # Simple Content
            "section": 2,  # Simple Section
            "two_content": 3,  # Simple Two Content
            "comparison": 4  # Simple Comparison
        },
        "creative": {
            "title": 5,  # Creative Title
            "content": 6,  # Creative Content
            "section": 7,  # Creative Section
            "two_content": 8,  # Creative Two Content
            "comparison": 9  # Creative Comparison
        },
        "corporate": {
            "title": 0,  # Corporate Title
            "content": 1,  # Corporate Content
            "section": 2,  # Corporate Section
            "two_content": 3,  # Corporate Two Content
            "comparison": 4  # Corporate Comparison
        }
    }
    return themes.get(theme, themes["professional"])

def create_title_slide(ppt, title, theme="professional"):
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["title"]]
    slide = ppt.slides.add_slide(layout)
    
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
    else:
        # If no title placeholder, create a text box for title
        left = Pt(36)  # 0.5 inch from left
        top = Pt(180)  # 2.5 inches from top
        width = Pt(648)  # 9 inches
        height = Pt(72)  # 1 inch
        title_box = slide.shapes.add_textbox(left, top, width, height)
        tf = title_box.text_frame
        tf.text = title
        tf.paragraphs[0].font.size = Pt(44)
        tf.paragraphs[0].font.name = 'Calibri'
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Add subtitle
    if subtitle_placeholder:
        subtitle_placeholder.text = "Generated with AI"
        subtitle_placeholder.text_frame.paragraphs[0].font.size = Pt(24)
        subtitle_placeholder.text_frame.paragraphs[0].font.name = 'Calibri'
    else:
        # If no subtitle placeholder, create a text box for subtitle
        left = Pt(36)  # 0.5 inch from left
        top = Pt(288)  # 4 inches from top
        width = Pt(648)  # 9 inches
        height = Pt(36)  # 0.5 inch
        subtitle_box = slide.shapes.add_textbox(left, top, width, height)
        tf = subtitle_box.text_frame
        tf.text = "Generated with AI"
        tf.paragraphs[0].font.size = Pt(24)
        tf.paragraphs[0].font.name = 'Calibri'
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    return slide

def create_content_slide(ppt, title, content, theme="professional"):
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["content"]]
    slide = ppt.slides.add_slide(layout)
    
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
    else:
        # If no title placeholder, create a text box for title
        left = Pt(36)  # 0.5 inch from left
        top = Pt(36)   # 0.5 inch from top
        width = Pt(648)  # 9 inches
        height = Pt(50)  # About 0.7 inch
        title_box = slide.shapes.add_textbox(left, top, width, height)
        title_box.text_frame.text = title
        title_box.text_frame.paragraphs[0].font.size = Pt(32)
        title_box.text_frame.paragraphs[0].font.bold = True
    
    # Add content if placeholder exists
    if content_placeholder:
        tf = content_placeholder.text_frame
    else:
        # If no content placeholder, create a text box for content
        left = Pt(36)  # 0.5 inch from left
        top = Pt(108)  # 1.5 inches from top
        width = Pt(648)  # 9 inches
        height = Pt(432)  # 6 inches
        content_box = slide.shapes.add_textbox(left, top, width, height)
        tf = content_box.text_frame
    
    # Format content
    tf.text = ""  # Clear any existing text
    p = tf.paragraphs[0]
    p.text = content
    p.font.size = Pt(18)
    p.font.name = 'Calibri'
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    return slide

def create_section_slide(ppt, title, theme="professional"):
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["section"]]
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
        
        # Generate outline
        outline_prompt = f"""Create a {num_slides}-slide presentation outline about {topic}. 
        For each slide, provide a clear heading and key points. Format the content to be concise and fit within a standard presentation slide.
        Each bullet point should be brief and not exceed 2 lines.
        Do not include 'Title:', 'Slide X:', or any slide numbers in the content.
        Format as JSON with 'slides' array containing 'heading' and 'content' for each slide."""
        outline_response = client.generate(outline_prompt)
        
        # Create title slide
        create_title_slide(ppt, topic, theme)
        
        # Create content slides based on the outline
        for i in range(num_slides):
            slide_prompt = f"""For a presentation about {topic}, generate concise content that will fit on a single slide.
            The content should be brief and well-formatted with:
            - A clear heading (without any 'Title:' prefix or slide numbers)
            - 3-4 key bullet points
            - Each bullet point should be 1-2 lines maximum
            - Total content should fit on a standard presentation slide without overflow"""
            slide_content = client.generate(slide_prompt)
            
            # Create different types of slides based on content
            if i == 0:
                create_section_slide(ppt, "Overview", theme)
            create_content_slide(ppt, slide_content.split('\n')[0], '\n'.join(slide_content.split('\n')[1:]), theme)
        
        # Save the presentation
        output_path = os.path.join(save_dir, filename)
        ppt.save(output_path)
        logging.info(f"Presentation saved as {filename}")
        
        return filename
        
    except Exception as e:
        logging.error(f"Error generating presentation: {str(e)}")
        raise
