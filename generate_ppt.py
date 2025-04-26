import os
import logging
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import MSO_AUTO_SIZE
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
    
    # Set title with appropriate formatting
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(44)
    title_shape.text_frame.paragraphs[0].font.name = 'Calibri'
    
    # Add subtitle if placeholder exists
    if len(slide.placeholders) > 1:
        subtitle = slide.placeholders[1]
        subtitle.text = "Generated with AI"
        subtitle.text_frame.paragraphs[0].font.size = Pt(24)
        subtitle.text_frame.paragraphs[0].font.name = 'Calibri'
    
    return slide

def create_content_slide(ppt, title, content, theme="professional"):
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["content"]]
    slide = ppt.slides.add_slide(layout)
    
    # Set title
    title_shape = slide.shapes.title
    title_shape.text = title
    
    # Format content shape
    content_shape = slide.placeholders[1]
    tf = content_shape.text_frame
    tf.text = ""  # Clear any existing text
    
    # Add content with proper formatting
    p = tf.paragraphs[0]
    p.text = content
    p.font.size = Pt(18)  # Adjust font size for better fit
    p.font.name = 'Calibri'
    
    # Auto-fit text
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    return slide

def create_section_slide(ppt, title, theme="professional"):
    layout = ppt.slide_layouts[get_theme_layout_ids(theme)["section"]]
    slide = ppt.slides.add_slide(layout)
    
    # Set title with larger font
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(44)
    title_shape.text_frame.paragraphs[0].font.name = 'Calibri'
    
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
