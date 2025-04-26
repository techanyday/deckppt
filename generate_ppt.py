import os
import re
import json
import random
import string
from pptx import Presentation
import logging
import tempfile

from apis.openai_api import OpenAIClient
from crawlers.icrawlercrawler import ICrawlerCrawler

def generate_ppt(topic, api_name, model_name, num_slides):
    # Clean the topic for file naming
    legal_topic = re.sub(r'[^\w\s-]', '', topic).strip().replace(' ', '_')
    
    # Create a temporary directory for processing
    temp_dir = tempfile.mkdtemp()
    timestamp = ''.join(random.choices(string.ascii_letters + string.digits, k=6))
    save_dir = os.path.join(temp_dir, f"{legal_topic}_{timestamp}")
    os.makedirs(save_dir, exist_ok=True)
    
    logging.info(f"Created temporary directory at: {save_dir}")

    # Copy the theme file to the temp directory
    theme_path = os.path.join(save_dir, "theme0.pptx")
    with open("theme0.pptx", "rb") as src, open(theme_path, "wb") as dst:
        dst.write(src.read())

    ppt = Presentation(theme_path)

    def delete_all_slides():
        for i in range(len(ppt.slides) - 1, -1, -1):
            r_id = ppt.slides._sldIdLst[i].rId
            ppt.part.drop_rel(r_id)
            del ppt.slides._sldIdLst[i]

    def create_title_slide(title, subtitle):
        layout = ppt.slide_layouts[0]
        slide = ppt.slides.add_slide(layout)
        slide.shapes.title.text = title
        slide.placeholders[1].text = subtitle

    def create_section_header_slide(title):
        layout = ppt.slide_layouts[2]
        slide = ppt.slides.add_slide(layout)
        slide.shapes.title.text = title

    def create_title_and_content_slide(title, content):
        layout = ppt.slide_layouts[1]
        slide = ppt.slides.add_slide(layout)
        slide.shapes.title.text = title
        slide.placeholders[1].text = content

    def create_title_and_content_and_image_slide(title, content, image_query):
        try:
            layout = ppt.slide_layouts[8]
            slide = ppt.slides.add_slide(layout)
            slide.shapes.title.text = title
            slide.placeholders[2].text = content
            
            # Try to get an image, but fallback gracefully if it fails
            try:
                crawler = ICrawlerCrawler(browser="google")
                image_name = crawler.get_image(image_query, save_dir)
                if image_name and os.path.isfile(os.path.join(save_dir, image_name)):
                    img_path = os.path.join(save_dir, image_name)
                    slide.shapes.add_picture(img_path, slide.placeholders[1].left, slide.placeholders[1].top,
                                          slide.placeholders[1].width, slide.placeholders[1].height)
                else:
                    logging.warning(f"Failed to download image for query: {image_query}")
            except Exception as e:
                logging.error(f"Error adding image to slide: {str(e)}")
                # Continue without the image
        except Exception as e:
            logging.error(f"Error creating slide: {str(e)}")
            # Fallback to a simple content slide
            layout = ppt.slide_layouts[1]
            slide = ppt.slides.add_slide(layout)
            slide.shapes.title.text = title
            slide.placeholders[1].text = content

    def find_text_in_between_tags(text, start_tag, end_tag):
        start_pos = text.find(start_tag)
        end_pos = text.find(end_tag)
        result = []
        while start_pos > -1 and end_pos > -1:
            text_between_tags = text[start_pos + len(start_tag):end_pos]
            result.append(text_between_tags)
            start_pos = text.find(start_tag, end_pos + len(end_tag))
            end_pos = text.find(end_tag, start_pos)
        res1 = "".join(result)
        res2 = re.sub(r"\[IMAGE\].*?\[/IMAGE\]", '', res1)
        if len(result) > 0:
            return res2
        else:
            return ""

    def search_for_slide_type(text):
        tags = ["[L_TS]", "[L_CS]", "[L_IS]", "[L_THS]"]
        found_text = next((s for s in tags if s in text), None)
        return found_text

    def parse_response(reply):
        try:
            list_of_slides = reply.split("[SLIDEBREAK]")
            for slide in list_of_slides:
                if not slide.strip():
                    continue
                    
                slide_type = search_for_slide_type(slide)
                if not slide_type:
                    continue

                title = find_text_in_between_tags(str(slide), "[TITLE]", "[/TITLE]")
                if not title:
                    continue

                if slide_type == "[L_TS]":
                    subtitle = find_text_in_between_tags(str(slide), "[SUBTITLE]", "[/SUBTITLE]")
                    create_title_slide(title, subtitle or "")
                elif slide_type == "[L_CS]":
                    content = find_text_in_between_tags(str(slide), "[CONTENT]", "[/CONTENT]")
                    create_title_and_content_slide(title, content or "")
                elif slide_type == "[L_IS]":
                    content = find_text_in_between_tags(str(slide), "[CONTENT]", "[/CONTENT]")
                    image_query = find_text_in_between_tags(str(slide), "[IMAGE]", "[/IMAGE]") or title
                    create_title_and_content_and_image_slide(title, content or "", image_query)
                elif slide_type == "[L_THS]":
                    create_section_header_slide(title)
        except Exception as e:
            logging.error(f"Error parsing presentation content: {str(e)}")
            raise

    def find_title():
        return ppt.slides[0].shapes.title.text

    # Initialize logging
    logging.basicConfig(level=logging.INFO)

    final_prompt = f"""Create an outline for a slideshow presentation on the topic of {topic} which is {num_slides}
        slides long. Make sure there are ONLY {num_slides} slides. This includes the title and thanks slide.

        You are allowed to use the following slide types:
        Title Slide - (Title, Subtitle)
        Content Slide - (Title, Content)
        Image Slide - (Title, Content, Image)
        Thanks Slide - (Title)
        
        Put this tag before the Title Slide: [L_TS]
        Put this tag before the Content Slide: [L_CS]
        Put this tag before the Image Slide: [L_IS]
        Put this tag before the Thanks Slide: [L_THS]
        
        Put this tag before the Title: [TITLE]
        Put this tag after the Title: [/TITLE]
        Put this tag before the Subtitle: [SUBTITLE]
        Put this tag after the Subtitle: [/SUBTITLE]
        Put this tag before the Content: [CONTENT]
        Put this tag after the Content: [/CONTENT]
        Put this tag before the Image: [IMAGE]
        Put this tag after the Image: [/IMAGE]

        Put "[SLIDEBREAK]" after each slide 

        For example:
        [L_TS]
        [TITLE]Among Us[/TITLE]

        [SLIDEBREAK]

        [L_CS] 
        [TITLE]What Is Among Us?[/TITLE]
        [CONTENT]
        1. Among Us is a popular online multiplayer game developed and published by InnerSloth.
        2. The game is set in a space-themed setting where players take on the roles of Crewmates and Impostors.
        3. The objective of Crewmates is to complete tasks and identify the Impostors among them, while the Impostors' goal is to sabotage the spaceship and eliminate the Crewmates without being caught.
        [/CONTENT]

        [SLIDEBREAK]


        Elaborate on the Content, provide as much information as possible.
        REMEMBER TO PLACE a [/CONTENT] at the end of the Content.
        Do not include any special characters (?, !, ., :, ) in the Title.
        Do not include any additional information in your response and stick to the format."""

    try:
        # Get API key from environment variable
        api_key = os.environ.get('OPENAI_API_KEY')
        if not api_key:
            raise ValueError("OPENAI_API_KEY environment variable is not set")
            
        api_client = OpenAIClient(api_key, model_name)
        presentation_content = api_client.generate(final_prompt)

        delete_all_slides()
        parse_response(presentation_content)

        # Generate unique filename
        filename = f"{legal_topic}_{timestamp}.pptx"
        output_file = os.path.join('generated_presentations', filename)
        
        # Save the presentation
        ppt.save(output_file)
        logging.info(f"Saved presentation to: {output_file}")
        
        return "Presentation generated successfully!", filename
    except Exception as e:
        logging.error(f"Error generating presentation: {str(e)}")
        raise
