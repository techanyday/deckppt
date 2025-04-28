from openai import OpenAI
import logging
import json
import requests
from io import BytesIO
from PIL import Image

class OpenAIClient:
    def __init__(self, api_key, model="gpt-3.5-turbo"):
        self.client = OpenAI(api_key=api_key)
        self.model = model

    def generate(self, prompt):
        """Generate text using the OpenAI API"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.7,
                max_tokens=500
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            logging.error(f"Error generating text: {str(e)}")
            raise

    def generate_image(self, prompt, size="1024x1024"):
        """Generate an image using DALL-E 2"""
        try:
            # Log the image generation attempt
            logging.info(f"[OpenAI] Generating image with prompt: {prompt}")

            response = self.client.images.generate(
                model="dall-e-2",  # Use DALL-E 2 which is more cost-effective
                prompt=prompt,
                size=size,
                n=1
            )
            
            # Log the response
            logging.info(f"[OpenAI] Image generation response: {response}")
            
            # Download the image
            image_url = response.data[0].url
            logging.info(f"[OpenAI] Downloading image from: {image_url}")
            
            image_response = requests.get(image_url)
            if image_response.status_code == 200:
                logging.info("[OpenAI] Image downloaded successfully")
                return image_response.content  # Return bytes directly instead of BytesIO
            else:
                error_msg = f"Failed to download image: {image_response.status_code}"
                logging.error(f"[OpenAI] {error_msg}")
                raise Exception(error_msg)
                
        except Exception as e:
            logging.error(f"[OpenAI] Error generating image: {str(e)}")
            # Return None instead of raising to prevent presentation generation failure
            return None
