import openai
import logging
import json
import base64
import requests
from io import BytesIO
from PIL import Image

class OpenAIClient:
    def __init__(self, api_key, model="gpt-3.5-turbo"):
        self.api_key = api_key
        self.model = model
        openai.api_key = api_key

    def generate(self, prompt):
        try:
            response = openai.ChatCompletion.create(
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
        """Generate an image using DALL-E 3"""
        try:
            response = openai.Image.create(
                model="dall-e-3",
                prompt=prompt,
                size=size,
                quality="standard",
                n=1
            )
            
            # Download the image
            image_url = response.data[0].url
            image_response = requests.get(image_url)
            if image_response.status_code == 200:
                return BytesIO(image_response.content)
            else:
                raise Exception(f"Failed to download image: {image_response.status_code}")
                
        except Exception as e:
            logging.error(f"Error generating image: {str(e)}")
            raise
