import os
import requests
import logging
from urllib.parse import urljoin
from crawlers.base_crawler import BaseCrawler

class PixabayCrawler(BaseCrawler):
    def __init__(self):
        super().__init__("pixabay")
        self.api_key = os.environ.get('PIXABAY_API_KEY')
        if not self.api_key:
            raise ValueError("PIXABAY_API_KEY environment variable is not set")
        self.base_url = "https://pixabay.com/api/"

    def get_image(self, query, save_dir):
        """
        Search and download an image from Pixabay
        Args:
            query (str): Search term for the image
            save_dir (str): Directory to save the image
        Returns:
            str: Filename of the downloaded image or None if failed
        """
        try:
            # Make API request
            params = {
                'key': self.api_key,
                'q': query,
                'image_type': 'photo',
                'orientation': 'horizontal',
                'per_page': 3,  # Get top 3 results to choose from
                'safesearch': True,
            }
            
            response = requests.get(self.base_url, params=params)
            response.raise_for_status()
            data = response.json()
            
            if not data.get('hits'):
                logging.warning(f"No images found for query: {query}")
                return None
            
            # Get the first image URL
            image_url = data['hits'][0].get('largeImageURL')
            if not image_url:
                logging.warning(f"No large image URL found for query: {query}")
                return None
            
            # Download the image
            image_response = requests.get(image_url)
            image_response.raise_for_status()
            
            # Create a unique filename
            image_ext = image_url.split('.')[-1]
            if image_ext not in ['jpg', 'jpeg', 'png']:
                image_ext = 'jpg'
            
            filename = f"pixabay_{query.replace(' ', '_')}_{data['hits'][0]['id']}.{image_ext}"
            filepath = os.path.join(save_dir, filename)
            
            # Save the image
            with open(filepath, 'wb') as f:
                f.write(image_response.content)
            
            logging.info(f"Successfully downloaded image: {filename}")
            return filename
            
        except requests.RequestException as e:
            logging.error(f"Error making request to Pixabay API: {str(e)}")
            return None
        except Exception as e:
            logging.error(f"Error downloading image from Pixabay: {str(e)}")
            return None
