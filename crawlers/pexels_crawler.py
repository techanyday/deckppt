import os
import requests
import logging
import re
from urllib.parse import urljoin
from crawlers.base_crawler import BaseCrawler

class PexelsCrawler(BaseCrawler):
    def __init__(self):
        super().__init__("pexels")
        self.api_key = os.environ.get('PEXELS_API_KEY')
        if not self.api_key:
            raise ValueError("PEXELS_API_KEY environment variable is not set")
        self.base_url = "https://api.pexels.com/v1/search"
        self.headers = {
            'Authorization': self.api_key
        }

    def _sanitize_filename(self, filename):
        """Sanitize filename to be safe for all operating systems"""
        # Replace any non-alphanumeric characters (except dashes and underscores) with underscore
        filename = re.sub(r'[^\w\-_]', '_', filename)
        # Remove multiple consecutive underscores
        filename = re.sub(r'_+', '_', filename)
        # Remove leading/trailing underscores
        filename = filename.strip('_')
        return filename

    def get_image(self, query, save_dir):
        """
        Search and download an image from Pexels
        Args:
            query (str): Search term for the image
            save_dir (str): Directory to save the image
        Returns:
            str: Filename of the downloaded image or None if failed
        """
        try:
            # Make sure save_dir exists
            os.makedirs(save_dir, exist_ok=True)

            # Make API request
            params = {
                'query': query,
                'per_page': 1,  # Get just one result
                'orientation': 'landscape'  # Better for presentations
            }
            
            response = requests.get(
                self.base_url, 
                headers=self.headers,
                params=params
            )
            response.raise_for_status()
            data = response.json()
            
            if not data.get('photos'):
                logging.warning(f"No images found for query: {query}")
                return None
            
            # Get the image URL (large size)
            photo = data['photos'][0]
            image_url = photo['src']['large']
            
            # Download the image
            image_response = requests.get(image_url)
            image_response.raise_for_status()
            
            # Create filename with photo ID for uniqueness
            safe_query = self._sanitize_filename(query)
            filename = f"pexels_{safe_query}_{photo['id']}.jpg"
            filepath = os.path.join(save_dir, filename)
            
            # Save the image
            with open(filepath, 'wb') as f:
                f.write(image_response.content)
            
            logging.info(f"Successfully downloaded image: {filename}")
            return filename
            
        except requests.RequestException as e:
            logging.error(f"Error making request to Pexels API: {str(e)}")
            return None
        except Exception as e:
            logging.error(f"Error downloading image from Pexels: {str(e)}")
            return None
