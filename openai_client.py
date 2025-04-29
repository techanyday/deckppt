import os
from openai import OpenAI
import logging

logger = logging.getLogger(__name__)

class OpenAIClient:
    def __init__(self):
        self.client = OpenAI(api_key=os.environ.get('OPENAI_API_KEY'))
        self.model = "gpt-3.5-turbo"  # Can be configured as needed
        
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
            logger.error(f"Error generating text: {str(e)}")
            raise
