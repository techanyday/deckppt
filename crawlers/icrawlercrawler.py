import os
import string
from random import choice
from urllib.parse import urlparse
from icrawler.builtin import ImageDownloader, BaiduImageCrawler, BingImageCrawler, GoogleImageCrawler
from crawlers import base_crawler
import logging

class ICrawlerDownloader(ImageDownloader):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.final_image_name = ""
        self.unique_image_name = ""
        self.download_success = False

    def get_filename(self, task, default_ext):
        self.final_image_name = "image_" + self.unique_image_name + "." + default_ext
        return self.final_image_name

    def download(self, task, default_ext, timeout=5, max_retry=3, **kwargs):
        try:
            success = super().download(task, default_ext, timeout, max_retry, **kwargs)
            self.download_success = success
            return success
        except Exception as e:
            logging.error(f"Failed to download image: {str(e)}")
            self.download_success = False
            return False

    def generate_new_name(self):
        self.unique_image_name = ''.join(choice(string.ascii_uppercase + string.ascii_lowercase + string.digits) for _
                                         in range(16))

    def get_image_name(self):
        return self.final_image_name if self.download_success else None


class ICrawlerCrawler(base_crawler.BaseCrawler):
    def __init__(self, browser):
        super().__init__(browser)
        self.browser = browser

    def get_image(self, query, save_dir):
        try:
            if self.browser == "google":
                crawler = GoogleImageCrawler(downloader_cls=ICrawlerDownloader, storage={'root_dir': save_dir if save_dir
                else os.getcwd()})
            elif self.browser == "bing":
                crawler = BingImageCrawler(downloader_cls=ICrawlerDownloader, storage={'root_dir': save_dir if save_dir
                else os.getcwd()})
            elif self.browser == "baidu":
                crawler = BaiduImageCrawler(downloader_cls=ICrawlerDownloader, storage={'root_dir': save_dir if save_dir
                else os.getcwd()})
            else:
                logging.error(f"Unsupported browser: {self.browser}")
                return None

            downloader = crawler.downloader
            downloader.generate_new_name()
            
            try:
                crawler.crawl(keyword=query, max_num=1)
            except Exception as e:
                logging.error(f"Failed to crawl for image: {str(e)}")
                return None

            final_image_name = downloader.get_image_name()
            if not final_image_name:
                logging.warning(f"Failed to download image for query: {query}")
                return None

            # Verify the file exists
            image_path = os.path.join(save_dir, final_image_name)
            if not os.path.exists(image_path):
                logging.warning(f"Image file not found after download: {image_path}")
                return None

            return final_image_name
        except Exception as e:
            logging.error(f"Error in get_image: {str(e)}")
            return None
