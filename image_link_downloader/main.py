import pandas as pd
import requests
import os
import time
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
from tqdm import tqdm
import argparse

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class ExcelImageDownloader:
    def __init__(self, excel_path, url_column='url', output_dir='downloaded_images',
                 max_workers=5, timeout=30, headers=None):
        """
        Initialize the image downloader
        
        Args:
            excel_path (str): Path to the Excel file
            url_column (str): Column name containing image URLs
            output_dir (str): Directory to save downloaded images
            max_workers (int): Number of parallel downloads
            timeout (int): Request timeout in seconds
            headers (dict): Custom headers for requests
        """
        self.excel_path = excel_path
        self.url_column = url_column
        self.output_dir = output_dir
        self.max_workers = max_workers
        self.timeout = timeout
        self.headers = headers or {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        # Create output directory if it doesn't exist
        os.makedirs(self.output_dir, exist_ok=True)
        
    def load_excel_data(self):
        """Load data from Excel file"""
        try:
            # Try different Excel readers
            if self.excel_path.endswith('.xlsx'):
                df = pd.read_excel(self.excel_path, engine='openpyxl')
            elif self.excel_path.endswith('.xls'):
                df = pd.read_excel(self.excel_path, engine='xlrd')
            else:
                raise ValueError("Unsupported file format. Use .xlsx or .xls")
            
            logger.info(f"Loaded Excel file with {len(df)} rows")
            
            # Check if URL column exists
            if self.url_column not in df.columns:
                available_columns = ', '.join(df.columns)
                raise ValueError(
                    f"Column '{self.url_column}' not found in Excel. "
                    f"Available columns: {available_columns}"
                )
            
            # Filter out empty URLs
            urls = df[self.url_column].dropna().astype(str).tolist()
            logger.info(f"Found {len(urls)} valid URLs")
            
            return df, urls
            
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            raise
    
    def get_filename(self, url, index, df_row=None):
        """Generate filename for the image"""
        try:
            # Try to extract filename from URL
            parsed_url = urlparse(url)
            filename = os.path.basename(parsed_url.path)
            
            # If no filename in URL, use index
            if not filename or '.' not in filename:
                filename = f"image_{index}"
            
            # Ensure the file has an extension
            if '.' not in filename:
                # Try to get extension from content type or use .jpg as default
                filename = f"{filename}.jpg"
            
            # Clean filename (remove query parameters, etc.)
            filename = filename.split('?')[0]
            
            # If we have dataframe row data, try to create meaningful name
            if df_row is not None:
                # Try to use an 'id' or 'name' column if available
                for col in ['id', 'name', 'title', 'product_id']:
                    if col in df_row and pd.notna(df_row[col]):
                        safe_name = str(df_row[col]).replace('/', '_').replace('\\', '_')
                        ext = os.path.splitext(filename)[1] or '.jpg'
                        filename = f"{safe_name}{ext}"
                        break
            
            # Ensure filename is safe
            safe_chars = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.")
            filename = ''.join(c for c in filename if c in safe_chars)
            
            return filename
            
        except Exception as e:
            logger.warning(f"Error generating filename for URL {url}: {e}")
            return f"image_{index}.jpg"
    
    def download_image(self, url, filename, index):
        """Download a single image"""
        try:
            # Skip if file already exists
            filepath = os.path.join(self.output_dir, filename)
            if os.path.exists(filepath):
                logger.debug(f"File already exists: {filename}")
                return {'index': index, 'url': url, 'filename': filename, 'status': 'skipped', 'error': None}
            
            # Send request with headers and timeout
            response = requests.get(
                url, 
                headers=self.headers, 
                timeout=self.timeout,
                stream=True
            )
            response.raise_for_status()
            
            # Check if content is an image
            content_type = response.headers.get('content-type', '')
            if 'image' not in content_type:
                logger.warning(f"URL does not point to an image: {url}")
                return {'index': index, 'url': url, 'filename': filename, 'status': 'failed', 'error': 'Not an image'}
            
            # Save the image
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            logger.debug(f"Downloaded: {filename}")
            return {'index': index, 'url': url, 'filename': filename, 'status': 'success', 'error': None}
            
        except requests.exceptions.RequestException as e:
            logger.warning(f"Failed to download {url}: {e}")
            return {'index': index, 'url': url, 'filename': filename, 'status': 'failed', 'error': str(e)}
        except Exception as e:
            logger.error(f"Unexpected error downloading {url}: {e}")
            return {'index': index, 'url': url, 'filename': filename, 'status': 'failed', 'error': str(e)}
    
    def download_all_images(self, df, urls):
        """Download all images with progress tracking"""
        results = []
        failed_downloads = []
        
        logger.info(f"Starting download of {len(urls)} images to '{self.output_dir}'")
        logger.info(f"Using {self.max_workers} parallel workers")
        
        # Prepare tasks
        tasks = []
        for i, url in enumerate(urls):
            # Get corresponding row from dataframe if available
            df_row = None
            if df is not None:
                # Find the row with this URL
                matching_rows = df[df[self.url_column] == url]
                if not matching_rows.empty:
                    df_row = matching_rows.iloc[0].to_dict()
            
            filename = self.get_filename(url, i, df_row)
            tasks.append((url, filename, i))
        
        # Download images in parallel
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all tasks
            future_to_task = {
                executor.submit(self.download_image, url, filename, i): (url, filename, i)
                for url, filename, i in tasks
            }
            
            # Process results as they complete
            with tqdm(total=len(tasks), desc="Downloading images") as pbar:
                for future in as_completed(future_to_task):
                    result = future.result()
                    results.append(result)
                    
                    if result['status'] == 'failed':
                        failed_downloads.append(result)
                    
                    pbar.update(1)
        
        # Generate summary report
        self.generate_report(results)
        
        # Save failed downloads to a CSV for retry
        if failed_downloads:
            self.save_failed_downloads(failed_downloads)
        
        return results
    
    def generate_report(self, results):
        """Generate download summary report"""
        total = len(results)
        successful = sum(1 for r in results if r['status'] == 'success')
        failed = sum(1 for r in results if r['status'] == 'failed')
        skipped = sum(1 for r in results if r['status'] == 'skipped')
        
        logger.info("\n" + "="*50)
        logger.info("DOWNLOAD SUMMARY")
        logger.info("="*50)
        logger.info(f"Total URLs processed: {total}")
        logger.info(f"Successfully downloaded: {successful}")
        logger.info(f"Failed downloads: {failed}")
        logger.info(f"Skipped (already exists): {skipped}")
        logger.info(f"Images saved in: {os.path.abspath(self.output_dir)}")
        
        if failed > 0:
            logger.info(f"\nFailed downloads saved to: failed_downloads.csv")
    
    def save_failed_downloads(self, failed_downloads):
        """Save failed downloads to a CSV file for retry"""
        failed_df = pd.DataFrame(failed_downloads)
        failed_df.to_csv('failed_downloads.csv', index=False)
        logger.info(f"Saved {len(failed_downloads)} failed downloads to failed_downloads.csv")

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Download images from URLs in an Excel file')
    parser.add_argument('excel_file', help='Path to the Excel file')
    parser.add_argument('--url-column', default='url', help='Column name containing URLs (default: url)')
    parser.add_argument('--output-dir', default='downloaded_images', help='Output directory (default: downloaded_images)')
    parser.add_argument('--workers', type=int, default=5, help='Number of parallel workers (default: 5)')
    parser.add_argument('--timeout', type=int, default=30, help='Request timeout in seconds (default: 30)')
    parser.add_argument('--verbose', action='store_true', help='Enable verbose logging')
    
    args = parser.parse_args()
    
    # Set logging level
    if args.verbose:
        logger.setLevel(logging.DEBUG)
    
    try:
        # Initialize downloader
        downloader = ExcelImageDownloader(
            excel_path=args.excel_file,
            url_column=args.url_column,
            output_dir=args.output_dir,
            max_workers=args.workers,
            timeout=args.timeout
        )
        
        # Load data from Excel
        df, urls = downloader.load_excel_data()
        
        # Download images
        if urls:
            results = downloader.download_all_images(df, urls)
            
            # Ask user if they want to retry failed downloads
            failed = [r for r in results if r['status'] == 'failed']
            if failed and input(f"\n{failed} downloads failed. Retry? (y/n): ").lower() == 'y':
                logger.info("Retrying failed downloads...")
                # Simple retry logic - you could enhance this
                for item in failed:
                    result = downloader.download_image(
                        item['url'], 
                        item['filename'], 
                        item['index']
                    )
                    logger.info(f"Retry result for {item['url']}: {result['status']}")
        else:
            logger.warning("No valid URLs found in the Excel file")
            
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())