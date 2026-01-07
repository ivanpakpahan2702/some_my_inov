import os
import sys
from PIL import Image
from pathlib import Path

class BulkImageConverter:
    def __init__(self):
        self.supported_formats = {
            'png': 'PNG',
            'jpg': 'JPEG',  # Pillow uses 'JPEG' for jpg files
            'jpeg': 'JPEG',
            'jfif': 'JPEG',
            'jpe': 'JPEG',
            'bmp': 'BMP',
            'webp': 'WEBP',
            'tiff': 'TIFF',
            'tif': 'TIFF',
            'gif': 'GIF',
            'ico': 'ICO',
            'ppm': 'PPM',
            'pgm': 'PGM'
        }
        
        # Mapping from user input to Pillow format
        self.format_mapping = {
            'png': 'PNG',
            'jpg': 'JPEG',
            'jpeg': 'JPEG', 
            'jfif': 'JPEG',
            'jpe': 'JPEG',
            'bmp': 'BMP',
            'webp': 'WEBP',
            'tiff': 'TIFF',
            'tif': 'TIFF',
            'gif': 'GIF',
            'ico': 'ICO'
        }
    
    def get_supported_extensions(self):
        """Return list of supported input extensions"""
        return list(self.supported_formats.keys())
    
    def get_pillow_format(self, format_name):
        """Convert user-friendly format name to Pillow format"""
        format_name = format_name.lower()
        return self.format_mapping.get(format_name, format_name.upper())
    
    def convert_image(self, input_path, output_path, output_format):
        """
        Convert single image to target format
        """
        try:
            # Convert format name to Pillow format
            pillow_format = self.get_pillow_format(output_format)
            
            with Image.open(input_path) as img:
                # Convert to RGB if necessary (for JPEG formats)
                if pillow_format == 'JPEG' and img.mode in ('RGBA', 'LA', 'P'):
                    rgb_img = Image.new('RGB', img.size, (255, 255, 255))
                    rgb_img.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                    img = rgb_img
                elif img.mode == 'P':
                    img = img.convert('RGB')
                
                # Save in target format
                img.save(output_path, format=pillow_format)
                print(f"Successfully converted: {input_path.name} -> {output_path.name}")
                return True
                
        except Exception as e:
            print(f"Error converting {input_path}: {str(e)}")
            return False
    
    def get_output_extension(self, format_name):
        """
        Get appropriate file extension for output format
        """
        format_name = format_name.lower()
        # Special cases
        if format_name in ['jpg', 'jpeg', 'jfif', 'jpe']:
            return 'jpg'
        elif format_name in ['tif', 'tiff']:
            return 'tiff'
        else:
            return format_name
    
    def bulk_convert(self, input_folder, output_folder, output_format, 
                    recursive=False, overwrite=False):
        """
        Convert all images in input folder to target format
        """
        if output_format.lower() not in self.supported_formats:
            print(f"Unsupported output format: {output_format}")
            print(f"Supported formats: {', '.join(self.supported_formats.keys())}")
            return
        
        # Create output folder if it doesn't exist
        Path(output_folder).mkdir(parents=True, exist_ok=True)
        
        # Get all image files
        image_files = []
        pattern = "**/*" if recursive else "*"
        
        for ext in self.get_supported_extensions():
            search_pattern = f"{pattern}.{ext}"
            image_files.extend(Path(input_folder).glob(search_pattern))
            # Also check uppercase extensions
            image_files.extend(Path(input_folder).glob(search_pattern.upper()))
        
        if not image_files:
            print("No supported image files found!")
            return
        
        print(f"Found {len(image_files)} image files to convert")
        print(f"Converting to: {output_format.upper()}")
        print(f"Using Pillow format: {self.get_pillow_format(output_format)}")
        print("-" * 50)
        
        converted_count = 0
        skipped_count = 0
        failed_count = 0
        
        for input_path in image_files:
            # Generate output filename
            output_extension = self.get_output_extension(output_format)
            output_filename = f"{input_path.stem}.{output_extension}"
            output_path = Path(output_folder) / output_filename
            
            # Check if output file already exists
            if output_path.exists() and not overwrite:
                print(f"Skipped (exists): {input_path.name}")
                skipped_count += 1
                continue
            
            # Convert image
            success = self.convert_image(input_path, output_path, output_format)
            
            if success:
                converted_count += 1
            else:
                failed_count += 1
        
        print("-" * 50)
        print(f"Conversion complete!")
        print(f"Successfully converted: {converted_count}")
        print(f"Skipped: {skipped_count}")
        print(f"Failed: {failed_count}")

def main():
    converter = BulkImageConverter()
    
    print("=== Bulk Image Format Converter ===")
    print(f"Supported input formats: {', '.join(converter.get_supported_extensions())}")
    print(f"Supported output formats: {', '.join(converter.supported_formats.keys())}")
    print()
    
    # Get user input
    input_folder = input("Enter input folder path: ").strip()
    if not os.path.exists(input_folder):
        print("Input folder does not exist!")
        return
    
    output_folder = input("Enter output folder path: ").strip()
    if not output_folder:
        output_folder = input_folder + "_converted"
    
    print("\nAvailable output formats:")
    for i, fmt in enumerate(converter.supported_formats.keys(), 1):
        print(f"{i}. {fmt.upper()}")
    
    output_format = input("\nEnter target format: ").strip().lower()
    
    if output_format not in converter.supported_formats:
        print("Invalid format selected!")
        return
    
    recursive = input("Search subdirectories? (y/n): ").strip().lower() == 'y'
    overwrite = input("Overwrite existing files? (y/n): ").strip().lower() == 'y'
    
    print("\nStarting conversion...")
    converter.bulk_convert(
        input_folder=input_folder,
        output_folder=output_folder,
        output_format=output_format,
        recursive=recursive,
        overwrite=overwrite
    )

if __name__ == "__main__":
    # If command line arguments are provided, use them
    if len(sys.argv) > 3:
        converter = BulkImageConverter()
        input_folder = sys.argv[1]
        output_folder = sys.argv[2]
        output_format = sys.argv[3]
        recursive = len(sys.argv) > 4 and sys.argv[4].lower() == '--recursive'
        overwrite = len(sys.argv) > 5 and sys.argv[5].lower() == '--overwrite'
        
        converter.bulk_convert(input_folder, output_folder, output_format, recursive, overwrite)
    else:
        main()