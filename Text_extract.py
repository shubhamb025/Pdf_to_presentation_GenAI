import pdfplumber
import os
import fitz  # PyMuPDF
import re
from google_images_search import GoogleImagesSearch
import requests
from io import BytesIO
from PIL import Image

def extract_text_from_pdf(pdf_path, output_dir):
    """
    Maintain backward compatibility with existing code
    """
    return extract_text_and_images_from_pdf(pdf_path, output_dir)

def download_web_images(search_query, output_dir, num_images=2):
    """Download images from Google Images search"""
    # Get API credentials from environment variables
    gis = GoogleImagesSearch(os.getenv('GOOGLE_API_KEY'), os.getenv('GOOGLE_CX'))

    # Define search params
    search_params = {
        'q': search_query,
        'num': num_images,
        'safe': 'high',
        'fileType': 'jpg|png',
        'imgType': 'photo',
        'imgSize': 'large'
    }

    try:
        # Create images directory if it doesn't exist
        images_dir = os.path.join(output_dir, 'images')
        os.makedirs(images_dir, exist_ok=True)

        # Perform the search
        gis.search(search_params)

        image_titles = {}
        for i, image in enumerate(gis.results(), 1):
            try:
                # Download and save the image
                image_data = requests.get(image.url).content
                img = Image.open(BytesIO(image_data))
                
                # Generate filename
                filename = f"topic_1_img_{i}.jpg"
                filepath = os.path.join(images_dir, filename)
                
                # Save image
                img.save(filepath, 'JPEG')
                
                # Store image title
                image_titles[f"page_1_img_{i}"] = search_query
                
                print(f"Downloaded image {i}: {image.url}")
                
            except Exception as e:
                print(f"Error downloading image {i}: {str(e)}")
                continue

        # Save image titles
        titles_file = os.path.join(output_dir, 'image_titles.txt')
        with open(titles_file, 'w', encoding='utf-8') as f:
            for key, title in image_titles.items():
                f.write(f"{key}|{title}\n")

        return images_dir, image_titles

    except Exception as e:
        print(f"Error searching for images: {str(e)}")
        return None, {}

def extract_text_and_images_from_pdf(pdf_path, output_dir, is_topic_mode=False, topic_query=None):
    """Extract text and images from PDF or download web images for topic"""
    if is_topic_mode and topic_query:
        return download_web_images(topic_query, output_dir)
        
    print(f"\nStarting PDF processing...")
    print(f"PDF Path: {pdf_path}")
    print(f"Output Directory: {output_dir}")
    
    # Create the output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    print(f"Created/verified output directory: {output_dir}")
    
    # Create a subdirectory for images
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)
    print(f"Created/verified images directory: {images_dir}")

    # Dictionary to store image titles
    image_titles = {}
    images_found = False

    try:
        # First pass: Extract text and identify potential image titles
        print("\nFirst pass: Extracting text and identifying image titles...")
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                print(f"\nProcessing page {page_num}...")
                text = page.extract_text()
                tables = page.extract_tables()

                # Save text content
                output_file = os.path.join(output_dir, f"page_{page_num}.txt")
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(text + "\n\n")
                    for table in tables:
                        for row in table:
                            f.write(" | ".join(str(cell) for cell in row) + "\n")
                        f.write("\n")
                print(f"Saved text content to: {output_file}")

                # Look for potential image titles in the text
                title_patterns = [
                    r'(?:Figure|Fig\.|FIGURE)\s*(\d+)[:\.]?\s*([^\n\.]+)',
                    r'(?:Diagram|DIAGRAM)\s*(\d+)[:\.]?\s*([^\n\.]+)',
                    r'(?:Image|IMAGE)\s*(\d+)[:\.]?\s*([^\n\.]+)',
                    r'(?:Illustration|ILLUSTRATION)\s*(\d+)[:\.]?\s*([^\n\.]+)'
                ]

                for pattern in title_patterns:
                    matches = re.finditer(pattern, text, re.IGNORECASE)
                    for match in matches:
                        fig_num = match.group(1)
                        title = match.group(2).strip()
                        key = f"page_{page_num}_img_{fig_num}"
                        image_titles[key] = title
                        print(f"Found image title: {key} -> {title}")

        # Second pass: Extract and save images
        print("\nSecond pass: Extracting and saving images...")
        doc = fitz.open(pdf_path)
        total_images = 0
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            image_list = page.get_images()
            
            print(f"\nPage {page_num + 1}: Found {len(image_list)} images")
            
            if image_list:
                images_found = True
                
            for img_index, img in enumerate(image_list, 1):
                xref = img[0]
                try:
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    # Generate image filename with the pattern expected by PowerPoint creation
                    image_filename = f"page_{page_num + 1}_img_{img_index}.{image_ext}"
                    image_path = os.path.join(images_dir, image_filename)
                    
                    print(f"Saving image {img_index} to: {image_path}")
                    
                    # Save image
                    with open(image_path, "wb") as image_file:
                        image_file.write(image_bytes)
                    print(f"Successfully saved image: {image_filename}")
                    total_images += 1

                    # Save image title if found, otherwise use a default title
                    title_key = f"page_{page_num + 1}_img_{img_index}"
                    if title_key not in image_titles:
                        # Try to extract title from surrounding text
                        rect = page.get_image_bbox(xref)
                        if rect:
                            surrounding_text = page.get_text("text", clip=rect)
                            # Look for potential title in the text above the image
                            lines = surrounding_text.split('\n')
                            for line in lines:
                                if line.strip() and len(line.strip()) > 5:  # Reasonable title length
                                    image_titles[title_key] = line.strip()
                                    print(f"Extracted title from context: {title_key} -> {line.strip()}")
                                    break
                        if title_key not in image_titles:
                            image_titles[title_key] = f"Figure {img_index}"
                            print(f"Using default title: {title_key} -> Figure {img_index}")
                except Exception as img_error:
                    print(f"Error extracting image {img_index} on page {page_num + 1}: {str(img_error)}")
                    continue

        # Save image titles to a file
        titles_file = os.path.join(output_dir, "image_titles.txt")
        with open(titles_file, 'w', encoding='utf-8') as f:
            for key, title in image_titles.items():
                f.write(f"{key}|{title}\n")
        print(f"\nSaved {len(image_titles)} image titles to: {titles_file}")

        if not images_found:
            print(f"\nNo images found in {pdf_path}")
            # Create a marker file to indicate no images were found
            with open(os.path.join(images_dir, "no_images.txt"), 'w') as f:
                f.write("No images were found in this PDF.")
        else:
            print(f"\nSuccessfully extracted {total_images} images")
            # List all extracted images
            print("\nExtracted images:")
            for img_file in os.listdir(images_dir):
                if img_file.endswith(('.png', '.jpg', '.jpeg', '.gif')):
                    print(f"- {img_file}")

        return images_dir, image_titles

    except Exception as e:
        print(f"\nError processing PDF: {str(e)}")
        import traceback
        traceback.print_exc()
        # Ensure the images directory exists even if there was an error
        os.makedirs(images_dir, exist_ok=True)
        # Create a marker file to indicate an error occurred
        with open(os.path.join(images_dir, "extraction_error.txt"), 'w') as f:
            f.write(f"Error processing PDF: {str(e)}\n")
            f.write(traceback.format_exc())
        return images_dir, {}

# Remove the example usage from here
