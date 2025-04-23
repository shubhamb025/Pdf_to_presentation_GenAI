import os
import traceback
import shutil
import requests
import uuid
import urllib.parse
from io import BytesIO
from flask import Flask, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename
import Text_extract
import txt_to_vba
import vba_to_ppt
from google_images_search import GoogleImagesSearch
from PIL import Image

app = Flask(__name__, static_folder='static')

UPLOAD_FOLDER = 'uploads'
EXTRACT_FOLDER = 'extract'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXTRACT_FOLDER'] = EXTRACT_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Define themes with preview images
THEMES = {
    'theme1': {
        'name': 'Theme 1',
        'file': 'Presentation1.pptx',
        'preview': 'presentation1.png'
    },
    'theme2': {
        'name': 'Theme 2',
        'file': 'Presentation2.pptx',
        'preview': 'presentation2.png'
    },
    'theme3': {
        'name': 'Theme 3',
        'file': 'Presentation3.pptx',
        'preview': 'presentation3.png'
    },
    'theme4': {
        'name': 'Theme 4',
        'file': 'Presentation4.pptx',
        'preview': 'presentation4.png'
    }
}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup_folders():
    """Clean up the uploads and extract folders after PowerPoint generation"""
    try:
        # Clean uploads folder
        uploads_folder = app.config['UPLOAD_FOLDER']
        for file in os.listdir(uploads_folder):
            file_path = os.path.join(uploads_folder, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print("Cleaned uploads folder")

        # Clean extract folder
        extract_folder = app.config['EXTRACT_FOLDER']
        for item in os.listdir(extract_folder):
            item_path = os.path.join(extract_folder, item)
            if os.path.isfile(item_path):
                os.remove(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
        print("Cleaned extract folder")

    except Exception as e:
        print(f"Error during cleanup: {str(e)}")
        traceback.print_exc()

def fetch_images_for_topic(topic, slide_titles, num_images=2):
    """Fetch relevant images for the given topic and slide titles."""
    try:
        # Create directory for images if it doesn't exist
        images_dir = os.path.join('extract', 'images')
        os.makedirs(images_dir, exist_ok=True)
        
        # Create a file to store image titles
        titles_file = os.path.join('extract', 'image_titles.txt')
        
        # Generate one search term per slide title, skipping slide 0 (title slide)
        search_terms = []
        generic_terms = ['challenges', 'limitations', 'future', 'conclusion', 'key takeaways', 'applications']
        
        for i, title in enumerate(slide_titles):
            # Skip slide 0 (title slide) and generic terms
            if i == 0 or title.lower() in ['agenda', 'summary', 'thank you', 'questions', 'index']:
                continue
            
            # Check if the title contains generic terms and doesn't mention the topic
            title_lower = title.lower()
            is_generic = any(term in title_lower for term in generic_terms)
            topic_in_title = any(word in title_lower for word in topic.lower().split())
            
            # If title is generic and doesn't contain the topic, append the topic
            if is_generic and not topic_in_title:
                search_term = f"{title} of {topic}"
            else:
                search_term = title
                
            search_terms.append(search_term)
        
        # Add the main topic as one search term
        search_terms.append(topic)
        
        # Remove duplicates while preserving order
        seen = set()
        search_terms = [x for x in search_terms if not (x in seen or seen.add(x))]
        
        # Limit the number of search terms
        max_terms = min(len(search_terms), num_images * 2)
        search_terms = search_terms[:max_terms]
        
        print(f"Generated {len(search_terms)} search terms: {search_terms}")
        
        # Initialize Google Images Search
        try:
            api_key = os.getenv('GOOGLE_API_KEY')
            cx = os.getenv('GOOGLE_CX')
            
            if not api_key or not cx:
                print("Google API credentials not found in environment variables")
                return
            
            from googleapiclient.discovery import build
            service = build("customsearch", "v1", developerKey=api_key)
            
        except Exception as e:
            print(f"Error initializing Google Images Search: {str(e)}")
            return
        
        # Track downloaded images
        downloaded_images = []
        
        # Try each search term
        for term in search_terms:
            if len(downloaded_images) >= num_images:
                break
                
            print(f"\nSearching for images with term: {term}")
            
            try:
                # Perform the search
                result = service.cse().list(
                    q=term,
                    cx=cx,
                    searchType='image',
                    safe='active',
                    num=1  # Get one image per term
                ).execute()
                
                if 'items' in result:
                    for item in result['items']:
                        if len(downloaded_images) >= num_images:
                            break
                            
                        image_url = item['link']
                        print(f"Found image URL: {image_url}")
                        
                        try:
                            # Download the image
                            response = requests.get(image_url, timeout=10)
                            if response.status_code == 200 and len(response.content) > 10000:  # Ensure it's a valid image (>10KB)
                    # Generate a unique filename
                                img_num = len(downloaded_images) + 1
                                img_key = f"page_{img_num}_img_{img_num}"
                                img_path = os.path.join(images_dir, f"{img_key}.jpg")
                    
                    # Save the image
                                with open(img_path, 'wb') as f:
                                    f.write(response.content)
                    
                                # Save the image title
                                with open(titles_file, 'a', encoding='utf-8') as f:
                                    f.write(f"{img_key}|{term}\n")
                                
                                downloaded_images.append(img_path)
                                print(f"Successfully downloaded image: {img_path}")
                                break  # Break after successful download
                            
                        except Exception as e:
                            print(f"Error downloading image: {str(e)}")
                            continue
                
            except Exception as e:
                print(f"Error searching for term '{term}': {str(e)}")
                continue
        
        print(f"\nDownloaded {len(downloaded_images)} images successfully")
        return downloaded_images
    
    except Exception as e:
        print(f"Error in fetch_images_for_topic: {str(e)}")
        traceback.print_exc()
        return []

def generate_topic_content(topic, details="", slide_count=8, presentation_rules=""):
    """Generate presentation content based on a topic using Gemini"""
    # Format the prompt for Gemini
    prompt = f"""
    Create a comprehensive presentation outline on the topic: {topic}.
    
    Additional details: {details}
    
    For this outline, include:
    1. An engaging title slide
    2. A brief agenda/overview slide
    3. {slide_count} detailed content slides with key points
    4. A conclusion slide with summary and takeaways
    """
    
    # Add presentation rules if provided
    if presentation_rules:
        prompt += f"""
    
    PRESENTATION RULES:
    {presentation_rules}
    """
    
    prompt += """
    Format your response using the following structure:
    
    TITLE: [Presentation Title]
    SUBTITLE: [Optional Subtitle]
    
    SLIDE 1: [Slide Title]
    - [Bullet point 1]
    - [Bullet point 2]
    
    SLIDE 2: [Slide Title]
    - [Bullet point 1]
    - [Bullet point 2]
    
    [... and so on for each slide]
    """
    
    # Call Gemini to generate the outline
    gemini_output = txt_to_vba.generate_outline_with_gemini(prompt, slide_count)
    return gemini_output

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    # Create themes list with correct image paths
    themes = [
        {
            'key': key,
            'name': theme['name'],
            'preview': f'/static/theme_images/{theme["preview"]}'
        }
        for key, theme in THEMES.items()
    ]
    
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('index.html', message='No file part', themes=themes)
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', message='No selected file', themes=themes)
        
        # Get selected theme
        theme_key = request.form.get('theme', 'theme1')
        
        # Get the creator name from the form
        creator_name = request.form.get('creator_name', '')
        
        # Get presentation rules from the form
        presentation_rules = request.form.get('presentation_rules', '')
        
        if file and allowed_file(file.filename):
            try:
                filename = secure_filename(file.filename)
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(pdf_path)
                print(f"Saved PDF to: {pdf_path}")
                
                # Extract text and images from PDF
                images_dir, image_titles = Text_extract.extract_text_and_images_from_pdf(pdf_path, app.config['EXTRACT_FOLDER'])
                print(f"Extracted text and images from PDF. Images directory: {images_dir}")
                print(f"Found {len(image_titles)} image titles")
                
                # Read the extracted content
                content = txt_to_vba.read_input_files(app.config['EXTRACT_FOLDER'])
                
                # Add presentation rules to the content if provided
                if presentation_rules:
                    content += f"\n\nPRESENTATION RULES:\n{presentation_rules}"
                    print(f"Added presentation rules: {presentation_rules}")
                    print("IMPORTANT: Applying specific presentation formatting rules to the content.")
                
                # Generate outline using Gemini
                num_content_slides = 6  # You can adjust this number as needed
                gemini_output = txt_to_vba.generate_outline_with_gemini(content, num_content_slides)
                
                # Generate VBA code with creator name
                slides = txt_to_vba.parse_gemini_output(gemini_output)
                vba_code = txt_to_vba.generate_vba_code(slides, creator_name)
                
                # Save the VBA code
                with open('create_presentation.vba', 'w', encoding='utf-8') as f:
                    f.write(vba_code)
                
                # Set output path for PowerPoint
                ppt_output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'generated_presentation.pptx')
                
                # Convert VBA to PowerPoint with selected theme
                ppt_path = vba_to_ppt.create_powerpoint(
                    vba_to_ppt.parse_vba_file('create_presentation.vba'),
                    ppt_output_path,
                    theme_key,
                    creator_name
                )
                print(f"PowerPoint path: {ppt_path}")
                
                if ppt_path and os.path.exists(ppt_path):
                    # Verify the PowerPoint was created successfully
                    print("PowerPoint created successfully")
                    
                    # Clean up after successful PowerPoint generation
                    cleanup_folders()
                    
                    # Send the file
                    response = send_file(ppt_path, as_attachment=True)
                    return response
                else:
                    raise FileNotFoundError(f"Generated PowerPoint file not found: {ppt_path}")
            except Exception as e:
                error_message = f"An error occurred: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
                print(error_message)  # Print to console for debugging
                # Clean up even if there's an error
                cleanup_folders()
                return render_template('index.html', message=error_message, themes=themes)
    
    return render_template('index.html', themes=themes)

@app.route('/generate-from-topic', methods=['POST'])
def generate_from_topic():
    # Create themes list with correct image paths
    themes = [
        {
            'key': key,
            'name': theme['name'],
            'preview': f'/static/theme_images/{theme["preview"]}'
        }
        for key, theme in THEMES.items()
    ]
    
    try:
        # Get form data
        topic = request.form.get('topic', '')
        details = request.form.get('details', '')
        theme_key = request.form.get('theme', 'theme1')
        creator_name = request.form.get('creator_name', '')
        
        # Get slide count (default to 8 if not provided)
        try:
            slide_count = int(request.form.get('slide_count', 8))
            # Ensure slide count is within reasonable limits
            slide_count = max(4, min(slide_count, 12))
        except ValueError:
            slide_count = 8
        
        # Get presentation rules from the form
        presentation_rules = request.form.get('presentation_rules', '')
        
        if not topic:
            return render_template('topic_generator.html', message='Error: No topic provided', themes=themes)
        
        print(f"Generating presentation on topic: {topic}")
        print(f"Additional details: {details}")
        print(f"Using theme: {theme_key}")
        print(f"Creator: {creator_name}")
        print(f"Slide count: {slide_count}")
        if presentation_rules:
            print(f"Presentation rules specified:")
            for line in presentation_rules.strip().split('\n'):
                print(f"  - {line.strip()}")
        else:
            print("No presentation rules specified, using default formatting")
        
        # Ensure extract directory exists
        os.makedirs(app.config['EXTRACT_FOLDER'], exist_ok=True)
        os.makedirs(os.path.join(app.config['EXTRACT_FOLDER'], 'images'), exist_ok=True)
        
        # Generate presentation content based on the topic
        gemini_output = generate_topic_content(topic, details, slide_count, presentation_rules)
        
        # Parse the output and generate VBA code
        slides = txt_to_vba.parse_gemini_output(gemini_output)
        vba_code = txt_to_vba.generate_vba_code(slides, creator_name)
        
        # Extract slide titles for image search
        slide_titles = [slide['title'] for slide in slides if slide['title']]
        
        # Fetch relevant images for the topic
        fetch_images_for_topic(topic, slide_titles, num_images=min(8, slide_count + 2))
        
        # Save the VBA code
        with open('create_presentation.vba', 'w', encoding='utf-8') as f:
            f.write(vba_code)
        
        # Set output path for PowerPoint
        output_filename = f"topic_{topic.replace(' ', '_')[:30]}.pptx"
        output_dir = app.config['OUTPUT_FOLDER']
        os.makedirs(output_dir, exist_ok=True)
        ppt_output_path = os.path.join(output_dir, output_filename)
        
        # Convert VBA to PowerPoint with selected theme
        ppt_path = vba_to_ppt.create_powerpoint(
            vba_to_ppt.parse_vba_file('create_presentation.vba'),
            ppt_output_path,
            theme_key
        )
        
        if ppt_path and os.path.exists(ppt_path):
            # Clean up after successful PowerPoint generation
            cleanup_folders()
            
            # Send the file
            response = send_file(ppt_path, as_attachment=True)
            return response
        else:
            raise FileNotFoundError(f"Generated PowerPoint file not found: {ppt_path}")
            
    except Exception as e:
        error_message = f"An error occurred: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(error_message)  # Print to console for debugging
        # Clean up even if there's an error
        cleanup_folders()
        return render_template('topic_generator.html', message=error_message, themes=themes)

@app.route('/topic-generator')
def topic_generator():
    # Create themes list with correct image paths
    themes = [
        {
            'key': key,
            'name': theme['name'],
            'preview': f'/static/theme_images/{theme["preview"]}'
        }
        for key, theme in THEMES.items()
    ]
    return render_template('topic_generator.html', themes=themes)

if __name__ == '__main__':
    # Create necessary directories
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(EXTRACT_FOLDER, exist_ok=True)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    # Clean up any leftover files from previous runs
    cleanup_folders()
    
    app.run(debug=True)
