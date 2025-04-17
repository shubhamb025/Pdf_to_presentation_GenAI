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

def fetch_images_for_topic(topic, slide_titles, num_images=4):
    """Fetch relevant images for the topic and slides from web search"""
    try:
        print(f"\nStarting image fetch for topic: {topic}")
        images_dir = os.path.join(app.config['EXTRACT_FOLDER'], 'images')
        os.makedirs(images_dir, exist_ok=True)
        print(f"Images will be saved to: {images_dir}")
        
        # Create image_titles.txt file
        titles_file = os.path.join(app.config['EXTRACT_FOLDER'], 'image_titles.txt')
        print(f"Image titles will be saved to: {titles_file}")
        
        # List to keep track of saved images
        saved_images = []
        image_titles = {}
        
        # Generate search terms from slide titles
        search_terms = []
        
        # Process slide titles to create meaningful search terms
        if slide_titles:
            for title in slide_titles:
                # Skip generic slides
                if any(term in title.lower() for term in ["agenda", "overview", "conclusion", "summary", "thank you", "key concept", "index"]):
                    continue
                
                # Clean and enhance the title for search
                clean_title = title.strip()
                if clean_title:
                    # Remove common suffixes after colon or hyphen
                    if ':' in clean_title:
                        clean_title = clean_title.split(':')[0].strip()
                    if '-' in clean_title:
                        clean_title = clean_title.split('-')[0].strip()
                    
                    # Remove common words and phrases that don't add to the search
                    words_to_remove = ['the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'key', 'concept', 'index', 'unveiling', 'power']
                    phrases_to_remove = ['unveiling power', 'unveiling the power']
                    
                    # Remove phrases first
                    for phrase in phrases_to_remove:
                        clean_title = clean_title.lower().replace(phrase, '').strip()
                    
                    # Then remove individual words
                    title_words = clean_title.split()
                    filtered_words = [word for word in title_words if word.lower() not in words_to_remove]
                    
                    if filtered_words:
                        # Join the remaining words to form the search term
                        search_term = ' '.join(filtered_words)
                        if len(search_term.split()) > 1:  # Only add if more than one word
                            search_terms.append(search_term)
                            print(f"Added search term from title: {search_term}")
        
        # Add main topic if not already covered
        if topic.lower() not in [term.lower() for term in search_terms]:
            # Clean the topic the same way
            clean_topic = topic.lower()
            for phrase in phrases_to_remove:
                clean_topic = clean_topic.replace(phrase, '').strip()
            
            topic_words = clean_topic.split()
            filtered_topic = ' '.join([word for word in topic_words if word.lower() not in words_to_remove])
            
            if filtered_topic and len(filtered_topic.split()) > 1:
                search_terms.append(filtered_topic)
                print(f"Added main topic search term: {filtered_topic}")
        
        # Remove duplicates while preserving order
        search_terms = list(dict.fromkeys(search_terms))
        
        # Limit the number of search terms
        search_terms = search_terms[:num_images]
        print(f"\nWill search for {len(search_terms)} terms to get {num_images} images")

        # Check if API credentials are available
        api_key = os.getenv('GOOGLE_API_KEY')
        cx = os.getenv('GOOGLE_CX')
        
        if not api_key or not cx:
            print("Error: Google API credentials not found in environment variables")
            return []

        print(f"Using Google API Key: {api_key[:10]}... and CX: {cx}")
        
        try:
            # Initialize Google Images Search
            gis = GoogleImagesSearch(api_key, cx)
            print("Successfully initialized Google Images Search")
        except Exception as init_error:
            print(f"Error initializing Google Images Search: {str(init_error)}")
            return []

        # Define base search params
        search_params = {
            'q': '',  # Will be updated for each search
            'num': 2,  # Number of images per search term
            'safe': 'active',  # Use 'active' instead of 'high'
            'fileType': 'jpg|png',
            'imgType': 'photo',
            'imgSize': 'large',
            'searchType': 'image'  # Explicitly specify image search
        }
        
        for idx, search_term in enumerate(search_terms):
            if len(saved_images) >= num_images:
                print(f"\nReached desired number of images ({num_images})")
                break
                
            try:
                print(f"\nProcessing search term {idx + 1}/{len(search_terms)}: {search_term}")
                
                # Update search query
                search_params['q'] = search_term
                print(f"Search parameters: {search_params}")
                
                try:
                    # Perform the search
                    gis.search(search_params)
                    results = list(gis.results())
                    print(f"Found {len(results)} results for search term: {search_term}")
                    
                    if not results:
                        print("No images found for this search term")
                        continue
                    
                    # Process results
                    for i, image in enumerate(results, 1):
                        try:
                            print(f"Processing image {i} from {image.url}")
                            
                            # Download and save the image
                            image_data = requests.get(image.url, timeout=10).content
                            img = Image.open(BytesIO(image_data))
                            
                            # Generate filename
                            page_num = idx + 1
                            img_num = len(saved_images) + 1
                            file_extension = '.jpg'
                            image_filename = f"page_{page_num}_img_{img_num}{file_extension}"
                            image_path = os.path.join(images_dir, image_filename)
                            
                            print(f"Saving image to: {image_path}")
                            
                            # Save image
                            img.save(image_path, 'JPEG')
                            
                            # Verify the image was saved
                            if os.path.exists(image_path):
                                file_size = os.path.getsize(image_path)
                                print(f"Image saved successfully. File size: {file_size} bytes")
                                
                                # Create image title key
                                img_key = f"page_{page_num}_img_{img_num}"
                                image_titles[img_key] = search_term
                                saved_images.append(image_path)
                                
                                print(f"Added image title: {search_term}")
                                
                                if len(saved_images) >= num_images:
                                    break
                            else:
                                print(f"Warning: Failed to save image to {image_path}")
                                
                        except Exception as img_error:
                            print(f"Error saving image: {str(img_error)}")
                            continue
                            
                except Exception as search_error:
                    print(f"Error performing search: {str(search_error)}")
                    continue
                    
            except Exception as term_error:
                print(f"Error processing search term '{search_term}': {str(term_error)}")
                continue
        
        # Write image titles to file if we have saved images
        if saved_images:
            print(f"\nWriting {len(image_titles)} image titles to {titles_file}")
            with open(titles_file, 'w', encoding='utf-8') as f:
                for key, title in image_titles.items():
                    f.write(f"{key}|{title}\n")
            
            print(f"\nSuccessfully saved {len(saved_images)} images for the presentation")
            print("Saved images:")
            for img in saved_images:
                print(f"- {os.path.basename(img)}")
        else:
            print("\nNo images were successfully downloaded")
            
        return saved_images
    
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
