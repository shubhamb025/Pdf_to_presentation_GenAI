import os
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
from dotenv import load_dotenv

# Load the .env file
load_dotenv()

# Get the API key from the .env file
api_key = os.getenv('GOOGLE_API_KEY')

if not api_key:
    raise ValueError("GOOGLE_API_KEY environment variable is not set in .env file")

# Configure the genai library with the API key
genai.configure(api_key=api_key)

class RuleProcessor:
    """Class to process and enforce presentation rules"""
    
    def __init__(self):
        self.rule_categories = {
            'bullet': ['bullet', 'point', 'bullet point', 'bulletpoint'],
            'question': ['question', 'answer', 'q&a', 'q & a', 'question and answer'],
            'length': ['sentence', 'word', 'character', 'length', 'short', 'concise'],
            'table': ['table', 'tabular', 'row', 'column'],
            'diagram': ['diagram', 'chart', 'graph', 'visual', 'figure', 'illustration'],
            'format': ['format', 'structure', 'organize'],
            'quantity': ['number', 'count', 'limit', 'max', 'maximum', 'min', 'minimum', 'exactly']
        }
        
        # Initialize default formatter settings
        self.formatter_settings = {
            'use_bullets': True,
            'max_bullet_length': 120,  # characters
            'min_bullets_per_slide': 3,
            'max_bullets_per_slide': 7,
            'use_qa_format': False,
            'use_tables': False,
            'use_diagrams': False,
            'format_type': 'standard'  # standard, qa, simplified
        }
    
    def extract_rules(self, content):
        """Extract presentation rules from content"""
        if "PRESENTATION RULES:" not in content:
            return content, []
            
        parts = content.split("PRESENTATION RULES:")
        content = parts[0].strip()
        rules_text = parts[1].strip()
        
        # Split into individual rules
        rules = [rule.strip() for rule in rules_text.split('\n') if rule.strip()]
        # Clean up rule format
        rules = [rule[2:].strip() if rule.startswith('- ') or rule.startswith('• ') else rule for rule in rules]
        
        print(f"Extracted {len(rules)} presentation rules")
        return content, rules
    
    def analyze_rules(self, rules):
        """Analyze the extracted rules and configure formatter settings"""
        if not rules:
            return self.formatter_settings
        
        settings = self.formatter_settings.copy()
        
        for rule in rules:
            rule_lower = rule.lower()
            
            # Check for bullet point rules
            if any(term in rule_lower for term in self.rule_categories['bullet']):
                settings['use_bullets'] = True
                
                # Check for "only bullet points" or similar phrasing
                if 'only' in rule_lower and any(term in rule_lower for term in self.rule_categories['bullet']):
                    settings['format_type'] = 'bullets_only'
                
                # Check for bullet point length restrictions
                if any(term in rule_lower for term in self.rule_categories['length']):
                    if '1-2 sentence' in rule_lower or 'one to two sentence' in rule_lower:
                        settings['max_bullet_length'] = 120
                    elif 'short' in rule_lower or 'brief' in rule_lower or 'concise' in rule_lower:
                        settings['max_bullet_length'] = 80
                    elif 'one sentence' in rule_lower or '1 sentence' in rule_lower:
                        settings['max_bullet_length'] = 60
                
            # Check for question and answer format
            if any(term in rule_lower for term in self.rule_categories['question']):
                settings['use_qa_format'] = True
                settings['format_type'] = 'qa'
            
            # Check for bullet point quantity rules
            if any(term in rule_lower for term in self.rule_categories['quantity']):
                if 'exactly 5' in rule_lower or 'only 5' in rule_lower:
                    settings['min_bullets_per_slide'] = 5
                    settings['max_bullets_per_slide'] = 5
                elif 'exactly 3' in rule_lower or 'only 3' in rule_lower:
                    settings['min_bullets_per_slide'] = 3
                    settings['max_bullets_per_slide'] = 3
                elif 'maximum' in rule_lower or 'max' in rule_lower or 'at most' in rule_lower:
                    for num in range(2, 11):  # Check for numbers 2-10
                        if str(num) in rule_lower:
                            settings['max_bullets_per_slide'] = num
                            break
                elif 'minimum' in rule_lower or 'min' in rule_lower or 'at least' in rule_lower:
                    for num in range(2, 11):  # Check for numbers 2-10
                        if str(num) in rule_lower:
                            settings['min_bullets_per_slide'] = num
                            break
            
            # Check for table format request
            if any(term in rule_lower for term in self.rule_categories['table']):
                settings['use_tables'] = True
                if 'all' in rule_lower or 'only' in rule_lower:
                    settings['format_type'] = 'tables'
            
            # Check for diagram/visual format request
            if any(term in rule_lower for term in self.rule_categories['diagram']):
                settings['use_diagrams'] = True
                if 'instead of text' in rule_lower or 'only' in rule_lower:
                    settings['format_type'] = 'diagrams'
                    
        print(f"Applied rule settings: {settings}")
        return settings
    
    def generate_rule_prompt(self, rules, settings):
        """Generate a prompt that specifically instructs the AI based on the analyzed rules"""
        if not rules:
            return ""
            
        rule_prompt = "FORMATTING INSTRUCTIONS:\n"
        
        # Apply specific formatting based on the rule analysis
        if settings['format_type'] == 'bullets_only':
            rule_prompt += "- Use ONLY bullet points. NO paragraph text allowed.\n"
            rule_prompt += f"- Each bullet point must be {settings['max_bullet_length']} characters or less.\n"
            rule_prompt += f"- Each slide must have {settings['min_bullets_per_slide']}-{settings['max_bullets_per_slide']} bullet points.\n"
        
        elif settings['format_type'] == 'qa':
            rule_prompt += "- Format all content as questions and answers.\n"
            rule_prompt += "- Each slide should have a question as the title.\n"
            rule_prompt += "- The bullet points should provide the answer to the question.\n"
        
        elif settings['format_type'] == 'tables':
            rule_prompt += "- Format content using tables wherever possible.\n"
            rule_prompt += "- Present information in a structured tabular format.\n"
            rule_prompt += "- Use bullet points only when tables are not suitable.\n"
        
        elif settings['format_type'] == 'diagrams':
            rule_prompt += "- Describe what diagrams should be created instead of providing text.\n"
            rule_prompt += "- For each slide, suggest a diagram type and describe what it should contain.\n"
            rule_prompt += "- Provide minimal text explanation along with diagram descriptions.\n"
        
        # Add general rule for bullet point length if specified
        if settings['max_bullet_length'] < 120 and settings['format_type'] != 'bullets_only':
            rule_prompt += f"- Keep all bullet points under {settings['max_bullet_length']} characters.\n"
        
        # Add specific quantity restrictions
        if settings['min_bullets_per_slide'] == settings['max_bullets_per_slide'] and settings['format_type'] != 'bullets_only':
            rule_prompt += f"- Each slide must have EXACTLY {settings['min_bullets_per_slide']} bullet points.\n"
        
        # Add original rules for reference
        rule_prompt += "\nORIGINAL RULES TO FOLLOW:\n"
        for rule in rules:
            rule_prompt += f"- {rule}\n"
            
        return rule_prompt

def read_input_files(folder_path):
    combined_content = ""
    for filename in sorted(os.listdir(folder_path)):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r', encoding='utf-8') as file:
                combined_content += file.read() + "\n\n"
    return combined_content[:8000] 

def generate_outline_with_gemini(content, num_content_slides):
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    # Process rules using the specialized rule processor
    rule_processor = RuleProcessor()
    cleaned_content, rules = rule_processor.extract_rules(content)
    formatter_settings = rule_processor.analyze_rules(rules)
    formatted_rules_prompt = rule_processor.generate_rule_prompt(rules, formatter_settings)
    
    # Create the base prompt
    prompt = f"""
    Create a detailed PowerPoint presentation outline based on the following content:

    {cleaned_content}

    For each slide, also suggest relevant image search queries that would enhance the content.
    Each slide should have 1-2 relevant images.

    Generate an outline with the following structure:
    1. Title Slide
    2. Index Slide (will be generated automatically, don't include in your output)
    3-{num_content_slides+2}. Content Slides (5-7 key points per slide)
    {num_content_slides+3}. Conclusion Slide

    For each slide (except the index slide), provide:
    - Slide Title
    - 5-7 Key Points (as bullet points, not paragraphs)
    - Image Search Query: [1-2 relevant search terms for images]

    YOU MUST FORMAT YOUR OUTPUT EXACTLY AS FOLLOWS (this format is required for parsing):
    [Slide 1]
    Title: [Presentation Title]
    Image Search Query: [search term for title slide image]

    [Slide 2]
    Title: Index
    - 1: [First Content Slide Title]
    - 2: [Second Content Slide Title]
    - 3: [Third Content Slide Title]
    - ...
    - N: [Last Content Slide Title]
    - Conclusion

    [Slide 3]
    Title: [First Content Slide Title]
    - [Key Point 1]
    - [Key Point 2]
    - [Key Point 3]
    - [Key Point 4]
    - [Key Point 5]
    Image Search Query: [search term for this slide's images]
    """
    
    # Add formatted rules if they exist
    if formatted_rules_prompt:
        prompt += f"""
    {formatted_rules_prompt}
    
    These formatting requirements are MANDATORY. Your content MUST strictly adhere to these specific requirements completely.
    """

    generation_config = {
        "temperature": 0.6,
        "top_p": 0.9,
        "top_k": 32,
        "max_output_tokens": 2048,
    }

    safety_settings = [
        {
            "category": HarmCategory.HARM_CATEGORY_HARASSMENT,
            "threshold": HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
        },
        {
            "category": HarmCategory.HARM_CATEGORY_HATE_SPEECH,
            "threshold": HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
        },
        {
            "category": HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
            "threshold": HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
        },
        {
            "category": HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
            "threshold": HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE
        },
    ]

    try:
        response = model.generate_content(
            prompt,
            generation_config=generation_config,
            safety_settings=safety_settings
        )
        
        # Process response and extract image search queries
        processed_response = process_response_with_images(response.text)
        return processed_response
        
    except Exception as e:
        print(f"Error generating content with Gemini: {str(e)}")
        return minimal_error_response()

def process_response_with_images(response_text):
    """Process the response and extract image search queries"""
    slides = []
    current_slide = None
    current_content = []
    
    for line in response_text.split('\n'):
        line = line.strip()
        if not line:
            continue
            
        if line.startswith('[Slide'):
            if current_slide:
                current_slide['content'] = current_content
                slides.append(current_slide)
            current_slide = {'title': '', 'content': [], 'image_query': ''}
            current_content = []
            
        elif line.startswith('Title:'):
            if current_slide:
                current_slide['title'] = line.split(':', 1)[1].strip()
                
        elif line.startswith('Image Search Query:'):
            if current_slide:
                current_slide['image_query'] = line.split(':', 1)[1].strip()
                
        elif line.startswith('-'):
            current_content.append(line)
    
    # Add the last slide
    if current_slide:
        current_slide['content'] = current_content
        slides.append(current_slide)
    
    return slides

def minimal_error_response():
    """Return a minimal valid response in case of errors"""
    return [
        {
            'title': 'Error in Content Generation',
            'content': ['- Could not generate presentation content.', '- Please try again.'],
            'image_query': 'error message'
        }
    ]

def parse_gemini_output(output):
    slides = []
    current_slide = None

    # Check if output is None or empty
    if not output:
        print("Error: Gemini output is empty or None")
        # Return a minimal valid structure with default content
        return [
            {'title': 'Error in Presentation Generation', 'content': ['- Could not generate presentation content.', '- Please check your input and try again.']},
            {'title': 'Index', 'content': []},
            {'title': 'Error Details', 'content': ['- The AI model returned an empty or invalid response.', '- This might be due to strict presentation rules that could not be applied.', '- Try with simpler or clearer presentation rules.']}
        ]

    # If output is already a list of slides, return it directly
    if isinstance(output, list):
        return output

    # Process string output
    for line in output.split('\n'):
        line = line.strip()
        if not line:  # Skip empty lines
            continue
            
        if line.startswith('[Slide'):
            if current_slide:
                slides.append(current_slide)
            current_slide = {'title': '', 'content': []}
        elif line.startswith('Title:'):
            if current_slide is None:  # Handle case where Title appears before [Slide]
                current_slide = {'title': '', 'content': []}
            current_slide['title'] = line.split(':', 1)[1].strip()
        elif line.startswith('-'):
            if current_slide is None:  # Handle case where bullet points appear before slide declaration
                current_slide = {'title': 'Untitled Slide', 'content': []}
            current_slide['content'].append(line)

    # Don't forget to add the last slide
    if current_slide:
        slides.append(current_slide)

    # Check if we have at least one slide
    if not slides:
        print("Error: No valid slides were parsed from Gemini output")
        return [
            {'title': 'Error in Presentation Generation', 'content': ['- Could not generate presentation content.', '- Please check your input and try again.']},
            {'title': 'Index', 'content': []},
            {'title': 'Error Details', 'content': ['- The AI model returned a response that could not be parsed into slides.', '- This might be due to presentation rules that were difficult to apply.', '- Try with simpler or clearer presentation rules.']}
        ]
        
    # Ensure we have at least a title slide and index slide
    if len(slides) < 2:
        if not slides[0]['title']:
            slides[0]['title'] = 'Presentation'
        slides.insert(1, {'title': 'Index', 'content': []})
    
    # Debug info
    print(f"Successfully parsed {len(slides)} slides")
    
    return slides

def generate_vba_code(slides, creator_name=None):
    # Print debugging information
    print("Debugging: Number of slides:", len(slides))
    for i, slide in enumerate(slides):
        print(f"Slide {i}: {slide['title']}")

    vba_code = f"""
Sub CreatePresentation()
    Dim ppt As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim tf As TextFrame
    Dim para As TextRange
    
    ' Create a new presentation
    Set ppt = Application.Presentations.Add

    ' Add title slide
    Set sld = ppt.Slides.Add(1, ppLayoutTitle)
    sld.Shapes.Title.TextFrame.TextRange.Text = "{slides[0]['title']}"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    
    ' Add creator name if provided
    If sld.Shapes.HasTitle Then
        Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 400, 600, 50)
        Set tf = shp.TextFrame
        tf.TextRange.Text = "Created by: {creator_name if creator_name else ''}"
        tf.TextRange.Font.Size = 14
        tf.TextRange.Font.Color.RGB = RGB(128, 128, 128)  ' Gray color
        tf.HorizontalAlignment = ppAlignCenter
    End If

    ' Add index slide
    Set sld = ppt.Slides.Add(2, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "Index"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.MarginLeft = 20
    tf.MarginRight = 20

    ' Add index content
"""

    # Add index content with each title on a new line
    for i, slide in enumerate(slides[2:], start=1):  # Skip title and index slides
        vba_code += f"""
    ' Add number
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "{i}."
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceAfter = 0
    para.ParagraphFormat.SpaceBefore = 6
    
    ' Add title on next line
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "{slide['title'].replace('"', '""')}"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Size = 14
    para.ParagraphFormat.LeftIndent = 20
    para.ParagraphFormat.SpaceAfter = 12
    para.ParagraphFormat.SpaceBefore = 0
"""

    # Add Conclusion with consistent formatting
    vba_code += """
    ' Add blank line before conclusion
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = ""
    para.ParagraphFormat.SpaceAfter = 6
    
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "Conclusion"
    para.ParagraphFormat.Bullet.Visible = False
    para.Font.Bold = True
    para.Font.Size = 14
    para.ParagraphFormat.SpaceBefore = 6
"""

    # Add content slides
    for index, slide in enumerate(slides[2:], start=3):  # Start from slide 3
        vba_code += f"""
    ' Add slide {index}
    Set sld = ppt.Slides.Add({index}, ppLayoutText)
    sld.Shapes.Title.TextFrame.TextRange.Text = "{slide['title'].replace('"', '""')}"
    sld.Shapes.Title.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 400)
    Set tf = shp.TextFrame
    tf.WordWrap = True
    tf.AutoSize = ppAutoSizeShapeToFitText
"""

        for point in slide['content']:
            vba_code += f"""
    Set para = tf.TextRange.Paragraphs.Add
    para.Text = "{point[1:].strip().replace('"', '""')}"
    para.ParagraphFormat.Bullet.Visible = True
    para.ParagraphFormat.Bullet.RelativeSize = 1
    para.Font.Size = 14
"""

    vba_code += """
End Sub
"""

    return vba_code

def main():
    input_folder = 'extract'  # Name of your input folder containing txt files
    output_file = 'create_presentation.vba'
    num_content_slides = 6  # Customize this number as needed

    if not os.path.exists(input_folder):
        print(f"Error: Input folder '{input_folder}' not found.")
        return

    content = read_input_files(input_folder)
    if not content:
        print("No text files found in the input folder.")
        return

    gemini_output = generate_outline_with_gemini(content, num_content_slides)
    slides = parse_gemini_output(gemini_output)
    vba_code = generate_vba_code(slides)
    
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write(vba_code)

    print(f"VBA code generated and saved to '{output_file}'")

if __name__ == "__main__":
    main()