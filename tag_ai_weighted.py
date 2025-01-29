import os
from dotenv import load_dotenv
import openpyxl
from datetime import datetime
from anthropic import Anthropic

# Create log file with timestamp
log_filename = f'data/tag_processing_ai_weighted_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
try:
    log_file = open(log_filename, 'w', encoding='utf-8')
except IOError:
    print("Error: Cannot create log file in data directory")
    exit()

def log_message(message):
    """Write message to both console and log file"""
    print(message)
    log_file.write(message + '\n')

# Initialize Anthropic client
try:
    # Load environment variables from .env file
    load_dotenv()
    
    # Get API key from environment variable
    anthropic_key = os.getenv('ANTHROPIC_API_KEY')
    if not anthropic_key:
        raise ValueError("ANTHROPIC_API_KEY environment variable not found")
        
    client = Anthropic(api_key=anthropic_key)
except Exception as e:
    log_message(f"Error initializing Anthropic client: {str(e)}")
    log_file.close()
    exit()

def analyze_with_ai(content, tags):
    """Use Claude to analyze content and suggest relevant tags"""
    try:
        prompt = f"""Given the following content and list of available tags, return ONLY the most relevant tags that match the content.
        Order them by relevance, most relevant first.
        Respond with just the tags, separated by commas, nothing else.
        
        Content: {content}
        Available tags: {', '.join(tags)}"""
        
        response = client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=150,
            temperature=0.3,
            messages=[{
                "role": "user",
                "content": prompt
            }]
        )
        suggested_tags = response.content[0].text.split(',')
        return [tag.strip() for tag in suggested_tags if tag.strip() in tags]
    except Exception as e:
        log_message(f"AI analysis error: {str(e)}")
        return []

# Load the XLSX file
workbook = openpyxl.load_workbook('data/input.xlsx')
worksheet = workbook.active

# Read tags from txt file
try:
    with open('data/tags.txt', 'r', encoding='utf-8') as file:
        tags = [line.strip() for line in file if line.strip()]
except FileNotFoundError:
    log_message("Error: tags.txt file not found in data directory")
    log_file.close()
    exit()

log_message(f"Found {len(tags)} tags to check against.")

# Check if tag columns exist, if not, add them
tag_columns = ['Tag 1', 'Tag 2', 'Tag 3', 'Tag 4', 'Tag 5']
existing_headers = [worksheet.cell(row=1, column=col).value for col in range(1, worksheet.max_column + 1)]

# Find the last column with content
last_content_column = 1
for col in range(1, worksheet.max_column + 1):
    has_content = False
    for row in range(1, worksheet.max_row + 1):
        if worksheet.cell(row=row, column=col).value is not None:
            has_content = True
            break
    if has_content:
        last_content_column = col
    else:
        break

log_message(f"Last column with content found at position {last_content_column}")

# Add missing tag column headers
for i, tag_header in enumerate(tag_columns):
    column = None
    # Check if header already exists
    if tag_header in existing_headers:
        column = existing_headers.index(tag_header) + 1
    else:
        # Add new header after the last column
        column = last_content_column + i + 1
        worksheet.cell(row=1, column=column, value=tag_header)
        log_message(f"Added header '{tag_header}' in column {column}")

# Iterate through each row in the XLSX file
for row in range(2, worksheet.max_row + 1):
    # Get the text content from all cells in the row
    row_content = ''
    for col in range(1, last_content_column + 1):
        cell_value = worksheet.cell(row=row, column=col).value
        if cell_value:
            row_content += str(cell_value) + ' '
    
    # Use both keyword matching and AI analysis
    keyword_matches = set()
    for tag in tags:
        if tag.lower() in row_content.lower():
            keyword_matches.add(tag)
    
    ai_matches = set(analyze_with_ai(row_content, tags))
    
    # Combine and prioritize matches
    # Give higher weight to keyword matches (they appear in both sets)
    weighted_tags = []
    for tag in keyword_matches.union(ai_matches):
        weight = 2 if tag in keyword_matches and tag in ai_matches else 1
        weighted_tags.append((tag, weight))
    
    # Sort by weight (descending) and then by AI suggestion order
    ai_order = {tag: idx for idx, tag in enumerate(ai_matches)}
    weighted_tags.sort(key=lambda x: (-x[1], ai_order.get(x[0], len(ai_matches))))
    
    matching_tags = [tag for tag, _ in weighted_tags]
    log_message(f"Row {row} - Keyword matches: {keyword_matches}")
    log_message(f"Row {row} - AI suggested tags: {ai_matches}")
    log_message(f"Row {row} - Final weighted order: {matching_tags[:5]}")
    
    # Add matching tags to the Tag columns
    for i, tag in enumerate(matching_tags[:5]):  # Limit to first 5 matching tags
        column = last_content_column + i + 1
        worksheet.cell(row=row, column=column, value=tag)
        log_message(f"Writing '{tag}' to row {row}, column {column}")

# Save the updated XLSX file
output_filename = 'data/output-with-tags-ai-weighted.xlsx'
workbook.save(output_filename)

log_message(f"Processing complete. Results saved in '{output_filename}'")
log_file.close()