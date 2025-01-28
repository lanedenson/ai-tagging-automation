import openpyxl

# Load the XLSX file
workbook = openpyxl.load_workbook('data/jcd-exports.xlsx')
worksheet = workbook.active

# Read tags from txt file
try:
    with open('data/tags.txt', 'r', encoding='utf-8') as file:
        tags = [line.strip() for line in file if line.strip()]
except FileNotFoundError:
    print("Error: tags.txt file not found in data directory")
    exit()

print(f"Found {len(tags)} tags to check against.")

# Iterate through each row in the XLSX file
for row in range(2, worksheet.max_row + 1):
    # Combine content from all columns for checking
    row_content = []
    for col in range(1, 28):  # Check columns 1-27
        cell_value = worksheet.cell(row=row, column=col).value
        if cell_value:
            row_content.append(str(cell_value))
    
    if row_content:  # Check if we found any content
        # Combine all content and convert to lowercase
        combined_content = ' '.join(row_content).lower()
        
        # Analyze the content and determine relevant tags
        relevant_tags = []
        for tag in tags:
            if tag.lower() in combined_content:
                relevant_tags.append(tag)
        
        if relevant_tags:  # Debug print
            print(f"Row {row}: Found tags: {relevant_tags}")
        
        # Add tags to the Tag columns (columns 28-32 for Tag 1-5)
        for i, tag in enumerate(relevant_tags[:5]):
            column = 28 + i  # Tag columns start after the existing 27 columns
            worksheet.cell(row=row, column=column, value=tag)
            print(f"Writing '{tag}' to row {row}, column {column}")  # Debug print

# Save the updated XLSX file
workbook.save('data/jcd-exports-with-tags.xlsx')

print("Processing complete. Results saved in 'jcd-exports-with-tags.xlsx'")