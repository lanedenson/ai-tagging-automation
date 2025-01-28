# Tagging Automation

A Python script that automatically tags Excel rows by analyzing their content and matching them with predefined categories from a tags list.

## Description

This tool helps automate the process of categorizing Excel data by:
- Reading content from multiple columns in an Excel file
- Matching content against a predefined list of tags
- Assigning relevant tags to dedicated tag columns

## Prerequisites

- Python 3.x
- openpyxl library

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/tagging-automation.git
   cd tagging-automation
   ```

2. Install required dependencies:
   ```bash
   pip install openpyxl
   ```

## Usage

1. Place your Excel file in the `data` directory as `input.xlsx`
2. Ensure your tags are listed in `data/tags.txt` (one tag per line)
3. Run the script:
   ```bash
   python tag.py
   ```

The script will:
- Process all 27 columns of each row
- Match content against the tags list
- Assign up to 5 relevant tags per row
- Save results in `data/output-with-tags.xlsx`

## Tag Matching

The script performs basic case-insensitive substring matching. For example:
- If a row contains "Topic" anywhere in its content
- The corresponding "Topic" tag will be assigned to one of the tag columns

## Output

The script adds 5 new columns to the Excel file:
- Tag 1
- Tag 2
- Tag 3
- Tag 4
- Tag 5

Each row can have up to 5 tags assigned based on its content.

## Future Enhancements

This is the first pass of the tagging automation system. Future refinements will include:
- Integration with ChatGPT or other Large Language Models (LLMs) for more sophisticated content analysis
- Holistic understanding of context and meaning beyond simple substring matching
- Improved accuracy in tag assignments through AI-powered natural language processing
- Potential for automated tag suggestions based on content patterns

## Contributing

Feel free to submit issues and enhancement requests.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.