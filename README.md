# Tagging Automation

A Python script that automatically tags Excel rows by analyzing their content and matching them with predefined categories from a tags list.

## Description

This tool helps automate the process of categorizing Excel data by:
- Reading content from multiple columns in an Excel file
- Matching content against a predefined list of tags using both keyword matching and AI analysis
- Assigning and prioritizing relevant tags to dedicated tag columns

## Prerequisites

- Python 3.x
- openpyxl library
- anthropic library
- Claude API key

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/tagging-automation.git
   cd tagging-automation
   ```

2. Install required dependencies:
   ```bash
   pip install openpyxl anthropic
   ```

3. Set up your environment:
   - Create a `data` directory in the project root
   - Place your Claude API key in `data/anthropic_key.txt`

## Usage

1. Place your Excel file in the `data` directory as `input.xlsx`
2. Ensure your tags are listed in `data/tags.txt` (one tag per line)
3. Run the script:
   ```bash
   python tag.py
   ```

The script will:
- Process all columns with data for each row
- Perform both keyword matching and AI analysis
- Assign up to 5 relevant tags per row, ordered by relevance
- Generate a detailed log file of the tagging process
- Save results in `data/output-with-tags.xlsx`

## Tag Matching

The script uses a sophisticated dual-matching system:

1. Keyword Matching
   - Performs case-insensitive substring matching
   - Identifies tags that directly appear in the content

2. AI-Powered Analysis
   - Uses Claude AI to analyze content context and meaning
   - Suggests relevant tags ordered by relevance
   - Provides more nuanced tag matching beyond simple substring matching

3. Tag Prioritization
   - Tags found by both methods (keyword and AI) receive a higher weight (2)
   - Tags found by only one method receive a lower weight (1)
   - Within each weight category, AI-suggested order is preserved
   - Ensures the 5 most relevant tags are selected for each row

## Output

The script produces two files:
1. `output-with-tags.xlsx`: The original Excel file with 5 new columns:
   - Tag 1
   - Tag 2
   - Tag 3
   - Tag 4
   - Tag 5

2. A timestamped log file (`tag_processing_YYYYMMDD_HHMMSS.log`) containing:
   - Processing details
   - Keyword matches for each row
   - AI-suggested tags
   - Final weighted tag selections

## Error Handling

The script includes robust error handling for:
- Missing input files
- Missing API key
- File access issues
- AI analysis errors
- Invalid tag formats

## Future Enhancements

Future refinements will include:
- Fine-tuning of the tag weighting system
- Additional AI models for comparison and validation
- Custom weighting rules for specific tag categories
- User-configurable priority settings
- Integration with domain-specific taxonomies
- Batch processing capabilities
- Support for additional file formats

## Contributing

Feel free to submit issues and enhancement requests.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.