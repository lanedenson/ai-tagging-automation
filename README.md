# Tagging Automation

A Python script that automatically tags Excel rows by analyzing their content and matching them with predefined categories from a tags list.

## Description

This tool helps automate the process of categorizing Excel data by:
- Reading content from multiple columns in an Excel file
- Matching content against a predefined list of tags using both keyword matching and AI analysis
- Assigning and prioritizing relevant tags to dedicated tag columns
- Supporting both simple keyword-based tagging (`tag.py`) and AI-enhanced weighted tagging (`tag_ai_weighted.py`)

## Prerequisites

- Python 3.x
- openpyxl library
- anthropic library (for AI-enhanced tagging)
- Claude API key (for AI-enhanced tagging)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/lanedenson/tagging-automation.git
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

The repository provides two tagging approaches:

### Simple Keyword Tagging (tag.py)
1. Place your Excel file in the `data` directory as `input.xlsx`
2. Ensure your tags are listed in `data/tags.txt` (one tag per line)
3. Run:
   ```bash
   python tag.py
   ```

### AI-Enhanced Weighted Tagging (tag_ai_weighted.py)
1. Follow the same setup steps as above
2. Ensure your Claude API key is in place
3. Run:
   ```bash
   python tag_ai_weighted.py
   ```

Both scripts will:
- Process all columns with data for each row
- Generate a detailed log file of the tagging process
- Save results in `data/output-with-tags-*.xlsx`

## Tag Matching Methods

### Simple Keyword Matching (tag.py)
- Performs case-insensitive substring matching
- Identifies tags that directly appear in the content
- Assigns tags in order of appearance
- Suitable for straightforward categorization needs

### AI-Enhanced Weighted Matching (tag_ai_weighted.py)
- Combines keyword matching with AI analysis
- Uses Claude AI to analyze content context and meaning
- Implements a sophisticated weighting system:
  - Tags found by both methods receive weight 2
  - Tags found by single method receive weight 1
  - Preserves AI-suggested ordering within weight categories
- Provides more nuanced and context-aware tagging

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
   - AI-suggested tags (for AI-enhanced tagging)
   - Final weighted tag selections (for AI-enhanced tagging)

## Performance Considerations

- Simple keyword tagging (`tag.py`) is faster and requires no API calls
- AI-enhanced tagging (`tag_ai_weighted.py`) provides more accurate results but:
  - Requires API access
  - Processes rows more slowly due to API calls
  - May incur API usage costs

## Error Handling

The script includes robust error handling for:
- Missing input files
- Missing API key (for AI-enhanced tagging)
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