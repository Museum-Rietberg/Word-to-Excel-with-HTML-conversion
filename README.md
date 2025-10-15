# Word to Excel with HTML conversion
A workaround Text to Table conversion for lack of an appropriate (headless) CMS

## Scripts Overview

### convertWordToExcel.py
Converts structured Word documents (see textSources/sample.docx) to Excel format with advanced text processing capabilities:

#### Text Formatting Features
- Preserves italics using HTML `<em>` tags
- Converts line breaks to HTML `<br>` tags
- Preserves explicit superscript formatting from Word
- Maintains list structures (both bullet points and numbered lists)

#### Smart Typography Processing
1. Superscript Detection and Formatting:
   - French century notation (e.g., "XIXe" → "XIX<sup>e</sup>")
   - English ordinals (e.g., "19th" → "19<sup>th</sup>")
   - Numerical measurements (e.g., "km2" → "km<sup>2</sup>")

2. Micro-typographic Rules:
   - French quotation marks spacing ("«" and "»")
   - Punctuation spacing (before semicolon, colon)
   - En dash spacing
   - Non-breaking spaces:
     * After ordinal superscripts (e.g., "XIX<sup>e</sup>&nbsp;siècle")
     * Around measurement units (km, m, cm, etc.)
     * Around multiplication symbol (×)
     * Before currency symbols (€, $, £, EUR, CHF, USD)
     * Before temperature, percentage, and degree symbols
     * With time units (h, min, s, Uhr, heures, hours)
     * In common abbreviations (z. B., d. h., i. e., etc.)

#### Excel Output
- Creates organized Excel workbooks
- Adjusts column widths automatically
- Enables text wrapping for content columns
- Preserves table structure from Word

### collectAllContentInExcel.py
Combines and processes content from multiple tables:

- Merges content from multiple input tables
- Maintains data structure and relationships
- Preserves formatting and special characters
- Creates a consolidated Excel output

## Configuration
### File Paths
- Input Word documents go in the `textSources` folder
- Processed Excel files are saved to the `output` folder
- Default paths can be modified in the script configuration:
  ```python
  # In convertWordToExcel.py
  INPUT_FILE = 'textSources/tabelle.docx'
  OUTPUT_FILE = 'output/tabelle.xlsx'
  ```

## Usage
1. Place your Word documents in the `textSources` folder
2. Configure input/output paths in the script if needed
3. Run the appropriate script:
   ```bash
   python convertWordToExcel.py   # For single document conversion
   python collectAllContentInExcel.py  # For combining multiple tables
   ```
4. Find the processed Excel files in the `output` folder
