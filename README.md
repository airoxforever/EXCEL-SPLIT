# Excel File Splitter & Merger

A powerful tool for managing multilingual Excel files, designed for translators and localization teams.

## 🌟 Features

- **Smart Language Detection**: Automatically detects language columns in Excel files
- **Format Preservation**: Maintains original Excel formatting including:
  - Column widths
  - Cell colors and styles
  - Text formatting
- **Flexible Processing**: 
  - Automatically finds header rows
  - Handles various Excel layouts
  - Supports multiple language codes
- **Interactive Interface**:
  - Visual confirmation of detected languages
  - Manual override options
  - Clear error messages and guidance

## 📋 Requirements

- Python 3.7+
- Required packages listed in `requirements.txt`

## 🚀 Getting Started

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the application:
   ```bash
   streamlit run app.py
   ```

## 💡 How to Use

### Splitting Files

1. Upload your multilingual Excel file
2. The app will automatically detect:
   - Header row with language codes
   - Source language column (English GB)
   - Available target languages
3. Verify the detected source language
4. Select target languages for extraction
5. Download the generated bilingual files

### Merging Translations

1. Upload your original multilingual Excel file
2. Upload the ZIP file containing translated bilingual files
3. The app will:
   - Preserve all original formatting
   - Match translations to correct columns
   - Maintain original layout and styling
4. Download the merged file

## 📝 File Format Requirements

- Excel files (.xlsx or .xls)
- Language codes in column headers (e.g., ENGB, FRFR)
- Headers must be within first 10 rows
- Source language (English GB) must be present

## 🔍 Troubleshooting

If language detection fails:
1. Ensure language codes are present in column headers
2. Check if headers are within first 10 rows
3. Verify file format and encoding
4. Use manual column selection if needed

## 🔒 Privacy

- All processing is done locally
- No data is stored or transmitted
- Files are automatically cleaned up after processing
