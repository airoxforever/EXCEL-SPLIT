# Excel File Splitter and Merger

A Streamlit application for splitting multilingual Excel files into bilingual pairs and merging them back after translation.

## Core Functionality

### 1. Split Operation
- Input: Multilingual Excel file
- Process: Creates bilingual Excel files (source + target pairs)
- Output: Multiple bilingual Excel files named by language codes
- Example: `en-DE.xlsx` contains English in column A, German in column B

### 2. Merge Operation
- Input: 
  - Original multilingual Excel file
  - ZIP file containing translated bilingual Excel files
- Process: Updates original file with translations from bilingual files
- Output: Updated multilingual Excel file
- Safety: Creates backup copy of original file

## Technical Requirements

### Dependencies
```
streamlit
pandas
openpyxl
zipfile
pathlib
```

### File Structure
```
project/
├── app.py              # Main Streamlit application
├── excel_handler.py    # Excel processing logic
├── config.py          # Configuration and language mappings
├── utils/
│   └── logging_config.py
└── temp/              # Temporary file processing (auto-cleaned)
```

### Reusable Components
From existing project:
1. `excel_handler.py`:
   - Language detection
   - Excel reading/writing
   - Basic file operations
2. `config.py`:
   - Language mappings
   - Configuration settings
3. `utils/logging_config.py`:
   - Logging setup
   - Error handling

## Processing Flow

### Split Operation
1. User uploads multilingual Excel file
2. System:
   - Detects language columns
   - Identifies source language
   - Shows preview for user confirmation
3. After confirmation:
   - Creates bilingual Excel files
   - Names files using language codes
   - Offers ZIP download of all files

### Merge Operation
1. User uploads:
   - Original multilingual Excel file
   - ZIP file with translated files
2. System:
   - Creates backup of original file
   - Extracts ZIP to temporary folder
   - Matches files by language codes
   - Updates translations in original file
3. Offers download of updated file

## Security Considerations

### File Handling
- All files processed in temporary directory
- Immediate cleanup after processing
- No permanent storage of uploaded files
- Secure file naming and handling

### Error Handling
- Validation of file formats
- Language code verification
- Structural integrity checks
- Detailed error messages

## User Interface

### Split Tab
- File upload area
- Language detection preview
- Source language confirmation
- Processing status
- Download button for ZIP

### Merge Tab
- Original file upload
- ZIP file upload
- Processing status
- Download button for updated file

## Future Enhancements
- Support for monolingual files
- Custom column mapping
- Advanced error recovery
- Batch processing
