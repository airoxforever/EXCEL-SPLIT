# Excel File Splitter and Merger

A Streamlit application for splitting multilingual Excel files into bilingual pairs and merging them back after translation.

## Features
- Split multilingual Excel files into bilingual pairs
- Merge translated bilingual files back into a multilingual Excel
- Secure in-memory processing
- No data storage - all files are processed temporarily
- Support for multiple language pairs
- Format preservation during processing

## Live Demo
Access the application at: [Streamlit Cloud Link - TBD]

## Usage
1. **Split Operation**
   - Upload your multilingual Excel file
   - Select target languages
   - Download the ZIP with bilingual files

2. **Merge Operation**
   - Upload your original Excel file
   - Upload ZIP with translated files
   - Download the merged Excel file

## Privacy & Security
- All files are processed in-memory
- No data is stored on servers
- Files are automatically cleaned up after processing
- Secure file handling throughout the process

## Local Development
```bash
# Clone the repository
git clone [repository-url]

# Install dependencies
pip install -r requirements.txt

# Run the application
streamlit run app.py
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[MIT](https://choosealicense.com/licenses/mit/) 
