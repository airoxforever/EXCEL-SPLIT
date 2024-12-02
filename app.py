import streamlit as st
import pandas as pd
import tempfile
import os
from pathlib import Path
import io
import zipfile
from datetime import datetime
from excel_handler import ExcelProcessor
from config import SUPPORTED_LANGUAGES, SOURCE_LANGUAGE
from utils.logging_config import setup_logging

# Configure logging
logger = setup_logging("streamlit_app")

# Page configuration
st.set_page_config(
    page_title="Excel File Splitter & Merger",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = []

def cleanup_temp_files():
    """Clean up any temporary files in session state"""
    for temp_file in st.session_state.processed_files:
        try:
            if os.path.exists(temp_file):
                os.remove(temp_file)
                logger.info(f"Cleaned up temporary file: {temp_file}")
        except Exception as e:
            logger.error(f"Error cleaning up file {temp_file}: {e}")
    st.session_state.processed_files = []

def create_bilingual_excel(df, source_lang, target_lang):
    """Create a bilingual Excel file in memory"""
    try:
        # Create in-memory Excel file
        output = io.BytesIO()
        
        # Select only the required columns
        source_col = next(col for col in df.columns if source_lang in col)
        target_col = next(col for col in df.columns if target_lang in col)
        
        # Create new dataframe with just source and target columns
        bilingual_df = pd.DataFrame({
            'Source': df[source_col],
            'Target': df[target_col]
        })
        
        # Save to in-memory buffer
        bilingual_df.to_excel(output, index=False)
        return output.getvalue()
        
    except Exception as e:
        logger.error(f"Error creating bilingual Excel: {e}")
        raise

def main():
    st.title("Excel File Splitter & Merger")
    
    # Add privacy notice
    with st.expander("📋 Privacy & Security Information"):
        st.info(
            "This application processes all files in-memory and does not store any data. "
            "Files are automatically deleted after processing. "
            "All operations are performed locally in your browser session."
        )
    
    # Create tabs
    tab1, tab2 = st.tabs(["Split Excel", "Merge Translations"])
    
    # Split Excel Tab
    with tab1:
        st.header("Split Multilingual Excel")
        
        uploaded_file = st.file_uploader(
            "Upload your multilingual Excel file",
            type=['xlsx', 'xls'],
            key="split_uploader"
        )
        
        if uploaded_file:
            try:
                # Process the uploaded file
                excel_processor = ExcelProcessor(uploaded_file)
                df = excel_processor.read_excel()
                
                # Get available languages
                available_languages = excel_processor.get_available_languages()
                
                if available_languages:
                    st.success(f"Found {len(available_languages)} languages in the file!")
                    
                    # Language selection
                    selected_languages = st.multiselect(
                        "Select target languages for splitting",
                        options=[SUPPORTED_LANGUAGES[lang] for lang in available_languages if lang != SOURCE_LANGUAGE],
                        default=[SUPPORTED_LANGUAGES[lang] for lang in available_languages if lang != SOURCE_LANGUAGE]
                    )
                    
                    if selected_languages and st.button("Generate Bilingual Files"):
                        try:
                            # Create ZIP file in memory
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                                # Process each selected language
                                for lang_name in selected_languages:
                                    # Get language code from name
                                    lang_code = next(code for code, name in SUPPORTED_LANGUAGES.items() if name == lang_name)
                                    
                                    # Create bilingual Excel
                                    excel_content = create_bilingual_excel(df, SOURCE_LANGUAGE, lang_code)
                                    
                                    # Add to ZIP
                                    filename = f"{SOURCE_LANGUAGE}-{lang_code}.xlsx"
                                    zf.writestr(filename, excel_content)
                            
                            # Offer ZIP download
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            st.download_button(
                                label="Download Bilingual Files (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name=f"bilingual_files_{timestamp}.zip",
                                mime="application/zip"
                            )
                            
                        except Exception as e:
                            st.error(f"Error generating files: {str(e)}")
                            logger.error(f"File generation error: {e}", exc_info=True)
                else:
                    st.warning("No supported languages found in the file!")
                    
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")
                logger.error(f"File processing error: {e}", exc_info=True)
    
    # Merge Translations Tab
    with tab2:
        st.header("Merge Translations")
        st.info("🚧 Merge functionality coming soon!")
        # We'll implement this in the next step
    
    # Cleanup on session end
    cleanup_temp_files()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
        logger.error("Application error", exc_info=True)
    finally:
        cleanup_temp_files()