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
        output = io.BytesIO()
        source_col = next(col for col in df.columns if source_lang in col)
        target_col = next(col for col in df.columns if target_lang in col)
        
        bilingual_df = pd.DataFrame({
            'Source': df[source_col],
            'Target': df[target_col]
        })
        
        bilingual_df.to_excel(output, index=False)
        return output.getvalue()
    except Exception as e:
        logger.error(f"Error creating bilingual Excel: {e}")
        raise

def merge_translations(original_df, translations_dict):
    """Merge translations back into the original dataframe"""
    try:
        result_df = original_df.copy()
        
        for lang_code, translation_df in translations_dict.items():
            # Find the target column in the original dataframe
            target_col = next(col for col in result_df.columns if lang_code in col)
            # Update the column with translations
            result_df[target_col] = translation_df['Target'].values
            
        return result_df
    except Exception as e:
        logger.error(f"Error merging translations: {e}")
        raise

def main():
    st.title("Excel File Splitter & Merger")
    
    # Add detailed app description
    st.markdown("""
    ### 📚 About This App
    This application helps you manage multilingual Excel files by:
    1. **Splitting** a multilingual Excel file into separate bilingual files for translation
    2. **Merging** the translated files back into the original format
    
    All formatting from the original file is preserved throughout the process.
    
    ### 🔒 Privacy & Security
    - All files are processed locally in your browser
    - No data is stored or transmitted anywhere
    - Files are automatically deleted after processing
    """)
    
    # Create tabs
    tab1, tab2 = st.tabs(["Split Excel", "Merge Translations"])
    
    # Split Excel Tab
    with tab1:
        st.header("Split Multilingual Excel")
        st.markdown("""
        ### 📝 How it works:
        1. Upload your multilingual Excel file
        2. The app will automatically detect language columns
        3. Verify the detected languages and source column
        4. Generate bilingual files for translation
        """)
        
        uploaded_file = st.file_uploader(
            "Upload your multilingual Excel file",
            type=['xlsx', 'xls'],
            key="split_uploader"
        )
        
        if uploaded_file:
            try:
                excel_processor = ExcelProcessor(uploaded_file)
                
                # Preserve original formatting
                if not excel_processor.preserve_workbook_format(uploaded_file):
                    st.error("Failed to preserve original file formatting")
                    return
                
                # Get column information
                col_info = excel_processor.get_column_info()
                
                if col_info:
                    st.success(f"✅ Found language columns in row {col_info['header_row']}")
                    
                    # Display and allow correction of source column
                    source_col = col_info['source_column']
                    if source_col:
                        st.info(f"📌 Source language column detected: {source_col['header']} (Column {source_col['letter']})")
                    else:
                        st.warning("⚠️ Source language column (English GB) not automatically detected")
                    
                    # Allow manual source column selection
                    col_options = {f"{info['header']} (Column {info['letter']})": code 
                                 for code, info in col_info['columns'].items()}
                    
                    selected_source = st.selectbox(
                        "Verify or select source language column:",
                        options=list(col_options.keys()),
                        index=list(col_options.keys()).index(f"{source_col['header']} (Column {source_col['letter']})") if source_col else 0
                    )
                    
                    # Show available target languages
                    st.subheader("Target Languages")
                    available_languages = [
                        f"{info['language_name']} ({info['header']})"
                        for code, info in col_info['columns'].items()
                        if code != col_options[selected_source]
                    ]
                    
                    selected_languages = st.multiselect(
                        "Select target languages for splitting:",
                        options=available_languages,
                        default=available_languages
                    )
                    
                    if selected_languages and st.button("Generate Bilingual Files"):
                        try:
                            source_lang = col_options[selected_source]
                            
                            with st.spinner("Generating bilingual files..."):
                                zip_buffer = io.BytesIO()
                                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                                    for lang_display in selected_languages:
                                        # Extract language code from display name
                                        target_lang = next(
                                            code for code, info in col_info['columns'].items()
                                            if f"{info['language_name']} ({info['header']})" == lang_display
                                        )
                                        
                                        # Create bilingual file with preserved formatting
                                        new_wb = excel_processor.create_bilingual_file(source_lang, target_lang)
                                        
                                        # Save to zip
                                        excel_buffer = io.BytesIO()
                                        new_wb.save(excel_buffer)
                                        filename = f"{source_lang}-{target_lang}.xlsx"
                                        zf.writestr(filename, excel_buffer.getvalue())
                                
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            st.download_button(
                                label="📥 Download Bilingual Files (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name=f"bilingual_files_{timestamp}.zip",
                                mime="application/zip"
                            )
                            
                            st.success("✅ Bilingual files generated successfully!")
                            
                        except Exception as e:
                            st.error(f"Error generating files: {str(e)}")
                            logger.error(f"File generation error: {e}", exc_info=True)
                else:
                    st.error("""
                    ❌ Could not detect language columns in the file.
                    
                    Please ensure:
                    1. The file contains language codes (e.g., ENGB, FRFR) in the column headers
                    2. The language codes are within the first 10 rows of the file
                    3. The file follows the expected format
                    """)
            except Exception as e:
                st.error(f"Error processing file: {str(e)}")
                logger.error(f"File processing error: {e}", exc_info=True)
    
    # Merge Translations Tab
    with tab2:
        st.header("Merge Translations")
        
        col1, col2 = st.columns(2)
        
        with col1:
            original_file = st.file_uploader(
                "Upload original multilingual Excel file",
                type=['xlsx', 'xls'],
                key="merge_original_uploader"
            )
        
        with col2:
            translations_zip = st.file_uploader(
                "Upload ZIP file with translations",
                type=['zip'],
                key="merge_translations_uploader"
            )
        
        if original_file and translations_zip:
            try:
                # Process original file
                excel_processor = ExcelProcessor(original_file)
                original_df = excel_processor.read_excel()
                
                # Process translations
                translations_dict = {}
                with zipfile.ZipFile(translations_zip) as zf:
                    for filename in zf.namelist():
                        if filename.endswith('.xlsx'):
                            # Extract language code from filename
                            target_lang = filename.split('-')[1].split('.')[0]
                            
                            # Read translation file
                            with zf.open(filename) as f:
                                translation_df = pd.read_excel(f)
                                translations_dict[target_lang] = translation_df
                
                if translations_dict:
                    try:
                        # Initialize Excel processor with original file
                        excel_processor = ExcelProcessor(original_file)
                        
                        # Preserve original formatting
                        if not excel_processor.preserve_workbook_format(original_file):
                            st.error("Failed to preserve original file formatting")
                            return
                        
                        # Apply translations while preserving formatting
                        if not excel_processor.apply_translations_to_workbook(translations_dict):
                            st.error("Failed to apply translations")
                            return
                        
                        # Save to bytes buffer
                        output = io.BytesIO()
                        excel_processor.save_workbook(output)
                        
                        # Offer download
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        st.download_button(
                            label="📥 Download Merged Excel File",
                            data=output.getvalue(),
                            file_name=f"merged_translations_{timestamp}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # Show preview
                        st.subheader("Preview of Merged File")
                        preview_df = pd.read_excel(output)
                        st.dataframe(preview_df.head())
                        
                        st.success("✅ Translations merged successfully!")
                        
                    except Exception as e:
                        st.error(f"Error merging translations: {str(e)}")
                        logger.error(f"Translation merge error: {e}", exc_info=True)
                else:
                    st.warning("No translation files found in the ZIP!")
            except Exception as e:
                st.error(f"Error processing files: {str(e)}")
                logger.error(f"File processing error: {e}", exc_info=True)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"Application error: {str(e)}")
        logger.error(f"Application error: {e}", exc_info=True)
    finally:
        cleanup_temp_files()