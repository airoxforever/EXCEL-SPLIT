import sys
if len(sys.argv) == 1:
    sys.argv.extend(["--server.port", "8502"])

import streamlit as st
import pandas as pd
from pathlib import Path
from xliff_handler import XliffHandler
from excel_handler import ExcelProcessor
import logging
import os
from config import SUPPORTED_LANGUAGES, SOURCE_LANGUAGE
from utils.logging_config import setup_logging
import glob

# Configure logging
logger = setup_logging("app")

st.set_page_config(
    page_title="Excel-to-XLIFF Converter",
    page_icon="🌐",
    layout="wide"
)

DEFAULT_XLIFF_PATH = Path("W:/1/png")

def get_base_folder_from_file(file_path):
    """
    Get the base folder path from any XLIFF file path
    Example: if file is 'C:/translations/de_DE/file.xlf',
    returns 'C:/translations'
    """
    path = Path(file_path)
    # Go up one level from the language folder
    return path.parent.parent

def scan_xliff_folders(base_path):
    """
    Scan for XLIFF files in language-specific subfolders
    Returns a dict of {language_code: file_path}
    """
    xliff_files = {}
    base_path = Path(base_path)
    logger.info(f"Scanning for XLIFF files in: {base_path}")
    
    # Debug: Check if base path exists and is accessible
    logger.debug(f"Base path exists: {base_path.exists()}")
    logger.debug(f"Base path is directory: {base_path.is_dir()}")
    
    # Create a mapping of hyphenated to underscore codes (case insensitive)
    hyphen_to_underscore = {
        code.replace('_', '-').lower(): code for code in SUPPORTED_LANGUAGES.keys()
    }
    
    # Special case for Norwegian
    hyphen_to_underscore['nbno'] = 'no_NO'
    
    # Debug: Print the mapping
    logger.debug(f"Hyphen to underscore mapping: {hyphen_to_underscore}")
    
    # Debug: List all subfolders
    all_subfolders = list(base_path.glob('*-*'))
    logger.debug(f"Found subfolders: {[f.name for f in all_subfolders]}")
    
    # Scan all subfolders
    for lang_folder in base_path.glob('*-*'):  # Look for folders with hyphen
        logger.debug(f"Checking folder: {lang_folder}")
        
        if not lang_folder.is_dir():
            logger.debug(f"Skipping {lang_folder} - not a directory")
            continue
            
        hyphen_code = lang_folder.name.lower()  # Convert to lowercase for matching
        logger.debug(f"Hyphen code (lowercase): {hyphen_code}")
        
        underscore_code = hyphen_to_underscore.get(hyphen_code)
        logger.debug(f"Mapped to underscore code: {underscore_code}")
        
        if underscore_code and underscore_code not in ['en_GB', 'en_US']:
            # Look for .xlf files in this language folder
            xlf_files = list(lang_folder.glob('translation_*.xlf'))
            logger.debug(f"Found XLF files in {lang_folder}: {[f.name for f in xlf_files]}")
            
            # Filter out .sdlxliff files
            xlf_files = [f for f in xlf_files if not f.name.endswith('.sdlxliff')]
            logger.debug(f"After filtering .sdlxliff: {[f.name for f in xlf_files]}")
            
            if xlf_files:
                xliff_files[underscore_code] = xlf_files[0]  # Take the first XLIFF file found
                logger.info(f"Found XLIFF file for {underscore_code}: {xlf_files[0]}")
            else:
                logger.warning(f"No XLIFF file found for language {underscore_code}")
    
    # Debug: Final results
    logger.debug(f"Total XLIFF files found: {len(xliff_files)}")
    logger.debug(f"Found languages: {list(xliff_files.keys())}")
    
    return xliff_files

def main():
    st.title("Excel-to-XLIFF Converter")
    
    # Move all settings to sidebar
    st.sidebar.header("Settings")
    
    # Add source column configuration
    source_col = st.sidebar.text_input(
        "Source Column (e.g., 'E')",
        value="E",
        help="The column containing source text (en_GB/en_US). This column will never be modified."
    )
    
    # Add translation start position
    trans_start_pos = st.sidebar.text_input(
        "Translation Start Position (e.g., 'F3')",
        value="F3",
        help="The cell where translations begin. Content above this position is preserved."
    )
    
    # Store these in session state
    if 'source_col' not in st.session_state:
        st.session_state.source_col = source_col
    if 'trans_start_pos' not in st.session_state:
        st.session_state.trans_start_pos = trans_start_pos
    
    # Display current configuration
    st.sidebar.info(f"""
    Current Configuration:
    - Source Column: {st.session_state.source_col}
    - Translations Start: {st.session_state.trans_start_pos}
    """)
    
    # Excel header settings
    skip_first_row = st.sidebar.checkbox(
        "Skip first row (use second row as headers)", 
        value=True,
        help="Enable if your language headers are in the second row"
    )
    
    # Show supported languages in sidebar
    st.sidebar.header("Supported Languages")
    if st.sidebar.checkbox("Show supported languages", value=False):
        st.sidebar.write(f"Source Language: {SUPPORTED_LANGUAGES[SOURCE_LANGUAGE]}")
        st.sidebar.write("Target Languages:")
        for code, name in SUPPORTED_LANGUAGES.items():
            if code != SOURCE_LANGUAGE:
                st.sidebar.write(f"- {name} ({code})")
    
    # Debug mode setting
    debug_mode = st.sidebar.checkbox(
        "Debug Mode", 
        value=False,
        help="Show additional debugging information"
    )
    
    # Move comment settings before segmentation settings
    st.sidebar.header("Comment Settings")
    comment_column = st.sidebar.text_input(
        "Comment Column",
        value="C",
        help="Column containing comments that will be included in XLIFF files"
    )

    # Store comment column in session state
    if 'comment_column' not in st.session_state:
        st.session_state.comment_column = comment_column

    # Add information about SDL Studio compatibility
    st.sidebar.info("""
    💡 Note about comments:
    Comments are stored in XLIFF files as <note> elements with 'from="reviewer"'.
    To ensure SDL Studio compatibility, you may need to:
    1. Use a compatible XLIFF filter
    2. Enable comment handling in Studio settings
    3. Consider using Studio's "XLIFF 1.2" filter settings
    """)

    # Then continue with segmentation settings...
    st.sidebar.header("Segmentation Settings")
    enable_splitting = st.sidebar.checkbox(
        "Enable automatic segment splitting",
        value=False,  # Default to disabled
        help="When enabled, long segments will be automatically split into smaller ones"
    )

    # Store in session state
    if 'enable_splitting' not in st.session_state:
        st.session_state.enable_splitting = enable_splitting

    # Initialize default values for segmentation settings
    if 'min_segment_length' not in st.session_state:
        st.session_state.min_segment_length = 5
    if 'max_unsplit_length' not in st.session_state:
        st.session_state.max_unsplit_length = 70

    # Update the existing segmentation settings to only show when splitting is enabled
    if st.session_state.enable_splitting:
        min_segment_length = st.sidebar.number_input(
            "Minimum segment length (characters)",
            value=st.session_state.min_segment_length,
            min_value=1,
            help="Segments shorter than this will be merged with previous segment"
        )

        max_unsplit_length = st.sidebar.number_input(
            "Maximum length without splitting",
            value=st.session_state.max_unsplit_length,
            min_value=10,
            help="Segments shorter than this will never be split"
        )
        
        # Update session state with new values
        st.session_state.min_segment_length = min_segment_length
        st.session_state.max_unsplit_length = max_unsplit_length

    # Create tabs for different operations
    tab1, tab2 = st.tabs(["Excel to XLIFF", "XLIFF to Excel"])
    
    with tab1:
        st.header("Convert Excel to XLIFF")
        # File upload for Excel
        excel_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'], key="excel_upload")
        
        if excel_file:
            try:
                logger.info(f"Processing Excel file: {excel_file.name}")
                processor = ExcelProcessor(excel_file)
                df = processor.read_excel(skip_first_row=skip_first_row)
                
                if df is not None:
                    st.success("Excel file loaded successfully!")
                    
                    # Always show formatting preview in a collapsible section
                    with st.expander("Preview Formatted Text", expanded=debug_mode):
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.subheader("Original Text")
                            st.dataframe(df.head())
                        
                        with col2:
                            st.subheader("Formatted Text (with tags)")
                            def highlight_tags(val):
                                if '<cfr>' in str(val):
                                    return 'background-color: #ffcccc'
                                elif '<cf>' in str(val):
                                    return 'background-color: yellow'
                                elif '<cr>' in str(val):
                                    return 'background-color: #ffdddd'
                                return ''
                            
                            styled_df = df.head().style.map(highlight_tags)
                            st.dataframe(styled_df)
                    
                    # Extract available languages
                    languages = processor.get_available_languages()
                    logger.info(f"Available languages: {languages}")
                    
                    if languages:
                        st.write("Detected languages:", ", ".join(languages))
                        
                        if st.button("Convert to XLIFF", key="convert_xliff"):
                            try:
                                xliff_handler = XliffHandler()
                                xliff_handler.update_settings(
                                    st.session_state.min_segment_length,
                                    st.session_state.max_unsplit_length
                                )
                                xliff_handler.processor = processor
                                xliff_handler.convert_to_xliff(
                                    df, 
                                    languages, 
                                    excel_path=excel_file.name
                                )
                                
                                # Show conversion summary
                                st.success("XLIFF files generated successfully!")
                                
                                # Display statistics
                                st.header("Conversion Statistics")
                                
                                # Create summary table
                                stats_data = []
                                for lang, stats in xliff_handler.processing_stats.items():
                                    stats_data.append({
                                        'Language': SUPPORTED_LANGUAGES.get(lang, lang),
                                        'Total Segments': stats['total_segments'],
                                        'Split Attempts': stats['split_attempts'],
                                        'Successful Splits': stats['successful_splits'],
                                        'With Comments': stats['segments_with_comments'],
                                        'Split Rate': f"{(stats['successful_splits']/stats['split_attempts']*100):.1f}%" if stats['split_attempts'] > 0 else "N/A"
                                    })
                                
                                if stats_data:
                                    stats_df = pd.DataFrame(stats_data)
                                    st.dataframe(stats_df, use_container_width=True)
                                    
                                    # Show overall summary
                                    total_segments = sum(s['total_segments'] for s in xliff_handler.processing_stats.values())
                                    total_splits = sum(s['successful_splits'] for s in xliff_handler.processing_stats.values())
                                    total_comments = sum(s['segments_with_comments'] for s in xliff_handler.processing_stats.values())
                                    
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Total Segments", total_segments)
                                    with col2:
                                        st.metric("Total Splits", total_splits)
                                    with col3:
                                        st.metric("Segments with Comments", total_comments)
                                    
                                    # Show detailed stats in expandable section
                                    with st.expander("Detailed Statistics", expanded=False):
                                        for lang, stats in xliff_handler.processing_stats.items():
                                            st.subheader(f"{SUPPORTED_LANGUAGES.get(lang, lang)}")
                                            st.write(f"""
                                            - Total segments: {stats['total_segments']}
                                            - Segments eligible for splitting: {stats['split_attempts']}
                                            - Successfully split segments: {stats['successful_splits']}
                                            - Segments with comments: {stats['segments_with_comments']}
                                            """)
                                            
                                            # Add comment details if there are any
                                            if stats['segments_with_comments'] > 0:
                                                st.write("**Comment Details:**")
                                                comment_data = pd.DataFrame(stats['comment_details'])
                                                comment_data.columns = ['Segment #', 'Source Text', 'Comment']
                                                st.dataframe(comment_data, use_container_width=True)
                                                
                                                # Add information about comment column
                                                st.info(f"Comments were read from column {st.session_state.comment_column}")
                                
                                logger.info("XLIFF conversion completed")
                            except Exception as e:
                                st.error(f"Error during XLIFF conversion: {str(e)}")
                                logger.error(f"XLIFF conversion failed: {str(e)}", exc_info=True)
                    else:
                        st.warning("No language columns found in the Excel file")
                        logger.warning("No language columns detected in Excel file")
                        
            except Exception as e:
                st.error(f"Error processing Excel file: {str(e)}")
                logger.error(f"Excel processing failed: {str(e)}", exc_info=True)

    with tab2:
        st.header("Update Excel from XLIFF")
        
        # File upload for original Excel
        st.subheader("1. Upload Original Excel File")
        original_excel = st.file_uploader("Upload original Excel file", type=['xlsx', 'xls'], key="original_excel")
        
        if original_excel:
            try:
                # Parse translation start position
                col_letter = ''.join(filter(str.isalpha, st.session_state.trans_start_pos))
                row_num = int(''.join(filter(str.isdigit, st.session_state.trans_start_pos)))
                
                # Add this information to the display
                st.info(f"""
                Processing Configuration:
                - Source Column: {st.session_state.source_col} (Protected)
                - Translations Start: Column {col_letter}, Row {row_num}
                """)
                
                # First try the default path
                if 'xliff_scan_done' not in st.session_state:
                    st.session_state.xliff_scan_done = False
                    st.session_state.using_default_path = True
                
                if not st.session_state.xliff_scan_done:
                    if DEFAULT_XLIFF_PATH.exists():
                        st.info(f"Checking default path: {DEFAULT_XLIFF_PATH}")
                        st.info(f"Path exists: {DEFAULT_XLIFF_PATH.exists()}")
                        st.info(f"Is directory: {DEFAULT_XLIFF_PATH.is_dir()}")
                        
                        # List contents of default path
                        try:
                            contents = list(DEFAULT_XLIFF_PATH.glob('*-*'))
                            st.info(f"Found folders: {[f.name for f in contents if f.is_dir()]}")
                        except Exception as e:
                            st.error(f"Error listing directory: {str(e)}")
                        
                        found_xliffs = scan_xliff_folders(DEFAULT_XLIFF_PATH)
                        if found_xliffs:
                            st.session_state.found_xliffs = found_xliffs
                            st.session_state.base_folder = DEFAULT_XLIFF_PATH
                            st.session_state.xliff_scan_done = True
                            st.success(f"Found XLIFF files in default location: {DEFAULT_XLIFF_PATH}")
                        else:
                            st.warning("No XLIFF files found in default location")
                            st.session_state.using_default_path = False
                    else:
                        st.warning("Default XLIFF folder not found")
                        st.session_state.using_default_path = False
                
                # If default path didn't work, show manual input
                if not st.session_state.using_default_path:
                    st.subheader("2. Enter XLIFF Folder Path")
                    custom_path = st.text_input(
                        "Enter the path to your XLIFF folder",
                        help="Enter the full path to the folder containing language subfolders"
                    )
                    
                    if custom_path and st.button("Scan Folder", key="scan_custom"):
                        try:
                            custom_path = Path(custom_path)
                            if not custom_path.exists():
                                st.error("Folder not found! Please check the path.")
                            else:
                                found_xliffs = scan_xliff_folders(custom_path)
                                if found_xliffs:
                                    st.session_state.found_xliffs = found_xliffs
                                    st.session_state.base_folder = custom_path
                                    st.session_state.xliff_scan_done = True
                                else:
                                    st.error("No XLIFF files found in the specified folder")
                        except Exception as e:
                            st.error(f"Error scanning folder: {str(e)}")
                            logger.error(f"Folder scan failed: {str(e)}", exc_info=True)
                
                # Show found files summary if we have any
                if hasattr(st.session_state, 'found_xliffs') and st.session_state.found_xliffs:
                    st.subheader("Found Translations")
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write("✅ Available translations:")
                        for lang_code, file_path in st.session_state.found_xliffs.items():
                            st.write(f"- {SUPPORTED_LANGUAGES[lang_code]}")
                    
                    with col2:
                        st.write("❌ Missing translations:")
                        missing_langs = [
                            lang for lang in SUPPORTED_LANGUAGES.keys()
                            if lang not in st.session_state.found_xliffs.keys() and lang not in ['en_GB', 'en_US']
                        ]
                        for lang_code in missing_langs:
                            st.write(f"- {SUPPORTED_LANGUAGES[lang_code]}")
                
                # Show the update button only if we have both Excel and XLIFF files
                if original_excel and hasattr(st.session_state, 'found_xliffs') and st.session_state.found_xliffs:
                    if st.button("Update Excel with Translations", key="update_excel"):
                        try:
                            # Save original Excel temporarily
                            temp_dir = Path("temp_uploads")
                            temp_dir.mkdir(exist_ok=True)
                            excel_path = temp_dir / original_excel.name
                            with open(excel_path, 'wb') as f:
                                f.write(original_excel.getvalue())
                            
                            # Process files with new parameters
                            xliff_handler = XliffHandler()
                            result_path = xliff_handler.xliff_to_excel(
                                excel_path, 
                                st.session_state.base_folder,
                                source_col=st.session_state.source_col,
                                trans_start_pos=st.session_state.trans_start_pos
                            )
                            
                            # Provide download link for the updated file
                            with open(result_path, 'rb') as f:
                                st.download_button(
                                    label="Download Updated Excel File",
                                    data=f,
                                    file_name=result_path.name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            
                            st.success(f"Excel file updated successfully with {len(st.session_state.found_xliffs)} translations!")
                            
                        except Exception as e:
                            st.error(f"Error updating Excel file: {str(e)}")
                            logger.error(f"Excel update failed: {str(e)}", exc_info=True)
                        finally:
                            # Cleanup temporary files
                            import shutil
                            shutil.rmtree(temp_dir, ignore_errors=True)
                            
            except Exception as e:
                st.error(f"Error processing Excel file: {str(e)}")
                logger.error(f"Excel processing failed: {str(e)}", exc_info=True)

if __name__ == "__main__":
    logger.info("Application started")
    main()