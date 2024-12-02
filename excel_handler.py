import pandas as pd
from utils.logging_config import setup_logging
from config import SUPPORTED_LANGUAGES, SOURCE_LANGUAGE
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils.cell import get_column_letter, column_index_from_string
import tempfile
import os
from pathlib import Path
from datetime import datetime
import re
import shutil
from openpyxl.styles import Color

# Set up logging to a file in a logs subdirectory
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)
logger = setup_logging("excel_processor")

class ExcelProcessor:
    def __init__(self, file):
        self.file = file
        self.df = None
        self.wb = None
        
    def clean_formatted_text(self, text_parts):
        """
        Clean and properly format text with tags, handling both bold and red text
        """
        result = []
        current_formatted_text = []
        
        # Check if the entire cell is red (all parts have red formatting)
        all_red = True
        for part in text_parts:
            if isinstance(part, TextBlock):
                if not (part.font and part.font.color and part.font.color.rgb == "FFFF0000"):
                    all_red = False
                    break
        
        # If entire cell is red, wrap the whole content
        if all_red:
            full_text = ''.join(str(part.text) if isinstance(part, TextBlock) else str(part) for part in text_parts)
            return f"<cr>{full_text}</cr>"
        
        # Otherwise, process parts individually
        for part in text_parts:
            if isinstance(part, TextBlock):
                text = str(part.text)
                is_bold = part.font and part.font.b
                is_red = part.font and part.font.color and part.font.color.rgb == "FFFF0000"  # Red color in RGB
                
                # If text is just whitespace, add it directly
                if text.isspace():
                    result.append(text)
                    continue
                    
                # Split text but preserve spaces using regex
                words = re.split(r'(\s+)', text)
                
                for word in words:
                    if word.isspace():  # Handle spacing
                        if current_formatted_text:
                            current_formatted_text.append(word)
                        else:
                            result.append(word)
                    elif word:  # Handle non-empty words
                        if is_bold and is_red:
                            result.append(f"<cfr>{word}</cfr>")  # Combined bold and red
                        elif is_bold:
                            result.append(f"<cf>{word}</cf>")    # Bold only
                        elif is_red:
                            result.append(f"<cr>{word}</cr>")    # Red only
                        else:
                            result.append(word)
            else:
                result.append(str(part))
        
        # Clean up multiple spaces and ensure proper spacing around tags
        text = ''.join(result)
        
        # Fix spacing around tags with punctuation and sentence awareness
        for tag in ['<cf>', '<cr>', '<cfr>']:
            # Add space before tag if preceded by non-space, non-punctuation, and not at start
            text = re.sub(r'(?<!^)(?<![\s\.,!?\)])(' + re.escape(tag) + ')', r' \1', text)
        
        for tag in ['</cf>', '</cr>', '</cfr>']:
            # Add space after tag if followed by non-space, non-punctuation
            text = re.sub(r'(' + re.escape(tag) + r')(?![\s\.,!?\)])(?!$)', r'\1 ', text)
        
        return text

    def read_excel(self, skip_first_row=True):
        """
        Read Excel file and immediately process formatting into tags
        """
        try:
            # Create temporary file for openpyxl
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                # Handle both file objects and file paths
                if isinstance(self.file, (str, Path)):
                    # If it's a path, copy the file
                    shutil.copy(str(self.file), tmp.name)
                else:
                    # If it's a file object, write its contents
                    self.file.seek(0)
                    tmp.write(self.file.read())
                tmp_path = tmp.name
            
            # Load with openpyxl first to get formatting
            self.wb = load_workbook(tmp_path, data_only=True, rich_text=True)
            ws = self.wb.active
            
            # Process formatting immediately
            formatted_cells = {}
            logger.info("Processing cell formatting...")
            
            for row_idx, row in enumerate(ws.iter_rows(min_row=2 if skip_first_row else 1)):
                for cell in row:
                    if isinstance(cell.value, CellRichText):
                        # Use the new clean_formatted_text method
                        formatted_text = self.clean_formatted_text(cell.value)
                        if any(tag in formatted_text for tag in ['<cf>', '<cr>', '<cfr>']):
                            logger.debug(f"Found rich text in cell {cell.coordinate}: {formatted_text}")
                            formatted_cells[cell.coordinate] = formatted_text
                    elif cell.font:
                        # Check for whole-cell formatting
                        text = str(cell.value)
                        if text.strip():  # Only process non-empty cells
                            if cell.font.color and cell.font.color.rgb == "FFFF0000":
                                # Whole cell is red
                                formatted_cells[cell.coordinate] = f"<cr>{text}</cr>"
                                logger.debug(f"Found red cell {cell.coordinate}: {text}")
                            elif cell.font.b:
                                # Whole cell is bold
                                formatted_cells[cell.coordinate] = f"<cf>{text}</cf>"
                                logger.debug(f"Found bold cell {cell.coordinate}: {text}")
            
            # Now read with pandas
            self.df = pd.read_excel(tmp_path, header=1 if skip_first_row else 0)
            
            # Apply formatted text to DataFrame
            for coord, text in formatted_cells.items():
                # Extract column letter and row number correctly
                col_letter = ''.join(c for c in coord if c.isalpha())
                row_num = int(''.join(c for c in coord if c.isdigit()))
                
                # Convert to zero-based index
                col_idx = column_index_from_string(col_letter) - 1
                row_idx = row_num - (2 if skip_first_row else 1) - 1
                
                # Apply the formatted text if indices are valid
                if row_idx >= 0 and row_idx < len(self.df) and col_idx >= 0 and col_idx < len(self.df.columns):
                    logger.debug(f"Applying formatted text at row {row_idx}, col {col_idx}: {text}")
                    self.df.iloc[row_idx, col_idx] = text
            
            # Save debug Excel file with tags
            debug_output = Path("debug_output")
            debug_output.mkdir(exist_ok=True)
            debug_file = debug_output / f"debug_tagged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.df.to_excel(debug_file, index=False)
            logger.info(f"Saved debug Excel file with tags: {debug_file}")
            
            # Create verification CSV with original and tagged text
            verification_data = []
            for row_idx, row in enumerate(ws.iter_rows(min_row=2 if skip_first_row else 1)):
                for cell in row:
                    if cell.coordinate in formatted_cells:
                        verification_data.append({
                            'Cell': cell.coordinate,
                            'Original': str(cell.value),
                            'Tagged': formatted_cells[cell.coordinate]
                        })
            
            if verification_data:
                verification_df = pd.DataFrame(verification_data)
                csv_file = debug_output / f"tag_verification_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                verification_df.to_csv(csv_file, sep=';', index=False, encoding='utf-8-sig')
                logger.info(f"Saved tag verification CSV: {csv_file}")
            
            return self.df
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}", exc_info=True)
            raise
        finally:
            if 'tmp_path' in locals():
                try:
                    os.unlink(tmp_path)
                except:
                    pass

    def get_cell_formatting(self, row_idx, col_name):
        """Get formatting information for a specific cell"""
        if self.wb is None:
            return None
            
        ws = self.wb.active
        # Adjust row index based on header configuration
        excel_row = row_idx + 3  # Adjust for 0-based index and header rows
        
        # Find column letter from column name
        col_letter = None
        for idx, cell in enumerate(ws[1], 1):  # Assuming headers are in row 1
            if cell.value == col_name:
                col_letter = cell.column_letter
                break
                
        if col_letter:
            cell = ws[f"{col_letter}{excel_row}"]
            logger.info(f"Checking formatting for cell {col_letter}{excel_row}")
            
            # Handle rich text
            if isinstance(cell.value, CellRichText):
                logger.info(f"Found rich text formatting")
                return cell.value
            elif cell.font and (cell.font.bold or cell.font.color):
                # Create rich text for cells with direct formatting
                rich_text = CellRichText()
                
                # Convert RGB color to Color object if needed
                color = None
                if cell.font.color:
                    if isinstance(cell.font.color, Color):
                        color = cell.font.color
                    else:
                        # Create a new Color object from the RGB value
                        rgb_value = cell.font.color.rgb if hasattr(cell.font.color, 'rgb') else None
                        if rgb_value:
                            color = Color(rgb=rgb_value)
                
                inline_font = InlineFont(
                    b=cell.font.bold,
                    color=color
                )
                rich_text.append(TextBlock(inline_font, str(cell.value)))
                return rich_text
            
        return None

    def detect_languages(self):
        """Detect language columns based on supported language codes"""
        if self.df is None:
            self.read_excel()
        
        detected_languages = []
        for col in self.df.columns:
            # Check if column name contains a language code
            for lang_code in SUPPORTED_LANGUAGES.keys():
                if lang_code in str(col):  # Convert to string to handle any non-string column names
                    detected_languages.append(lang_code)
                    logger.info(f"Detected language column: {lang_code} in '{col}'")
                    break
        
        if not detected_languages:
            logger.warning("No supported language columns detected")
        
        return detected_languages

    def validate_source_language(self):
        """Validate that source language (English GB) exists in the Excel file"""
        languages = self.detect_languages()
        if SOURCE_LANGUAGE not in languages:
            logger.error(f"Source language {SOURCE_LANGUAGE} not found in Excel file")
            raise ValueError(f"Source language {SOURCE_LANGUAGE} not found in Excel file")
        return True

    def get_available_languages(self):
        """Get list of available languages in the Excel file"""
        return self.detect_languages()