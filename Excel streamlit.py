import streamlit as st
import pandas as pd
import io
from typing import Optional, Tuple
import traceback

def load_excel_file(uploaded_file) -> Optional[pd.DataFrame]:
    """
    Load an Excel file and return a DataFrame.
    Supports both .xlsx and .xls formats.
    """
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        elif uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file, engine='xlrd')
        else:
            st.error("Unsupported file format. Please upload .xlsx or .xls files only.")
            return None
        return df
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return None

def validate_file_size(uploaded_file, max_size_mb: int = 50) -> bool:
    """
    Validate file size to prevent memory issues.
    """
    file_size_mb = uploaded_file.size / (1024 * 1024)
    if file_size_mb > max_size_mb:
        st.error(f"File size ({file_size_mb:.1f} MB) exceeds maximum allowed size ({max_size_mb} MB)")
        return False
    return True

def fill_missing_data(complete_df: pd.DataFrame, blank_df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """
    Fill missing data STRICTLY by position so the result has EXACTLY the same
    rows (count, labels, and order) and columns (order) as blank_df.

    Steps:
    - Reset indices (positional alignment).
    - Align complete_df to blank_df's shape: same columns (order) and same row count.
    - Fill only where blank_df has NaN/None.
    - Restore original index from blank_df.
    Returns the filled DataFrame and the count of cells that were filled.
    """
    # Preserve original structure
    original_index = blank_df.index
    original_columns = blank_df.columns
    original_rows = len(blank_df)
    original_cols = len(original_columns)

    # Count missing cells before processing
    original_missing = blank_df.isna().sum().sum()

    # Work on positional copies (ignore original index labels during fill)
    b_pos = blank_df.reset_index(drop=True)
    c_pos = complete_df.reset_index(drop=True)

    # Ensure columns match: keep only blank_df's columns, preserve order.
    # Missing columns in complete_df become NaN and simply won't fill anything.
    c_pos = c_pos.reindex(columns=original_columns)

    # Ensure same number of rows as blank_df: trim extra, pad missing with NaN
    c_pos = c_pos.reindex(range(original_rows))

    # Fill only where blank is NA, using positional alignment
    filled_pos = b_pos.where(~b_pos.isna(), c_pos)

    # Restore original index and columns exactly
    filled_df = filled_pos.copy()
    filled_df.columns = original_columns
    filled_df.index = original_index

    # Count how many cells were filled
    filled_count = int(((b_pos.isna()) & (c_pos.notna())).sum().sum())

    # Final checks (shape must match exactly)
    assert len(filled_df) == original_rows, f"Output must have {original_rows} rows, got {len(filled_df)}"
    assert len(filled_df.columns) == original_cols, f"Output must have {original_cols} columns, got {len(filled_df.columns)}"

    return filled_df, filled_count

def create_download_link(df: pd.DataFrame, filename: str) -> bytes:
    """
    Create a downloadable Excel file from DataFrame.
    """
    # Use a temporary file approach that's compatible with current pandas version
    import tempfile
    import os
    
    # Create temporary file
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
        temp_path = tmp_file.name
    
    try:
        # Write to temporary file
        df.to_excel(temp_path, index=False, sheet_name='Filled_Data', engine='openpyxl')
        
        # Read the file back as bytes
        with open(temp_path, 'rb') as f:
            excel_data = f.read()
        
        return excel_data
    
    finally:
        # Clean up temporary file
        if os.path.exists(temp_path):
            os.unlink(temp_path)

def main():
    st.set_page_config(
        page_title="Excel Data Filler",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Excel Data Filler")
    st.markdown("Fill missing data in Excel files by combining two datasets")
    
    # Create two columns for file uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📋 Complete Dataset")
        st.markdown("Upload the Excel file with complete data")
        complete_file = st.file_uploader(
            "Choose complete dataset file",
            type=['xlsx', 'xls'],
            key="complete",
            help="This file should contain the complete data that will be used to fill missing values"
        )
    
    with col2:
        st.subheader("📝 Dataset with Blanks")
        st.markdown("Upload the Excel file with missing data to be filled")
        blank_file = st.file_uploader(
            "Choose file with blanks",
            type=['xlsx', 'xls'],
            key="blank",
            help="This file contains missing data that will be filled using the complete dataset"
        )
    
    # Process files when both are uploaded
    if complete_file is not None and blank_file is not None:
        # Validate file sizes
        if not validate_file_size(complete_file) or not validate_file_size(blank_file):
            return
        
        try:
            # Show progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("Loading complete dataset...")
            progress_bar.progress(25)
            complete_df = load_excel_file(complete_file)
            if complete_df is None:
                return
            
            status_text.text("Loading dataset with blanks...")
            progress_bar.progress(50)
            blank_df = load_excel_file(blank_file)
            if blank_df is None:
                return
            
            status_text.text("Processing and filling missing data...")
            progress_bar.progress(75)
            
            # Fill missing data (STRICT positional, exact row count of blank file)
            filled_df, filled_count = fill_missing_data(complete_df, blank_df)
            
            progress_bar.progress(100)
            status_text.text("Processing complete!")
            
            # Display results
            st.success(f"✅ Successfully filled {filled_count} missing cells!")
            
            # Create tabs for data preview
            tab1, tab2, tab3 = st.tabs(["📋 Complete Dataset", "📝 Original with Blanks", "✅ Filled Result"])
            
            with tab1:
                st.subheader("Complete Dataset Preview")
                st.dataframe(complete_df, use_container_width=True)
                st.info(f"Shape: {complete_df.shape[0]} rows × {complete_df.shape[1]} columns")
            
            with tab2:
                st.subheader("Original Dataset with Blanks")
                st.dataframe(blank_df, use_container_width=True)
                st.info(f"Shape: {blank_df.shape[0]} rows × {blank_df.shape[1]} columns")
            
            with tab3:
                st.subheader("Filled Dataset")
                st.dataframe(filled_df, use_container_width=True)
                st.info(f"Shape: {filled_df.shape[0]} rows × {filled_df.shape[1]} columns")
                
                # Download button
                if filled_count > 0:
                    excel_data = create_download_link(filled_df, "filled_data.xlsx")
                    st.download_button(
                        label="📥 Download Filled Dataset",
                        data=excel_data,
                        file_name="filled_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Download the processed file with filled missing data"
                    )
                else:
                    st.info("No missing data was found to fill.")
            
            # Clear progress indicators
            progress_bar.empty()
            status_text.empty()
            
        except Exception as e:
            st.error(f"An error occurred during processing: {str(e)}")
            st.error("Please check your files and try again.")
            # Show detailed error for debugging
            with st.expander("Show detailed error information"):
                st.code(traceback.format_exc())
    
    # Instructions and help section
    with st.expander("ℹ️ How to use this application"):
        st.markdown("""
        ### Instructions:
        1. **Upload Complete Dataset**: Upload an Excel file (.xlsx or .xls) that contains the complete data
        2. **Upload Dataset with Blanks**: Upload an Excel file that has missing/blank cells you want to fill
        3. **Processing**: The application will automatically fill blank cells in the second file using values from the first file at matching positions (same row and column)
        4. **Preview**: Review the original files and the filled result in the preview tabs
        5. **Download**: Download the processed file with filled missing data
        
        ### File Requirements:
        - Supported formats: .xlsx and .xls
        - Maximum file size: 50 MB per file
        - Files should have similar structure for best results
        
        ### Your Requirement (GUARANTEED):
        - **Fill data in Sheet2 from Sheet1**
        - **Output must have exactly the same number of rows as Sheet2**
        - Sheet1 = Complete dataset (used as source for filling)
        - Sheet2 = File with missing data (determines output size)
        
        ### Technical Details:
        - Empty cells are identified as blank, NaN, or None values
        - Only cells that are truly empty will be filled
        - **Strict positional matching** (row-by-row, column order of Sheet2)
        - Extra rows from Sheet1 are ignored to preserve Sheet2 size
        """)

if __name__ == "__main__":
    main()