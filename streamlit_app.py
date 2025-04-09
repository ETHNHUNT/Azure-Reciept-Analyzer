import streamlit as st
import os
import tempfile
import time
import json
import logging
from datetime import datetime
import pandas as pd
from io import BytesIO
import uuid
import glob
from azure_receipt_analyzer import AzureReceiptAnalyzer
from config import VISION_ENDPOINT, VISION_API_KEY, OUTPUT_DIR, validate_config

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Validate configuration
try:
    validate_config()
except ValueError as e:
    logger.error(str(e))
    st.error(str(e))
    st.stop()

# Set page config
try:
    st.set_page_config(
        page_title="Receipt Analyzer",
        page_icon="🧾",
        layout="wide",
        initial_sidebar_state="expanded"
    )
except Exception as e:
    logger.error(f"Error setting page config: {str(e)}")
    st.error("Error initializing the application. Please try refreshing the page.")

# Custom CSS for better UI
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        margin-top: 1rem;
    }
    .receipt-stats {
        padding: 1rem;
        background-color: #f0f2f6;
        border-radius: 0.5rem;
        margin: 1rem 0;
    }
    .stProgress .st-bo {
        background-color: #4CAF50;
    }
    .success-message {
        color: #4CAF50;
        padding: 0.5rem;
        border-radius: 0.3rem;
        margin: 1rem 0;
    }
    .warning-message {
        color: #FFA500;
        padding: 0.5rem;
        border-radius: 0.3rem;
        margin: 1rem 0;
    }
    .error-message {
        color: #FF0000;
        padding: 0.5rem;
        border-radius: 0.3rem;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Define functions
def create_analyzer():
    """Create and return an AzureReceiptAnalyzer instance"""
    try:
        logger.info("Creating AzureReceiptAnalyzer instance")
        return AzureReceiptAnalyzer(VISION_ENDPOINT, VISION_API_KEY)
    except Exception as e:
        logger.error(f"Error creating analyzer: {str(e)}")
        st.error(f"Failed to create analyzer: {str(e)}")
        st.stop()

def process_uploaded_receipts(files):
    """Process uploaded receipt files using Azure Document Intelligence"""
    if not files:
        return None
    
    # Create a temporary directory to store uploaded files
    temp_dir = tempfile.mkdtemp()
    st.session_state.temp_dir = temp_dir
    
    # Save files to temp directory
    file_paths = []
    skipped_files = []
    
    # Define size limit (500MB for images, 100MB for PDF as per Azure DI docs)
    # Using a slightly lower practical limit, e.g., 80MB for PDF, 450MB for images
    MAX_PDF_SIZE_MB = 80
    MAX_IMAGE_SIZE_MB = 450
    MAX_PDF_SIZE_BYTES = MAX_PDF_SIZE_MB * 1024 * 1024
    MAX_IMAGE_SIZE_BYTES = MAX_IMAGE_SIZE_MB * 1024 * 1024
    
    for file in files:
        file_size = len(file.getbuffer())
        file_ext = os.path.splitext(file.name)[1].lower()
        supported_formats = ['.jpg', '.jpeg', '.png', '.pdf', '.tif', '.tiff', '.bmp']

        # Check file size based on type
        limit_exceeded = False
        if file_ext == '.pdf' and file_size > MAX_PDF_SIZE_BYTES:
            skipped_files.append((file.name, f"File too large (max {MAX_PDF_SIZE_MB}MB for PDF)"))
            limit_exceeded = True
        elif file_ext != '.pdf' and file_size > MAX_IMAGE_SIZE_BYTES:
            skipped_files.append((file.name, f"File too large (max {MAX_IMAGE_SIZE_MB}MB for images)"))
            limit_exceeded = True
            
        if limit_exceeded:
            continue
            
        # Check file extension
        if file_ext not in supported_formats:
            skipped_files.append((file.name, f"Unsupported format (supported: {', '.join(supported_formats)})"))
            continue
        
        # Save valid file
        file_path = os.path.join(temp_dir, file.name)
        with open(file_path, "wb") as f:
            f.write(file.getbuffer())
        file_paths.append(file_path)
    
    # Show warning for skipped files
    if skipped_files:
        skip_msg = "The following files were skipped:\n"
        for name, reason in skipped_files:
            skip_msg += f"- {name}: {reason}\n"
        st.warning(skip_msg)
    
    if not file_paths:
        st.error("No valid files to process. Please upload valid receipt images.")
        return None
    
    # Create analyzer
    try:
        analyzer = create_analyzer()
    except Exception as e:
        st.error(f"Failed to create analyzer: {str(e)}")
        return None
    
    # Process receipts directly without using blob storage
    results = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    total_files = len(file_paths)
    
    # Add a counter for total files and processed files
    progress_counter = st.empty()
    progress_counter.text(f"Processing 0/{total_files} receipts")
    
    # Display each file being processed
    current_file_display = st.empty()
    
    for i, file_path in enumerate(file_paths):
        file_name = os.path.basename(file_path)
        status_text.text(f"Processing receipt {i+1}/{len(file_paths)}: {file_name}")
        current_file_display.image(file_path, caption=f"Processing: {file_name}", width=300)
        progress_counter.text(f"Processing {i+1}/{total_files} receipts")
        
        try:
            # Process the receipt directly from local file
            result = analyzer.analyze_local_receipt(file_path)
            if result:
                results.append(result)
                
                # Show success message
                if result.get("merchant", {}).get("name") or result.get("transaction", {}).get("total") or result.get("items"):
                    st.success(f"✅ Successfully analyzed {file_name}")
                else:
                    st.warning(f"⚠️ Limited data extracted from {file_name}")
        except Exception as e:
            st.error(f"Error processing {file_name}: {str(e)}")
            # Add empty result with the filename
            empty_result = analyzer._get_empty_result(image_id=file_name)
            results.append(empty_result)
        
        # Update progress
        progress_bar.progress((i + 1) / len(file_paths))
        
        # Add delay between processing to avoid rate limiting
        if i < len(file_paths) - 1:  # Skip delay after last file
            with st.spinner(f"Waiting before processing next receipt (Azure rate limiting)..."):
                time.sleep(15)  # Shorter wait time for UI but still respects limits
    
    # Clear displays after processing
    current_file_display.empty()
    status_text.text("Receipt processing complete!")
    progress_counter.text(f"Processed {total_files}/{total_files} receipts")
    
    # Save results
    if results:
        output_dir = OUTPUT_DIR
        os.makedirs(output_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_file = os.path.join(output_dir, f"receipt_results_{timestamp}.json")
        excel_file = os.path.join(output_dir, f"receipt_results_{timestamp}.xlsx")
        
        # Save JSON
        with open(json_file, 'w') as f:
            json.dump(results, f, indent=2)
            
        # Create Excel file using the analyzer method
        try:
            # Construct the full file path for Excel
            excel_file_path = os.path.join(output_dir, f"receipt_results_{timestamp}.xlsx")
            # Call with just 2 args (plus self implicitly)
            excel_path = analyzer._save_results_excel(results, excel_file_path)
            excel_file = excel_path if excel_path else None
        except Exception as e:
            st.error(f"Error creating Excel file: {str(e)}")
            excel_file = None
            
        return {
            "results": results,
            "json_file": json_file,
            "excel_file": excel_file
        }
    
    return None

def download_as_json(data):
    """Create a download button for JSON data"""
    json_str = json.dumps(data, indent=2)
    return st.download_button(
        label="Download JSON",
        data=json_str,
        file_name=f"receipt_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
        mime="application/json"
    )

def download_excel(file_path):
    """Create a download button for Excel file"""
    with open(file_path, "rb") as f:
        data = f.read()
    return st.download_button(
        label="Download Excel",
        data=data,
        file_name=os.path.basename(file_path),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def load_css():
    st.markdown("""
        <style>
        .main {
            padding: 2rem;
        }
        .stButton>button {
            width: 100%;
            margin-top: 1rem;
        }
        .receipt-stats {
            padding: 1rem;
            background-color: #f0f2f6;
            border-radius: 0.5rem;
            margin: 1rem 0;
        }
        </style>
    """, unsafe_allow_html=True)

def initialize_analyzer():
    """Initialize the Azure Receipt Analyzer with credentials."""
    # Try to get credentials from Streamlit secrets first
    endpoint = st.secrets.get("VISION_ENDPOINT", os.getenv("VISION_ENDPOINT"))
    api_key = st.secrets.get("VISION_API_KEY", os.getenv("VISION_API_KEY"))
    
    if not endpoint or not api_key:
        st.error("❌ Missing Azure credentials. Please set VISION_ENDPOINT and VISION_API_KEY in your environment or Streamlit secrets.")
        st.info("To set up credentials:")
        st.code("""
1. Create a .streamlit/secrets.toml file with:
    VISION_ENDPOINT = "your_endpoint_here"
    VISION_API_KEY = "your_api_key_here"
        
2. Or set environment variables:
    export VISION_ENDPOINT="your_endpoint_here"
    export VISION_API_KEY="your_api_key_here"
        """)
        st.stop()
    
    try:
        return AzureReceiptAnalyzer(endpoint, api_key)
    except Exception as e:
        st.error(f"❌ Error initializing Azure Receipt Analyzer: {str(e)}")
        st.stop()

def display_receipt_summary(results):
    """Display receipt analysis summary."""
    if not results:
        st.warning("No results to display")
        return

    try:
        # Generate summary report
        analyzer = initialize_analyzer()
        summary = analyzer.generate_summary_report(results)
        
        # Display summary statistics in columns
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("💰 Total Spending", f"${summary['total_spending']:.2f}")
        with col2:
            st.metric("🧾 Receipts Processed", summary['receipt_count'])
        with col3:
            st.metric("🏪 Unique Merchants", summary['merchant_count'])
        with col4:
            st.metric("📦 Total Items", summary['item_count'])

        # Display date range if available
        if summary['date_range']['start'] and summary['date_range']['end']:
            st.info(f"📅 Date Range: {summary['date_range']['start']} to {summary['date_range']['end']}")

        # Display spending by category
        if summary.get('categories'):
            st.subheader("📊 Spending by Category")
            df_categories = pd.DataFrame(
                [(k, v) for k, v in summary['categories'].items()],
                columns=['Category', 'Amount']
            ).sort_values('Amount', ascending=False)
            
            # Create a bar chart
            st.bar_chart(df_categories.set_index('Category'))
            
            # Show the data table
            st.dataframe(
                df_categories.style.format({'Amount': '${:,.2f}'})
                .set_properties(**{'text-align': 'left'})
                .set_table_styles([
                    {'selector': 'th', 'props': [('text-align', 'left')]},
                    {'selector': 'td', 'props': [('text-align', 'left')]}
                ])
            )

    except Exception as e:
        st.error(f"Error displaying summary: {str(e)}")

def save_results(results):
    """Save results to JSON and Excel files."""
    if not results:
        return None, None
    
    # Create timestamp for filenames
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Save JSON
    json_path = f"receipt_results_{timestamp}.json"
    with open(json_path, 'w') as f:
        json.dump(results, f, indent=2)
    
    # Save Excel
    analyzer = initialize_analyzer()
    # Construct the full file path for Excel
    excel_file_path = f"receipt_results_{timestamp}.xlsx"
    # Call with just 2 args (plus self implicitly)
    excel_path = analyzer._save_results_excel(results, excel_file_path)
    
    return json_path, excel_path

# Main app
def main():
    try:
        logger.info("Starting Receipt Analyzer application")
        load_css()
        st.title("🧾 Receipt Analyzer")
        st.write("Upload your receipts for analysis and budget tracking")
        
        # Initialize session state
        if "processed_result" not in st.session_state:
            st.session_state.processed_result = None
        if "temp_dir" not in st.session_state:
            st.session_state.temp_dir = None
        
        # File uploader
        st.subheader("Upload Receipts")
        uploaded_files = st.file_uploader(
            "Upload receipt images (JPG, PNG, PDF)",
            type=['jpg', 'jpeg', 'png', 'pdf'],
            accept_multiple_files=True
        )
        
        # Process button
        if uploaded_files:
            if st.button("Process Receipts"):
                with st.spinner("Processing receipts... This may take a while due to API rate limits."):
                    logger.info(f"Processing {len(uploaded_files)} files")
                    st.session_state.processed_result = process_uploaded_receipts(uploaded_files)
        
        # Display results if available
        if st.session_state.processed_result:
            logger.info("Displaying results")
            st.subheader("Results")
            
            results = st.session_state.processed_result["results"]
            json_file = st.session_state.processed_result["json_file"]
            excel_file = st.session_state.processed_result["excel_file"]
            
            st.success(f"Successfully processed {len(results)} receipt(s)")
            
            # Create tabs for different views
            tab1, tab2, tab3 = st.tabs(["Summary", "Receipt Details", "Downloads"])
            
            with tab1:
                # Generate and display summary report
                try:
                    display_receipt_summary(results)
                except Exception as e:
                    logger.error(f"Error generating summary report: {str(e)}")
                    st.error("Error generating summary report")
            
            with tab2:
                # Display receipt details
                for i, result in enumerate(results):
                    with st.expander(f"Receipt {i+1}"):
                        st.json(result)
            
            with tab3:
                # Download options
                st.subheader("Download Results")
                if json_file:
                    download_as_json(results)
                if excel_file:
                    download_excel(excel_file)
        
        # Add footer
        st.markdown("---")
        st.markdown(
            "Made with ❤️ using Streamlit and Azure Document Intelligence"
        )
    
    except Exception as e:
        logger.error(f"Error in main application: {str(e)}")
        st.error("An error occurred in the application. Please try refreshing the page.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"Fatal error in application: {str(e)}")
        st.error("A fatal error occurred. Please check the logs and try again.") 