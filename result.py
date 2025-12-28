"""
Mismatch Analysis - Gradio Interface

This script provides a web interface for the mismatch analysis workflow.
Upload your input files and run the complete analysis.

Required inputs:
- ETOF file
- LC file(s)
- Rate Card file(s)
- Mismatch report
- Shipper name

Optional inputs:
- Order file
- Ignore rate card columns
- Include positive discrepancy
"""

import os
import sys
import shutil
import gradio as gr
from pathlib import Path
from datetime import datetime


def setup_python_path():
    """Setup Python path to include the script directory for imports."""
    try:
        if '__file__' in globals():
            script_dir = os.path.dirname(os.path.abspath(__file__))
        else:
            script_dir = os.getcwd()
        
        if script_dir and script_dir not in sys.path:
            sys.path.insert(0, script_dir)
            print(f"📁 Added to Python path: {script_dir}")
            
    except Exception as e:
        print(f"⚠️ Warning: Could not auto-detect script directory: {e}")


# Run setup when module is imported
setup_python_path()


def run_mismatch_analysis_gradio(
    etof_file,
    lc_files,
    rate_card_files,
    mismatch_file,
    shipper_name,
    order_file=None,
    ignore_rate_card_columns=None,
    include_positive_discrepancy=False
):
    """
    Main workflow for Gradio interface.
    Accepts uploaded files and user input; returns downloadable file and status messages.
    """
    # Capture status messages
    status_messages = []
    errors = []
    warnings = []
    
    def log_status(msg, level="info"):
        """Log status messages with different levels"""
        try:
            timestamp = datetime.now().strftime("%H:%M:%S")
        except:
            timestamp = ""
        
        formatted_msg = f"[{timestamp}] {msg}"
        status_messages.append(formatted_msg)
        
        if level == "error":
            errors.append(msg)
        elif level == "warning":
            warnings.append(msg)
        
        print(formatted_msg)
    
    def _handle_upload(uploaded, allow_multiple=False):
        """Handle file upload - convert Gradio file objects to paths."""
        if uploaded is None:
            return None if not allow_multiple else []
        
        if isinstance(uploaded, list):
            if not allow_multiple:
                return _handle_upload(uploaded[0] if uploaded else None, allow_multiple=False)
            result = []
            for item in uploaded:
                if item is None:
                    continue
                if hasattr(item, "name"):
                    result.append(item.name)
                elif isinstance(item, str):
                    result.append(item)
            return result if result else []
        
        if hasattr(uploaded, "name"):
            return uploaded.name
        if isinstance(uploaded, str):
            return uploaded
        return None if not allow_multiple else []
    
    # Convert file paths
    etof_path = _handle_upload(etof_file)
    lc_paths = _handle_upload(lc_files, allow_multiple=True)
    rate_card_paths = _handle_upload(rate_card_files, allow_multiple=True)
    mismatch_path = _handle_upload(mismatch_file)
    order_path = _handle_upload(order_file)
    
    # Validate required fields
    if not etof_path:
        error_msg = "❌ Error: ETOF File is required."
        log_status(error_msg, "error")
        return None, error_msg
    
    if not lc_paths:
        error_msg = "❌ Error: LC File(s) are required."
        log_status(error_msg, "error")
        return None, error_msg
    
    if not rate_card_paths:
        error_msg = "❌ Error: Rate Card File(s) are required."
        log_status(error_msg, "error")
        return None, error_msg
    
    if not mismatch_path:
        error_msg = "❌ Error: Mismatch Report is required."
        log_status(error_msg, "error")
        return None, error_msg
    
    if not shipper_name or not shipper_name.strip():
        error_msg = "❌ Error: Shipper Name is required."
        log_status(error_msg, "error")
        return None, error_msg
    
    log_status("✅ Validation passed. Starting workflow...", "info")
    
    # Create directories
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.getcwd()
    
    input_dir = os.path.join(script_dir, "input")
    output_dir = os.path.join(script_dir, "output")
    result_dir = os.path.join(script_dir, "result")
    partly_df_dir = os.path.join(script_dir, "partly_df")
    
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(result_dir, exist_ok=True)
    os.makedirs(partly_df_dir, exist_ok=True)
    
    # Copy uploaded files to input directory
    etof_filename = None
    lc_filenames = []
    rate_card_filenames = []
    mismatch_filename = None
    order_filename = None
    
    try:
        # ETOF file
        if etof_path:
            etof_filename = os.path.basename(etof_path)
            # Standardize name
            etof_ext = os.path.splitext(etof_filename)[1] or ".xlsx"
            etof_filename = f"etofs{etof_ext}"
            input_etof_path = os.path.join(input_dir, etof_filename)
            shutil.copy2(etof_path, input_etof_path)
            log_status(f"✓ ETOF file ready: {etof_filename}", "info")
        
        # LC files
        for idx, lc_path in enumerate(lc_paths):
            if lc_path:
                lc_filename = os.path.basename(lc_path)
                input_lc_path = os.path.join(input_dir, lc_filename)
                shutil.copy2(lc_path, input_lc_path)
                lc_filenames.append(lc_filename)
        log_status(f"✓ {len(lc_filenames)} LC file(s) ready", "info")
        
        # Rate card files
        for idx, rc_path in enumerate(rate_card_paths):
            if rc_path:
                rc_filename = os.path.basename(rc_path)
                input_rc_path = os.path.join(input_dir, rc_filename)
                shutil.copy2(rc_path, input_rc_path)
                rate_card_filenames.append(rc_filename)
        log_status(f"✓ {len(rate_card_filenames)} Rate Card file(s) ready", "info")
        
        # Mismatch file
        if mismatch_path:
            mismatch_filename = os.path.basename(mismatch_path)
            # Standardize name
            mismatch_ext = os.path.splitext(mismatch_filename)[1] or ".xlsx"
            mismatch_filename = f"mismatch{mismatch_ext}"
            input_mismatch_path = os.path.join(input_dir, mismatch_filename)
            shutil.copy2(mismatch_path, input_mismatch_path)
            log_status(f"✓ Mismatch file ready: {mismatch_filename}", "info")
        
        # Order file (optional)
        if order_path:
            order_filename = os.path.basename(order_path)
            input_order_path = os.path.join(input_dir, order_filename)
            shutil.copy2(order_path, input_order_path)
            log_status(f"✓ Order file ready: {order_filename}", "info")
            
    except Exception as e:
        error_msg = f"❌ Error copying files: {e}"
        log_status(error_msg, "error")
        return None, error_msg
    
    # Change to script directory
    original_cwd = os.getcwd()
    
    try:
        os.chdir(script_dir)
        
        # Import and run main workflow
        from main import run_workflow
        
        log_status("🚀 Starting mismatch analysis workflow...", "info")
        
        # Parse ignore columns
        ignore_cols = None
        if ignore_rate_card_columns and ignore_rate_card_columns.strip():
            ignore_cols = [col.strip() for col in ignore_rate_card_columns.split(',') if col.strip()]
        
        # Prepare LC files parameter
        lc_param = lc_filenames if len(lc_filenames) > 1 else (lc_filenames[0] if lc_filenames else None)
        
        # Run the workflow
        result_file = run_workflow(
            etof_file=etof_filename,
            lc_files=lc_param,
            rate_card_files=rate_card_filenames,
            mismatch_file=mismatch_filename,
            shipper_name=shipper_name.strip(),
            order_file=order_filename,
            ignore_rate_card_columns=ignore_cols,
            include_positive_discrepancy=include_positive_discrepancy
        )
        
        if result_file and os.path.exists(result_file):
            log_status(f"✅ Workflow completed successfully!", "info")
            log_status(f"📁 Result file: {result_file}", "info")
            
            # Copy to output directory for download
            output_result = os.path.join(output_dir, "Result.xlsx")
            shutil.copy2(result_file, output_result)
            
            final_file_path = output_result
        else:
            log_status("⚠️ Workflow completed but no result file generated", "warning")
            final_file_path = None
            
    except Exception as e:
        import traceback
        error_msg = f"❌ Workflow failed: {e}"
        log_status(error_msg, "error")
        log_status(f"Traceback: {traceback.format_exc()}", "error")
        final_file_path = None
        
    finally:
        os.chdir(original_cwd)
    
    # Prepare status summary
    status_summary = []
    status_summary.append("=" * 60)
    status_summary.append("WORKFLOW SUMMARY")
    status_summary.append("=" * 60)
    status_summary.append("")
    
    if final_file_path and os.path.exists(final_file_path):
        status_summary.append(f"✅ SUCCESS: Output file created")
        status_summary.append(f"   Location: {final_file_path}")
    else:
        status_summary.append(f"❌ Workflow did not complete successfully")
    
    status_summary.append("")
    
    if errors:
        status_summary.append(f"❌ ERRORS ({len(errors)}):")
        for i, error in enumerate(errors[:10], 1):
            status_summary.append(f"  {i}. {error}")
        if len(errors) > 10:
            status_summary.append(f"  ... and {len(errors) - 10} more errors")
        status_summary.append("")
    
    if warnings:
        status_summary.append(f"⚠️ WARNINGS ({len(warnings)}):")
        for i, warning in enumerate(warnings[:10], 1):
            status_summary.append(f"  {i}. {warning}")
        if len(warnings) > 10:
            status_summary.append(f"  ... and {len(warnings) - 10} more warnings")
        status_summary.append("")
    
    # Add key status messages
    key_messages = [msg for msg in status_messages if any(keyword in msg for keyword in 
                    ['✓', '✅', '❌', '⚠️', 'Error', 'Warning', 'SUCCESS', 'completed', 'failed', 'STEP'])]
    
    if key_messages:
        status_summary.append("Key Steps:")
        status_summary.append("-" * 60)
        status_summary.extend(key_messages[-20:])
    
    status_text = "\n".join(status_summary)
    
    return (final_file_path, status_text) if final_file_path and os.path.exists(final_file_path) else (None, status_text)


# ---- Gradio UI Definition ----
with gr.Blocks(title="Mismatch Analyzer", theme=gr.themes.Soft()) as demo:
    gr.Markdown("# 📊 Mismatch Analyzer")
    gr.Markdown("### Analyze cost mismatches against rate cards")
    
    with gr.Accordion("📖 Instructions & Information", open=False):
        gr.Markdown("""
        ## How to Use This Workflow
        
        ### Step 1: Upload Required Files
        - **ETOF File** (Required): Excel file containing ETOF shipment data (.xlsx)
        - **LC File(s)** (Required): XML files with LC data (can upload multiple)
        - **Rate Card File(s)** (Required): Excel file(s) containing rate card data (.xlsx)
        - **Mismatch Report** (Required): Excel file with mismatch data (.xlsx)
        - **Shipper Name** (Required): Enter the shipper identifier (e.g., "dairb")
        
        ### Step 2: Upload Optional Files
        - **Order Files Export** (Optional): Excel file with order data mapping
        
        ### Step 3: Configure Options
        - **Ignore Rate Card Columns**: Comma-separated column names to exclude
        - **Include Positive Discrepancy**: Check to include positive discrepancies in report
        
        ### Step 4: Run Workflow
        - Click "🚀 Run Analysis" button
        - Wait for processing to complete
        - Check the Status section for any issues
        - Download the Result.xlsx file when ready
        
        ## Workflow Steps
        1. **ETOF Processing**: Process ETOF shipment file
        2. **LC Processing**: Process LC XML file(s)
        3. **Rate Card Processing**: Process rate card file(s)
        4. **Mapping**: Create Order-LC-ETOF mapping
        5. **Vocabulary Mapping**: Map and rename columns
        6. **Matching**: Match shipments with rate card lanes
        7. **Mismatch Report**: Generate mismatch report
        8. **Rate Costs Analysis**: Analyze rate costs
        9. **Accessorial Costs**: Analyze accessorial costs
        10. **Mismatches Filing**: File mismatches
        11. **Conditions Checking**: Check conditions and add reasons
        12. **Cleaning**: Create final cleaned result
        
        ## Output File
        - **result.xlsx**: Final cleaned result with conditions checked
          - Separate tabs per carrier agreement
          - Pivot summary tabs
          - Color-coded by cost type
        """)
    
    gr.Markdown("---")
    gr.Markdown("### 📁 Required Files")
    
    with gr.Row():
        etof_input = gr.File(
            label="ETOF File (.xlsx) *Required",
            file_types=[".xlsx", ".xls"]
        )
        mismatch_input = gr.File(
            label="Mismatch Report (.xlsx) *Required",
            file_types=[".xlsx", ".xls"]
        )
        shipper_input = gr.Textbox(
            label="Shipper Name *Required",
            placeholder="e.g., dairb, apple, shipper"
        )
    
    with gr.Row():
        lc_input = gr.File(
            label="LC Files (.xml) *Required - can upload multiple",
            file_types=[".xml"],
            file_count="multiple"
        )
        rate_card_input = gr.File(
            label="Rate Card Files (.xlsx) *Required - can upload multiple",
            file_types=[".xlsx", ".xls"],
            file_count="multiple"
        )
    
    gr.Markdown("---")
    gr.Markdown("### 📁 Optional Files & Settings")
    
    with gr.Row():
        order_input = gr.File(
            label="Order Files Export (.xlsx) - Optional",
            file_types=[".xlsx", ".xls"]
        )
        ignore_columns_input = gr.Textbox(
            label="Ignore Rate Card Columns (Optional)",
            placeholder="Column1, Column2, Column3",
            info="Comma-separated column names to exclude from processing"
        )
        include_positive_input = gr.Checkbox(
            label="Include Positive Discrepancy",
            value=False,
            info="If checked, includes both positive and negative discrepancies"
        )
    
    gr.Markdown("---")
    
    launch_button = gr.Button("🚀 Run Analysis", variant="primary", size="lg")
    
    with gr.Row():
        output_file = gr.File(label="📥 Result.xlsx (Download)")
        status_output = gr.Textbox(
            label="📋 Status & Logs",
            lines=25,
            max_lines=40,
            interactive=False,
            placeholder="Workflow status and messages will appear here...",
            show_copy_button=True
        )
    
    def launch_workflow(etof_file, lc_files, rate_card_files, mismatch_file, 
                       shipper_name, order_file, ignore_columns, include_positive):
        try:
            result_file, status_text = run_mismatch_analysis_gradio(
                etof_file=etof_file,
                lc_files=lc_files,
                rate_card_files=rate_card_files,
                mismatch_file=mismatch_file,
                shipper_name=shipper_name,
                order_file=order_file,
                ignore_rate_card_columns=ignore_columns,
                include_positive_discrepancy=include_positive
            )
            return result_file, status_text
        except Exception as e:
            import traceback
            error_details = f"❌ CRITICAL ERROR:\n{str(e)}\n\nTraceback:\n{traceback.format_exc()}"
            return None, error_details
    
    launch_button.click(
        launch_workflow,
        inputs=[
            etof_input, lc_input, rate_card_input, mismatch_input,
            shipper_input, order_input, ignore_columns_input, include_positive_input
        ],
        outputs=[output_file, status_output]
    )


if __name__ == "__main__":
    # Create folders when program starts
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.getcwd()
    
    input_dir = os.path.join(script_dir, "input")
    output_dir = os.path.join(script_dir, "output")
    result_dir = os.path.join(script_dir, "result")
    partly_df_dir = os.path.join(script_dir, "partly_df")
    
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(result_dir, exist_ok=True)
    os.makedirs(partly_df_dir, exist_ok=True)
    
    print(f"📁 Input folder: {input_dir}")
    print(f"📁 Output folder: {output_dir}")
    print(f"📁 Result folder: {result_dir}")
    
    # Check if running in Colab
    in_colab = 'google.colab' in sys.modules
    
    if in_colab:
        print("🚀 Launching Gradio interface for Google Colab...")
        demo.launch(server_name="0.0.0.0", share=False, debug=False, show_error=True)
    else:
        print("🚀 Launching Gradio interface locally...")
        print(f"💡 Upload your files through the web interface")
        demo.launch(server_name="127.0.0.1", share=False)

