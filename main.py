"""
Main Orchestration Script

This script runs the complete mismatch analysis workflow:
1. Process input files (ETOF, LC, Rate Card)
2. Create mappings (Order-LC-ETOF)
3. Run vocabulary mapping and matching
4. Generate mismatch report and analyze costs
5. Check conditions and create final result

Required inputs:
- ETOF file
- LC file(s)
- Rate Card file(s)
- Mismatch report
- Shipper name

Optional inputs:
- Order file
- Ignore rate card columns
"""

import os
import sys
import shutil
from pathlib import Path
from datetime import datetime


def setup_folders():
    """Create input, output, partly_df, and result folders if they don't exist."""
    script_dir = Path(__file__).parent
    
    folders = {
        'input': script_dir / 'input',
        'output': script_dir / 'output',
        'partly_df': script_dir / 'partly_df',
        'result': script_dir / 'result'
    }
    
    for name, folder in folders.items():
        folder.mkdir(exist_ok=True)
    
    return folders


def log_step(step_num, message, level="info"):
    """Log a step with timestamp."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    prefix = {
        "info": "📄",
        "success": "✅",
        "warning": "⚠️",
        "error": "❌",
        "section": "="*60
    }.get(level, "  ")
    
    if level == "section":
        print(f"\n{prefix}")
        print(f"STEP {step_num}: {message}")
        print(f"{prefix}")
    else:
        print(f"[{timestamp}] {prefix} {message}")


def validate_inputs(etof_file, lc_files, rate_card_files, mismatch_file, shipper_name):
    """Validate that all required inputs are provided."""
    errors = []
    
    if not etof_file:
        errors.append("ETOF file is required")
    elif not os.path.exists(os.path.join("input", etof_file)):
        errors.append(f"ETOF file not found: input/{etof_file}")
    
    if not lc_files:
        errors.append("LC file(s) are required")
    else:
        lc_list = lc_files if isinstance(lc_files, list) else [lc_files]
        for lc_file in lc_list:
            if not os.path.exists(os.path.join("input", lc_file)):
                errors.append(f"LC file not found: input/{lc_file}")
    
    if not rate_card_files:
        errors.append("Rate Card file(s) are required")
    else:
        rc_list = rate_card_files if isinstance(rate_card_files, list) else [rate_card_files]
        for rc_file in rc_list:
            if not os.path.exists(os.path.join("input", rc_file)):
                errors.append(f"Rate Card file not found: input/{rc_file}")
    
    if not mismatch_file:
        errors.append("Mismatch report file is required")
    elif not os.path.exists(os.path.join("input", mismatch_file)):
        errors.append(f"Mismatch file not found: input/{mismatch_file}")
    
    if not shipper_name:
        errors.append("Shipper name is required")
    
    return errors


def run_workflow(
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
    Run the complete mismatch analysis workflow.
    
    Args:
        etof_file: Path to ETOF file relative to input/ folder (e.g., "etofs.xlsx")
        lc_files: Path(s) to LC file(s) relative to input/ folder (string or list)
        rate_card_files: Path(s) to Rate Card file(s) relative to input/ folder (string or list)
        mismatch_file: Path to mismatch report file relative to input/ folder (e.g., "mismatch.xlsx")
        shipper_name: Shipper identifier (e.g., "dairb")
        order_file: Optional path to order files export relative to input/ folder
        ignore_rate_card_columns: Optional list of column names to ignore in rate card
        include_positive_discrepancy: If True, include positive discrepancies; if False, only negative
    
    Returns:
        str: Path to the final result file, or None if workflow failed
    """
    print("\n" + "="*80)
    print("MISMATCH ANALYSIS WORKFLOW")
    print("="*80)
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Setup folders
    folders = setup_folders()
    
    # Change to script directory
    script_dir = Path(__file__).parent
    original_cwd = os.getcwd()
    os.chdir(script_dir)
    
    try:
        # Validate inputs
        log_step(0, "Validating inputs...", "info")
        validation_errors = validate_inputs(etof_file, lc_files, rate_card_files, mismatch_file, shipper_name)
        
        if validation_errors:
            for error in validation_errors:
                log_step(0, error, "error")
            return None
        
        log_step(0, "All inputs validated", "success")
        
        # Convert to lists for consistent handling
        lc_list = lc_files if isinstance(lc_files, list) else [lc_files]
        rc_list = rate_card_files if isinstance(rate_card_files, list) else [rate_card_files]
        
        # ========================================
        # STEP 1: ETOF Processing
        # ========================================
        log_step(1, "ETOF FILE PROCESSING", "section")
        try:
            from part1_etof_file_processing import process_etof_file, save_dataframe_to_excel
            log_step(1, f"Processing ETOF file: {etof_file}", "info")
            etof_df, etof_columns = process_etof_file(etof_file)
            save_dataframe_to_excel(etof_df, "etof_processed.xlsx")
            log_step(1, f"ETOF processed: {etof_df.shape[0]} rows, {etof_df.shape[1]} columns", "success")
        except Exception as e:
            log_step(1, f"ETOF processing failed: {e}", "error")
            raise
        
        # ========================================
        # STEP 2: LC Processing
        # ========================================
        log_step(2, "LC FILE PROCESSING", "section")
        try:
            from part2_lc_processing import process_lc_input, save_dataframe_to_excel
            log_step(2, f"Processing {len(lc_list)} LC file(s)...", "info")
            lc_input_param = lc_list if len(lc_list) > 1 else lc_list[0]
            lc_df, lc_columns = process_lc_input(lc_input_param, recursive=False)
            save_dataframe_to_excel(lc_df, "lc_processed.xlsx")
            log_step(2, f"LC processed: {lc_df.shape[0]} rows, {lc_df.shape[1]} columns", "success")
        except Exception as e:
            log_step(2, f"LC processing failed: {e}", "error")
            raise
        
        # ========================================
        # STEP 3: Rate Card Processing
        # ========================================
        log_step(3, "RATE CARD PROCESSING", "section")
        try:
            from part4_rate_card_processing import process_multiple_rate_cards
            log_step(3, f"Processing {len(rc_list)} Rate Card file(s)...", "info")
            rc_results = process_multiple_rate_cards(rc_list)
            log_step(3, f"Rate Cards processed: {len(rc_results)} files", "success")
        except Exception as e:
            log_step(3, f"Rate Card processing failed: {e}", "error")
            raise
        
        # ========================================
        # STEP 4: Order-LC-ETOF Mapping
        # ========================================
        log_step(4, "ORDER-LC-ETOF MAPPING", "section")
        try:
            from part7_optional_order_lc_etof_mapping import process_order_lc_etof_mapping
            log_step(4, "Creating Order-LC-ETOF mapping...", "info")
            lc_input_param = lc_list if len(lc_list) > 1 else lc_list[0]
            mapping_df, mapping_columns = process_order_lc_etof_mapping(
                lc_input_path=lc_input_param,
                etof_path=etof_file,
                order_files_path=order_file
            )
            log_step(4, f"Mapping completed: {mapping_df.shape[0]} rows", "success")
        except Exception as e:
            log_step(4, f"Mapping failed: {e}", "warning")
        
        # ========================================
        # STEP 5: Vocabulary Mapping
        # ========================================
        log_step(5, "VOCABULARY MAPPING", "section")
        try:
            from vocabular import map_and_rename_columns
            log_step(5, "Running vocabulary mapping...", "info")
            
            # Parse ignore columns
            ignore_cols = None
            if ignore_rate_card_columns:
                if isinstance(ignore_rate_card_columns, str):
                    ignore_cols = [col.strip() for col in ignore_rate_card_columns.split(',') if col.strip()]
                else:
                    ignore_cols = ignore_rate_card_columns
            
            lc_input_param = lc_list if len(lc_list) > 1 else lc_list[0]
            vocab_result = map_and_rename_columns(
                rate_card_file_path=rc_list[0] if rc_list else None,
                etof_file_path=etof_file,
                order_files_path=order_file,
                lc_input_path=lc_input_param,
                shipper_id=shipper_name,
                output_txt_path="column_mapping_results.txt",
                ignore_rate_card_columns=ignore_cols
            )
            log_step(5, "Vocabulary mapping completed", "success")
        except Exception as e:
            log_step(5, f"Vocabulary mapping failed: {e}", "warning")
        
        # ========================================
        # STEP 6: Matching
        # ========================================
        log_step(6, "MATCHING", "section")
        try:
            from matching import run_matching
            log_step(6, "Running matching process...", "info")
            matching_file = run_matching(rate_card_file_path=rc_list[0] if rc_list else None)
            log_step(6, f"Matching completed: {matching_file}", "success")
        except Exception as e:
            log_step(6, f"Matching failed: {e}", "warning")
        
        # ========================================
        # STEP 7: Mismatch Report
        # ========================================
        log_step(7, "MISMATCH REPORT", "section")
        try:
            from mismatch_report import main as mismatch_report_main
            log_step(7, f"Generating mismatch report (include_positive_discrepancy={include_positive_discrepancy})...", "info")
            mismatch_df = mismatch_report_main(include_positive_discrepancy=include_positive_discrepancy)
            log_step(7, f"Mismatch report generated: {len(mismatch_df)} rows", "success")
        except Exception as e:
            log_step(7, f"Mismatch report failed: {e}", "warning")
        
        # ========================================
        # STEP 8: Rate Costs Analysis
        # ========================================
        log_step(8, "RATE COSTS ANALYSIS", "section")
        try:
            from rate_costs import process_multiple_rate_cards as rate_costs_process
            log_step(8, "Analyzing rate costs...", "info")
            rate_costs_results = rate_costs_process(rc_list)
            log_step(8, f"Rate costs analyzed: {len(rate_costs_results)} files", "success")
        except Exception as e:
            log_step(8, f"Rate costs analysis failed: {e}", "warning")
        
        # ========================================
        # STEP 9: Accessorial Costs Analysis
        # ========================================
        log_step(9, "ACCESSORIAL COSTS ANALYSIS", "section")
        try:
            from rate_accesorial_costs import process_multiple_rate_cards as accessorial_process
            log_step(9, "Analyzing accessorial costs...", "info")
            accessorial_results = accessorial_process(rc_list)
            log_step(9, f"Accessorial costs analyzed: {len(accessorial_results)} files", "success")
        except Exception as e:
            log_step(9, f"Accessorial costs analysis failed: {e}", "warning")
        
        # ========================================
        # STEP 10: Mismatches Filing
        # ========================================
        log_step(10, "MISMATCHES FILING", "section")
        try:
            from mismacthes_filing import main as mismatches_filing_main
            log_step(10, "Filing mismatches...", "info")
            filing_result = mismatches_filing_main()
            log_step(10, "Mismatches filed", "success")
        except Exception as e:
            log_step(10, f"Mismatches filing failed: {e}", "warning")
        
        # ========================================
        # STEP 11: Conditions Checking
        # ========================================
        log_step(11, "CONDITIONS CHECKING", "section")
        try:
            from conditions_checking import main as conditions_main
            log_step(11, "Checking conditions...", "info")
            conditions_result = conditions_main(debug=False)
            log_step(11, f"Conditions checked: {len(conditions_result)} rows", "success")
        except Exception as e:
            log_step(11, f"Conditions checking failed: {e}", "error")
            raise
        
        # ========================================
        # STEP 12: Cleaning and Final Result
        # ========================================
        log_step(12, "CLEANING AND FINAL RESULT", "section")
        try:
            from cleaning import main as cleaning_main
            log_step(12, "Creating final result...", "info")
            result_path = cleaning_main()
            log_step(12, f"Final result created: {result_path}", "success")
        except Exception as e:
            log_step(12, f"Cleaning failed: {e}", "error")
            raise
        
        # ========================================
        # WORKFLOW COMPLETE
        # ========================================
        print("\n" + "="*80)
        print("WORKFLOW COMPLETED SUCCESSFULLY")
        print("="*80)
        print(f"Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"\nResult file: {result_path}")
        print("="*80)
        
        return str(result_path)
        
    except Exception as e:
        print("\n" + "="*80)
        print("WORKFLOW FAILED")
        print("="*80)
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None
        
    finally:
        # Restore original working directory
        os.chdir(original_cwd)


def main():
    """
    Main entry point - configure your input files here.
    """
    # ========================================
    # CONFIGURATION - Edit these values
    # ========================================
    
    # Required: ETOF file (in input/ folder)
    ETOF_FILE = "etofs.xlsx"
    
    # Required: LC file(s) (in input/ folder)
    # Can be a single file or list of files
    LC_FILES = "LC.xml"
    # LC_FILES = ["LC1.xml", "LC2.xml"]  # For multiple files
    
    # Required: Rate Card file(s) (in input/ folder)
    # Can be a single file or list of files
    RATE_CARD_FILES = ["rate.xlsx", "rate_3.xlsx"]
    # RATE_CARD_FILES = "rate.xlsx"  # For single file
    
    # Required: Mismatch report file (in input/ folder)
    MISMATCH_FILE = "mismatch.xlsx"
    
    # Required: Shipper name/identifier
    SHIPPER_NAME = "shipper"
    
    # Optional: Order files export (in input/ folder)
    ORDER_FILE = None  # Set to "Order_files_export.xlsx" if available
    
    # Optional: Columns to ignore in rate card processing
    IGNORE_RATE_CARD_COLUMNS = None  # Set to "Column1, Column2" or ["Column1", "Column2"]
    
    # Optional: Include positive discrepancy in mismatch report
    # True = include all non-zero discrepancies (positive and negative)
    # False = include only negative discrepancies
    INCLUDE_POSITIVE_DISCREPANCY = False
    
    # ========================================
    # RUN WORKFLOW
    # ========================================
    
    result = run_workflow(
        etof_file=ETOF_FILE,
        lc_files=LC_FILES,
        rate_card_files=RATE_CARD_FILES,
        mismatch_file=MISMATCH_FILE,
        shipper_name=SHIPPER_NAME,
        order_file=ORDER_FILE,
        ignore_rate_card_columns=IGNORE_RATE_CARD_COLUMNS,
        include_positive_discrepancy=INCLUDE_POSITIVE_DISCREPANCY
    )
    
    if result:
        print(f"\n✅ Success! Result file: {result}")
    else:
        print("\n❌ Workflow failed. Check the error messages above.")
    
    return result


if __name__ == "__main__":
    main()

