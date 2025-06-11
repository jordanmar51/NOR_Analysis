import pandas as pd
import os
import argparse
import sys
import re
from collections import defaultdict
from pathlib import Path
from datetime import datetime

def convert_csv_to_excel(input_folder='.', output_file=None, verbose=False):
    """
    Convert CSV files in a folder to a single Excel workbook with multiple sheets.
    
    Parameters:
    input_folder (str): Folder containing CSV files (default: current directory)
    output_file (str): Output Excel file path (default: auto-generated)
    verbose (bool): Whether to print verbose output
    
    Returns:
    tuple: (success: bool, output_path: str)
    """
    # Generate default output filename if not provided
    if output_file is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f'combined_excel_{timestamp}.xlsx'
    
    if verbose:
        print(f"Converting CSV files from {input_folder} to {output_file}")
    
    # Flag to track if any sheet is written
    sheets_written = False
    
    # Create an Excel writer object to write multiple sheets
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # Get all CSV files in the input folder
        csv_files = [f for f in os.listdir(input_folder) if f.endswith('.csv')]
        
        if not csv_files:
            print("No CSV files found in the current directory!")
            return False, None
        
        if verbose:
            print(f"Found {len(csv_files)} CSV files")
        
        for filename in csv_files:
            # Construct the file path
            file_path = os.path.join(input_folder, filename)
            
            try:
                # Read the CSV file into a DataFrame
                df = pd.read_csv(file_path)
                # Use the filename (without extension) as the sheet name
                sheet_name = os.path.splitext(filename)[0][:31]  # Sheet name limit is 31 chars
                # Save the DataFrame as a new sheet in the combined Excel file
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                # Mark that a sheet was written
                sheets_written = True
                if verbose:
                    print(f"Added sheet '{sheet_name}' with {len(df)} rows")
            except pd.errors.EmptyDataError:
                print(f"Skipping {filename}: EmptyDataError - No data found in the CSV file.")
            except pd.errors.ParserError:
                print(f"Skipping {filename}: ParserError - Could not parse the CSV file.")
            except Exception as e:
                print(f"Skipping {filename} due to unexpected error: {e}")
        
        # If no sheets were written, create a dummy sheet to avoid an empty workbook error
        if not sheets_written:
            pd.DataFrame({"Message": ["No valid data found"]}).to_excel(writer, sheet_name="NoData")
            return False, None
    
    print(f"Successfully created: {output_file}")
    if verbose:
        print("All valid sheets have been saved into the compiled Excel file")
    return True, output_file

def process_excel_file(input_path, output_path=None, verbose=False):
    """
    Process the Excel file by separating sheets by object_id and combining related data.
    
    Parameters:
    input_path (str): Path to original input Excel file
    output_path (str): Path for final output file (default: auto-generated)
    verbose (bool): Whether to print verbose output
    
    Returns:
    bool: True if successful, False otherwise
    """
    try:
        # Generate default output filename if not provided
        if output_path is None:
            input_stem = Path(input_path).stem
            output_path = f'{input_stem}_processed.xlsx'
        
        # Get absolute path to see where file is being saved
        output_abs_path = os.path.abspath(output_path)
        
        if verbose:
            print(f"\nProcessing Excel file: {input_path}")
            print(f"Output will be saved to: {output_abs_path}")
            print("Loading Excel file and processing sheets...")
        
        # Load the Excel file
        excel_data = pd.ExcelFile(input_path)
        
        # Dictionary to store separated data by sheet and object_id
        separated_data = {}
        has_object_id = False
        
        # Step 1: Separate data by object_id in memory
        for sheet_name in excel_data.sheet_names:
            data = pd.read_excel(input_path, sheet_name=sheet_name)
            if verbose:
                print(f"Loaded {len(data)} rows from sheet '{sheet_name}'")
            
            if 'object_id' in data.columns:
                has_object_id = True
                # Group data by 'object_id'
                for object_id, group in data.groupby('object_id'):
                    key = f"{sheet_name}_obj_{object_id}"
                    separated_data[key] = group.copy()
                    if verbose:
                        print(f"Separated {len(group)} rows for '{key}'")
            else:
                if verbose:
                    print(f"'object_id' column not found in sheet '{sheet_name}' - skipping")
        
        if not has_object_id:
            print("No sheets with 'object_id' column found. No processing needed.")
            return False
        
        # Step 2: Group sheets by their stem names and combine
        sheet_groups = defaultdict(list)
        for sheet_key in separated_data.keys():
            # Extract stem name (everything before _obj_)
            stem = re.split(r'_obj_[12]', sheet_key)[0]
            sheet_groups[stem].append(sheet_key)
        
        if verbose:
            print(f"\nCombining {len(sheet_groups)} sheet groups...")
        
        # Step 3: Create final Excel file
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for stem, sheet_list in sheet_groups.items():
                # Sort sheets to ensure _obj_1 comes before _obj_2
                sheet_list.sort()
                
                # Start with first sheet data
                combined_df = separated_data[sheet_list[0]].copy()
                
                # Add empty columns and second sheet if it exists
                if len(sheet_list) > 1:
                    # Add 4 empty columns
                    for i in range(4):
                        combined_df[f'empty_{i+1}'] = ''
                    
                    # Get second sheet data and reset its index
                    second_df = separated_data[sheet_list[1]].copy().reset_index(drop=True)
                    
                    # Reset index of first dataframe to ensure proper alignment
                    combined_df = combined_df.reset_index(drop=True)
                    
                    # Concatenate horizontally (side by side) starting from the same row
                    combined_df = pd.concat([combined_df, second_df], axis=1, ignore_index=False)
                
                # Write combined data to new sheet
                combined_df.to_excel(writer, sheet_name=stem, index=False)
                if verbose:
                    print(f"Created combined sheet '{stem}' with {len(combined_df)} rows")
        
        print(f"Processing completed! Output saved to: {output_abs_path}")
        return True
            
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(
        description="Convert CSV files to Excel and optionally process by object_id.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
This script is designed to be run from a directory containing CSV files.

Default behavior (no arguments):
  %%(prog)s
    - Converts all CSV files in current directory to Excel
    - Output: combined_excel_YYYYMMDD_HHMMSS.xlsx

Common usage:
  %%(prog)s -o myfile.xlsx          # Convert CSVs to myfile.xlsx
  %%(prog)s -p                       # Convert CSVs and process by object_id
  %%(prog)s -o report.xlsx -p -v     # Full processing with custom name and verbose output
  %%(prog)s -e existing.xlsx         # Process an existing Excel file
  %%(prog)s -e existing.xlsx -o new.xlsx  # Process existing file with custom output name

Examples:
  %%(prog)s                          # Quick convert in current directory
  %%(prog)s --output combined.xlsx   # Convert with custom filename
  %%(prog)s --process --verbose      # Full processing with details
  %%(prog)s -pv                      # Same as above (short form)
        """
    )
    
    # Input options
    parser.add_argument(
        '--excel-file', '-e',
        help='Process an existing Excel file instead of converting CSVs'
    )
    
    parser.add_argument(
        '--csv-folder', '-c',
        default='.',
        help='Folder containing CSV files (default: current directory)'
    )
    
    # Output argument
    parser.add_argument(
        '--output', '-o',
        help='Output file name (default: auto-generated with timestamp)'
    )
    
    # Process flag
    parser.add_argument(
        '--process', '-p',
        action='store_true',
        help='Also process the Excel file by object_id after conversion'
    )
    
    # Verbose flag
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='Show detailed progress information'
    )
    
    # Parse arguments
    args = parser.parse_args()
    
    # Check if we're processing an existing Excel file
    if args.excel_file:
        # Mode: Process existing Excel file
        if not os.path.exists(args.excel_file):
            print(f"Error: Excel file '{args.excel_file}' not found!")
            sys.exit(1)
        
        if args.verbose:
            print("Mode: Processing existing Excel file")
            print(f"Input file: {args.excel_file}")
            if args.output:
                print(f"Output file: {args.output}")
            print()
        
        # Process the Excel file
        success = process_excel_file(args.excel_file, args.output, args.verbose)
        if not success:
            sys.exit(1)
    
    else:
        # Mode: Convert CSV files (and optionally process)
        
        # Check if CSV folder exists
        if not os.path.exists(args.csv_folder):
            print(f"Error: Folder '{args.csv_folder}' not found!")
            sys.exit(1)
        
        # Check if there are CSV files
        csv_count = len([f for f in os.listdir(args.csv_folder) if f.endswith('.csv')])
        if csv_count == 0:
            print(f"Error: No CSV files found in '{args.csv_folder}'!")
            sys.exit(1)
        
        print(f"Found {csv_count} CSV files in '{args.csv_folder}'")
        
        if args.verbose:
            print("Mode: CSV to Excel conversion")
            if args.process:
                print("Will also process by object_id after conversion")
            print()
        
        # Convert CSVs to Excel
        # If processing is requested, use a temporary filename
        if args.process:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            temp_excel = f'temp_combined_{timestamp}.xlsx'
            success, excel_path = convert_csv_to_excel(args.csv_folder, temp_excel, args.verbose)
        else:
             success, excel_path = convert_csv_to_excel(args.csv_folder, args.output, args.verbose)
        
        if not success:
            print("Conversion failed!")
            sys.exit(1)
        
        # If process flag is set, also process the Excel file
        if args.process:
            print("\nStarting object_id processing...")
            
            # When processing, use the user's output name directly (no _processed suffix)
            final_output = args.output if args.output else None
            
            # Process the Excel file
            success = process_excel_file(excel_path, final_output, args.verbose)
            
            if not success:
                print("Processing failed!")
                sys.exit(1)
            
            # Delete the intermediate Excel file since we only want the processed version
            try:
                os.remove(excel_path)
                if args.verbose:
                    print(f"Removed intermediate file: {excel_path}")
            except Exception as e:
                print(f"Warning: Could not remove intermediate file {excel_path}: {e}")
    
    print("\nAll operations completed successfully!")
    
    # Final check for output files
    if args.output:
        if os.path.exists(args.output):
            abs_path = os.path.abspath(args.output)
            print(f"\nYour output file is located at:")
            print(f"  {abs_path}")
            print(f"  Size: {os.path.getsize(args.output)} bytes")
        else:
            print(f"\nWARNING: Expected output file '{args.output}' not found!")
            print("\nChecking for all Excel files in current directory:")
            xlsx_files = sorted([f for f in os.listdir('.') if f.endswith('.xlsx')])
            if xlsx_files:
                for f in xlsx_files:
                    print(f"  - {f}")
            else:
                print("  No .xlsx files found in current directory")

if __name__ == "__main__":
    main()