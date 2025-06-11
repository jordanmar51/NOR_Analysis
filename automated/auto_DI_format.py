import pandas as pd
import os
import argparse

def main():
    parser = argparse.ArgumentParser(description="Combine paired _Obj1 and _Obj2 Excel sheets side-by-side.")
    parser.add_argument(
        "--input", "-i", required=True,
        help="Path to the input Excel file (e.g., Compiled_ExplorationBouts.xlsx)"
    )
    parser.add_argument(
        "--output", "-o", required=True,
        help="Path to the output Excel file (e.g., Processed_compiled_ExplorationBouts.xlsx)"
    )
    args = parser.parse_args()

    file_path = args.input
    output_path = args.output

    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file does not exist: {file_path}")

    excel_file = pd.ExcelFile(file_path)
    sheet_names = excel_file.sheet_names

    # Identify all unique prefixes with _Obj1 or _Obj2
    prefixes = set()
    for sheet in sheet_names:
        if sheet.endswith('_Obj1'):
            prefixes.add(sheet[:-5])
        elif sheet.endswith('_Obj2'):
            prefixes.add(sheet[:-5])
    sorted_prefixes = sorted(prefixes)

    combined_sheets = {}
    for prefix in sorted_prefixes:
        obj1_sheet = f"{prefix}_Obj1"
        obj2_sheet = f"{prefix}_Obj2"

        if obj1_sheet in sheet_names and obj2_sheet in sheet_names:
            obj1_data = pd.read_excel(excel_file, sheet_name=obj1_sheet)
            obj2_data = pd.read_excel(excel_file, sheet_name=obj2_sheet)

            if not obj1_data.empty and not obj2_data.empty:
                combined_data = pd.concat([obj1_data, obj2_data], axis=1)
            elif not obj1_data.empty:
                combined_data = obj1_data
            elif not obj2_data.empty:
                combined_data = obj2_data
            else:
                continue

            combined_sheets[prefix] = combined_data

    # Write output Excel
    with pd.ExcelWriter(output_path) as writer:
        wrote_sheet = False
        for prefix in sorted_prefixes:
            if prefix in combined_sheets and not combined_sheets[prefix].empty:
                combined_sheets[prefix].to_excel(writer, sheet_name=prefix, index=False)
                wrote_sheet = True
        if not wrote_sheet:
            pd.DataFrame({"Info": ["No data available"]}).to_excel(writer, sheet_name="Default", index=False)

    print(f"✓ Combined file saved at: {output_path}")

if __name__ == "__main__":
    main()