import pandas as pd
import os
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, filename='anova_significance.log', filemode='w', format='%(levelname)s:%(message)s')

# Function to extract significant posthoc connections (p-unc < 0.001) and save them in a file
def extract_significant_uncorrected_posthoc(sheets):
    # Create output directory for results
    output_dir = 'anova_results'
    os.makedirs(output_dir, exist_ok=True)

    try:
        for sheet_name, df in sheets.items():
            if '_posthoc' in sheet_name:
                posthoc_results = []

                for _, row in df.iterrows():
                    try:
                        # Extract p-unc for uncorrected p-values
                        p_value = pd.to_numeric(row['p-unc'], errors='coerce')
                        hedges_value = pd.to_numeric(row['hedges'], errors='coerce')

                        # If uncorrected p-value is significant (p < 0.001)
                        if pd.notna(p_value) and p_value < 0.001:
                            higher = row['A'] if hedges_value > 0 else row['B']
                            contrast_type = row['Contrast']
                            
                            # Handle interaction effects to include both group and condition
                            if 'Condition * Group' in contrast_type:
                                condition = row['Condition']
                                posthoc_results.append(f"{contrast_type}: Higher value for {higher} in condition {condition}")
                            # Handle regular group or condition effects
                            else:
                                posthoc_results.append(f"{contrast_type}: Higher value between {row['A']} and {row['B']} - Higher: {higher}")
                    except ValueError:
                        continue

                # Save uncorrected significant results to file
                if posthoc_results:
                    with open(os.path.join(output_dir, f'{sheet_name}_uncorrected.txt'), 'w') as f:
                        for result in posthoc_results:
                            f.write(result + '\n')

    except Exception as e:
        logging.error(f"Error in extracting uncorrected posthoc connections: {e}")
        raise e

# Load the Excel file
file_path = 'anova.xlsx'  # Adjust the file path to your needs
xls = pd.ExcelFile(file_path)

# Load all sheets into a dictionary
sheets = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}

# Run the extraction for uncorrected p-value significant results
extract_significant_uncorrected_posthoc(sheets)

# Check the output files created in the directory
output_files = os.listdir('anova_results')
print(f"Significant results saved in the following files: {output_files}")
