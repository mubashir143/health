import pandas as pd
import difflib
import sys

def standardize_job_titles(file_path, column_name='job_title', standard_title='Community Health Inspector', threshold=80):
    """
    Standardizes job titles in an Excel file by replacing variations of a standard title
    with the correct spelling using fuzzy string matching.

    Parameters:
    - file_path: Path to the Excel file
    - column_name: Name of the column containing job titles (default: 'job_title')
    - standard_title: The correct job title to standardize to (default: 'Community Health Inspector')
    - threshold: Similarity threshold (0-100) for fuzzy matching (default: 80)

    Returns:
    - Saves the updated Excel file with standardized titles
    - Prints the count of standardized titles
    """

    try:
        # Read the Excel file
        df = pd.read_excel(file_path)

        if column_name not in df.columns:
            print(f"Error: Column '{column_name}' not found in the Excel file.")
            return

        # Function to check similarity and standardize
        def standardize_title(title):
            if isinstance(title, str):
                # Calculate similarity ratio using difflib
                similarity = difflib.SequenceMatcher(None, title.lower(), standard_title.lower()).ratio() * 100
                if similarity >= threshold:
                    return standard_title
            return title

        # Apply standardization
        original_count = df[column_name].value_counts().get(standard_title, 0)
        df[column_name] = df[column_name].apply(standardize_title)
        new_count = df[column_name].value_counts().get(standard_title, 0)

        # Save the updated file
        output_path = file_path.replace('.xlsx', '_standardized.xlsx')
        df.to_excel(output_path, index=False)

        print(f"Standardization complete!")
        print(f"Original count of '{standard_title}': {original_count}")
        print(f"New count of '{standard_title}': {new_count}")
        print(f"Updated file saved as: {output_path}")

        # Show filtering example
        filtered_df = df[df[column_name] == standard_title]
        print(f"\nFiltered records with '{standard_title}': {len(filtered_df)} rows")

    except Exception as e:
        print(f"Error processing file: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python standardize_job_titles.py <excel_file_path> [column_name] [standard_title] [threshold]")
        sys.exit(1)

    file_path = sys.argv[1]
    column_name = sys.argv[2] if len(sys.argv) > 2 else 'job_title'
    standard_title = sys.argv[3] if len(sys.argv) > 3 else 'Community Health Inspector'
    threshold = int(sys.argv[4]) if len(sys.argv) > 4 else 80

    standardize_job_titles(file_path, column_name, standard_title, threshold)