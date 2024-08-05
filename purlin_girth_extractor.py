import sys
import pandas as pd
import pyperclip

def extract_purlin_girth_data(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name="Purlin _ Girth")

        # Remove the first two rows and reset index
        df = df.iloc[2:].reset_index(drop=True)

        # Set column names
        df.columns = [
            "CHK", "CHK.1", "Apply Member To", "Material", "Member Type", "Section",
            "Unnamed: 6", "Unnamed: 7", "Unnamed: 8", "Unnamed: 9", "Unnamed: 10",
            "Unnamed: 11", "Unnamed: 12", "Unnamed: 13", "Span", "Unnamed: 15",
            "Unnamed: 16", "Unbraced Length", "Unnamed: 18", "Factor", "Unnamed: 20",
            "Design Load", "Unnamed: 22", "Unnamed: 23", "Unnamed: 24", "Unnamed: 25",
            "Unnamed: 26", "Unnamed: 27", "Unnamed: 28", "Unnamed: 29", "Unnamed: 30",
            "Unnamed: 31", "Unnamed: 32", "Unnamed: 33", "Unnamed: 34", "Unnamed: 35",
            "Defl. Criteria", "Width-Thick Ratio", "Unnamed: 38", "Unnamed: 39",
            "Moment Strength", "Unnamed: 41", "Unnamed: 42", "Unnamed: 43", "Unnamed: 44",
            "Unnamed: 45", "Unnamed: 46", "Shear Strength", "Unnamed: 48", "Unnamed: 49",
            "Unnamed: 50", "Unnamed: 51", "Unnamed: 52", "Deflection", "Unnamed: 54"
        ]

        # Extract and convert Deflection and Ratio
        df["Deflection"] = df["Deflection"].astype(str).str.extract(r"(\d+\.\d+)").astype(float)
        df["Ratio"] = df["Unnamed: 54"].astype(str).str.extract(r"(\d+\.\d+)").astype(float)

        # Calculate Span/300
        df["Calculated Span/300"] = (df["Span"] * 1000 / 300).round(2)

        # Select required columns and remove rows with NaN values
        output_df = df[["Deflection", "Calculated Span/300", "Ratio"]].dropna()

        # Convert to list and then to string
        result = output_df.to_string(index=False, header=False)

        # Copy to clipboard
        pyperclip.copy(result)

        print("Data successfully extracted and copied to clipboard.")
        print("Extracted data:")
        print(result)

    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python program_name.py <excel_file_path>")
    else:
        file_path = sys.argv[1]
        extract_purlin_girth_data(file_path)