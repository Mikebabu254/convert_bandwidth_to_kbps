import pandas as pd

# Load the Excel file
file_path = 'Net Monitoring.xlsx'  # Replace with your actual file name
sheet_name = 'Sheet1'                # Replace with the sheet you want to process

# Read the sheet
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Function to normalize bandwidth to Kbps (numbers only)
def normalize_bandwidth(cell):
    if isinstance(cell, str):
        cell_lower = cell.lower().strip()

        try:
            # Convert Mbps to Kbps
            if 'mbps' in cell_lower:
                number = float(cell_lower.replace('mbps', '').strip())
                return round(number * 1024, 2)

            # Clean up existing Kbps
            elif 'kbps' in cell_lower:
                number = float(cell_lower.replace('kbps', '').strip())
                return round(number, 2)

            # Convert bps to Kbps
            elif 'bps' in cell_lower:
                number = float(cell_lower.replace('bps', '').strip())
                return round(number / 1024, 2)

        except ValueError:
            return cell  # Return original cell if conversion fails

    return cell  # Return original if not a string

# Apply the function to all cells
df_converted = df.applymap(normalize_bandwidth)

# Save to a new sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_converted.to_excel(writer, sheet_name='CleanedSheet-Kbps', index=False)

print("Converted sheet saved successfully (numeric Kbps only).")
