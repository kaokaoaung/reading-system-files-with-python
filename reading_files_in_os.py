import os
import pandas as pd

# --- Config ---
# Change this to your target folder
folder_path = r"your folder path"

# Output Excel file path (change this to any directory you want)
output_dir = r"your folder path"   # Replace with your desired folder
os.makedirs(output_dir, exist_ok=True)  # Create folder if not exists
output_excel = os.path.join(output_dir, "file_list.xlsx")

# --- Collect file paths ---
file_list = []
for root, dirs, files in os.walk(folder_path):
    for file in files:
        file_list.append({
            "File Name": file,
            "Folder Path": root
        })

# --- Create DataFrame ---
df = pd.DataFrame(file_list)

# --- Export to Excel ---
df.to_excel(output_excel, index=False)

print(f"✅ Exported {len(df)} files to {output_excel}")
