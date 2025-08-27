import os
import pandas as pd

# --- Config ---
# Change this to your target folder
folder_path = r"D:\payMe\kokoAung\projects\pensionProject\data\pensionerTransactionData"

# Output Excel file path (change this to any directory you want)
output_dir = r"D:\payMe\kokoAung\projects\pensionProject\task\fileManagementProject\report"   # Replace with your desired folder
os.makedirs(output_dir, exist_ok=True)  # Create folder if not exists
output_excel = os.path.join(output_dir, "file_list.xlsx")

# --- Collect file paths ---
file_list = []
for root, dirs, files in os.walk(folder_path):
    for file in files:
        file_path = os.path.join(root, file)
        file_list.append({
            "File Name": file,
            "Folder Path": root,
            "Hyperlink": f'=HYPERLINK("{file_path}", "{file}")'
        })

# --- Create DataFrame ---
df = pd.DataFrame(file_list)

# --- Export to Excel ---
with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    df.to_excel(writer, index=False)

print(f"✅ Exported {len(df)} files to {output_excel}")
