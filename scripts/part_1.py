import pandas as pd
import os
import sys
import shutil

def main():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    file_path = os.path.join(base_dir, 'kunder', 'kundregister.xlsx')
    print(f"Reading Excel file from: {file_path}")

    # Read the Excel file
    df = pd.read_excel(file_path)

    # Ensure 'Saljare' and 'Kopare' columns are present and handle missing values
    df['Saljare'] = pd.to_numeric(df.get('Saljare', 0), errors='coerce').fillna(0).astype(int)
    df['Kopare'] = pd.to_numeric(df.get('Kopare', 0), errors='coerce').fillna(0).astype(int)

    # Group by 'Fastighetsbeteckning' and get the max value for 'Saljare' and 'Kopare'
    df_grouped = df.groupby('Fastighetsbeteckning', as_index=False).agg({'Saljare': 'max', 'Kopare': 'max'})

    mallar_dir = os.path.join(base_dir, 'kunder', 'mallar')

    total_dirs = len(df_grouped)
    new_dirs_created = 0
    total_files_copied = 0

    for i, row in df_grouped.iterrows():
        fastighetsbeteckning = row['Fastighetsbeteckning']
        saljare = row['Saljare']
        kopare = row['Kopare']

        new_directory = fastighetsbeteckning.replace(':', '_').replace(' ', '_')
        new_directory_path = os.path.join(base_dir, 'kunder', new_directory)
        
        if not os.path.exists(new_directory_path):
            os.makedirs(new_directory_path)
            new_dirs_created += 1

        files_copied = 0

        # Determine which files to copy based on 'Saljare' and 'Kopare' values
        files_to_copy = []
        if saljare == 1:
            # Copy files that contain 'saljare_' in their name
            files_to_copy.extend([f for f in os.listdir(mallar_dir) if f.startswith('saljare_')])
        if kopare == 1:
            # Copy files that contain 'kopare_' in their name
            files_to_copy.extend([f for f in os.listdir(mallar_dir) if f.startswith('kopare_')])

        # Remove duplicates from files_to_copy
        files_to_copy = list(set(files_to_copy))

        # Copy the selected files
        for filename in files_to_copy:
            src_file = os.path.join(mallar_dir, filename)
            dst_file = os.path.join(new_directory_path, filename)
            if os.path.isfile(src_file):
                if not os.path.exists(dst_file):
                    shutil.copy2(src_file, dst_file)
                    files_copied += 1

        total_files_copied += files_copied

        if files_copied > 0:
            print(f"Copied {files_copied} files to {new_directory_path}")

        # Print progress every 10% of directories processed
        if (i + 1) % max(1, total_dirs // 10) == 0:
            print(f"Progress: {i + 1}/{total_dirs} directories processed")

    print(f"\nPart 1 completed.")
    print(f"Total directories processed: {total_dirs}")
    print(f"New directories created: {new_dirs_created}")
    print(f"Total files copied: {total_files_copied}")
    print("\nStarting next script (Part 2)...")

if __name__ == "__main__":
    main()
