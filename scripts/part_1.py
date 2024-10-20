import pandas as pd
import os
import sys
import shutil

def copy_file(src, dst):
    if not os.path.exists(dst):
        shutil.copy2(src, dst)
        return True
    return False

def copy_directory(src, dst):
    if not os.path.exists(dst):
        os.makedirs(dst)
    
    files_copied = 0
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            files_copied += copy_directory(s, d)
        else:
            if copy_file(s, d):
                files_copied += 1
    return files_copied

def main():
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
    else:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    file_path = os.path.join(base_dir, 'kunder', 'kundregister.xlsx')
    print(f"Reading Excel file from: {file_path}")

    df = pd.read_excel(file_path)
    fastighetsbeteckningar = df['Fastighetsbeteckning'].drop_duplicates().tolist()
    mallar_dir = os.path.join(base_dir, 'kunder', 'mallar')

    total_dirs = len(fastighetsbeteckningar)
    new_dirs_created = 0
    total_files_copied = 0

    for i, fastighetsbeteckning in enumerate(fastighetsbeteckningar, 1):
        new_directory = fastighetsbeteckning.replace(':', '_').replace(' ', '_')
        new_directory_path = os.path.join(base_dir, 'kunder', new_directory)
        
        if not os.path.exists(new_directory_path):
            files_copied = copy_directory(mallar_dir, new_directory_path)
            new_dirs_created += 1
            total_files_copied += files_copied
            print(f"Created new directory: {new_directory_path} (Copied {files_copied} files)")

        # Print progress every 10% of directories processed
        if i % max(1, total_dirs // 10) == 0:
            print(f"Progress: {i}/{total_dirs} directories processed")

    print(f"\nPart 1 completed.")
    print(f"Total directories processed: {total_dirs}")
    print(f"New directories created: {new_dirs_created}")
    print(f"Total files copied: {total_files_copied}")
    print("\nStarting next script (Part 2)...")

if __name__ == "__main__":
    main()