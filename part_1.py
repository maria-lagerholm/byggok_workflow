import pandas as pd
import os
import shutil

# Get the current working directory
base_dir = os.getcwd()

# Path to the Excel file
file_path = os.path.join(base_dir, 'kunder', 'kundregister.xlsx')

# Read the Excel file
df = pd.read_excel(file_path)

# Extract the column 'Fastighetsbeteckning' all values without duplicates and save them in a list
fastighetsbeteckningar = df['Fastighetsbeteckning'].drop_duplicates().tolist()

# Ensure that all these newly created directories have copies of all files from kunder/mallar
mallar_dir = os.path.join(base_dir, 'kunder', 'mallar')

def copy_file(src, dst):
    if not os.path.exists(dst):
        shutil.copy2(src, dst)
        print(f"Copied: {dst}")
    else:
        print(f"Skipped existing file: {dst}")

def copy_directory(src, dst):
    if not os.path.exists(dst):
        os.makedirs(dst)
        print(f"Created directory: {dst}")
    
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            copy_directory(s, d)
        else:
            copy_file(s, d)

for fastighetsbeteckning in fastighetsbeteckningar:
    new_directory = fastighetsbeteckning.replace(':', '_').replace(' ', '_')
    new_directory_path = os.path.join(base_dir, 'kunder', new_directory)
    
    if not os.path.exists(new_directory_path):
        print(f"Creating new directory: {new_directory_path}")
        copy_directory(mallar_dir, new_directory_path)
    else:
        print(f"Directory already exists, skipping: {new_directory_path}")

print("Process completed.")