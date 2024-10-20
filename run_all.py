import os
import subprocess
import sys

# Get the directory of the current script (which should be the workflow directory)
if getattr(sys, 'frozen', False):  # Running as a PyInstaller bundle
    workflow_dir = sys._MEIPASS  # This points to the temporary directory
else:
    workflow_dir = os.path.dirname(os.path.abspath(__file__))

print(f"Workflow directory: {workflow_dir}")  # Debug: Print the workflow directory

# Define the path to kunder
kunder = os.path.join(workflow_dir, 'kunder')
print(f"Main directory: {kunder}")  # Debug: Print the main directory

# Verify that kunder exists
if not os.path.exists(kunder):
    print(f"Error: kunder not found at {kunder}")
    sys.exit(1)

# Define the paths to the scripts in the exact order they must be executed
scripts = [
    os.path.join(workflow_dir, 'part_1.py'),
    os.path.join(workflow_dir, 'part_2.py'),
    os.path.join(workflow_dir, 'part_3.py')
]

# Function to run a script
def run_script(script_path, order):
    try:
        print(f"Running script {order}: {os.path.basename(script_path)}...")
        print(f"Script path: {script_path}")  # Debug: Print the script path
        if not os.path.isfile(script_path):
            print(f"Script not found: {script_path}")
            sys.exit(1)  # Exit if script is not found
        # Pass kunder as an environment variable to the scripts
        env = os.environ.copy()
        env['KUNDER'] = kunder
        subprocess.run([sys.executable, script_path], check=True, env=env)
        print(f"Completed script {order}: {os.path.basename(script_path)}.\n")
    except subprocess.CalledProcessError as e:
        print(f"Error running script {order}: {os.path.basename(script_path)}: {e}")
        sys.exit(1)  # Exit if script execution fails

# Main execution
if __name__ == "__main__":
    print("Starting the workflow...\n")
    print("Contents of workflow directory:", os.listdir(workflow_dir))  # Debug: Print workflow directory contents
    print("Contents of main directory:", os.listdir(kunder))  # Debug: Print main directory contents
    
    for index, script in enumerate(scripts, start=1):
        run_script(script, index)
    
    print("All scripts have been executed successfully in the specified order.")
