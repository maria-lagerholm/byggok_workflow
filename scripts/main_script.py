import os
import sys

def main():
    try:
        # Get the base directory of the current script or executable
        if getattr(sys, 'frozen', False):  # Running as a bundled executable
            base_dir = sys._MEIPASS
        else:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        print(f"Base directory: {base_dir}")  # Debug: Print the base directory

        # Define the path to 'kunder'
        kunder_dir = os.path.join(base_dir, 'kunder')
        print(f"Kunder directory: {kunder_dir}")  # Debug: Print the kunder directory

        # Verify that 'kunder' exists
        if not os.path.exists(kunder_dir):
            print(f"Error: 'kunder' directory not found at {kunder_dir}")
            return  # Exit the function if 'kunder' is not found

        # Add 'scripts' directory to sys.path to import modules
        scripts_dir = os.path.join(base_dir, 'scripts')
        sys.path.insert(0, scripts_dir)

        # Import scripts dynamically
        try:
            import part_1
            import part_2
            import part_3
        except ImportError as e:
            print(f"Error importing scripts: {e}")
            return  # Exit if scripts cannot be imported

        # Run each part in order
        print("Running script 1: part_1.py...")
        part_1.main()

        print("Running script 2: part_2.py...")
        part_2.main()

        print("Running script 3: part_3.py...")
        part_3.main()

        print("All scripts have been executed successfully in the specified order.")
    except Exception as e:
        print(f"An error occurred during workflow execution: {e}")

if __name__ == "__main__":
    main()
