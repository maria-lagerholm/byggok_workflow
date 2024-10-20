import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import os
import sys
import threading

def main():
    # Create the main window
    root = tk.Tk()
    root.title("Workflow Execution")

    # Create a ScrolledText widget
    text_area = ScrolledText(root, wrap='word', width=80, height=20)
    text_area.pack(expand=True, fill='both')

    # Redirect stdout and stderr to the text area
    class TextRedirector(object):
        def __init__(self, widget):
            self.widget = widget

        def write(self, s):
            self.widget.insert('end', s)
            self.widget.see('end')

        def flush(self):
            pass  # For compatibility with Python's flush behavior

    sys.stdout = TextRedirector(text_area)
    sys.stderr = TextRedirector(text_area)

    # Function to run the workflow
    def run_workflow():
        try:
            # Get the directory of the current script
            if getattr(sys, 'frozen', False):  # Running as a PyInstaller bundle
                base_dir = sys._MEIPASS
            else:
                # Go up one level from 'scripts' to reach 'workflow'
                base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

            print(f"Base directory: {base_dir}")  # Debug: Print the base directory

            # Define the path to kunder
            kunder_dir = os.path.join(base_dir, 'kunder')
            print(f"Kunder directory: {kunder_dir}")  # Debug: Print the kunder directory

            # Verify that kunder exists
            if not os.path.exists(kunder_dir):
                print(f"Error: kunder not found at {kunder_dir}")
                return  # Exit the function if kunder is not found

            # Add scripts directory to sys.path to import modules
            scripts_dir = os.path.join(base_dir, 'scripts')
            sys.path.insert(0, scripts_dir)

            try:
                import part_1
                import part_2
                import part_3
            except ImportError as e:
                print(f"Error importing scripts: {e}")
                return  # Exit the function if scripts cannot be imported

            print("Starting the workflow...\n")
            print("Contents of base directory:", os.listdir(base_dir))  # Debug: Print base directory contents
            print("Contents of kunder directory:", os.listdir(kunder_dir))  # Debug: Print kunder directory contents

            # Run each part
            print("Running script 1: part_1.py...")
            part_1.main()

            print("Running script 2: part_2.py...")
            part_2.main()

            print("Running script 3: part_3.py...")
            part_3.main()

            print("All scripts have been executed successfully in the specified order.")
        except Exception as e:
            print(f"An error occurred during workflow execution: {e}")

    # Run the workflow in a separate thread
    threading.Thread(target=run_workflow).start()

    # Start the Tkinter event loop
    root.mainloop()

if __name__ == "__main__":
    main()
