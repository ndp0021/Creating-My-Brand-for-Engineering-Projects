import tkinter as tk
import random
from itertools import cycle
from time import sleep
import pandas as pd

# ================================================================
# Star Name Picker Wheel
# ================================================================
# Project by: Neer Patel
# Organization: Fusion Space
# Date Created: May 3, 2023
# Description:
# This application creates a graphical user interface (GUI) using tkinter
# to randomly select star names from an Excel file via a simulated spinning wheel.
# ================================================================

# Load star names from an Excel file
def load_star_names():
    """Loads star names from a specified Excel sheet into a list, skipping the header."""
    try:
        # Reads star names from the first column of the 'Reference' sheet
        df = pd.read_excel("IAU-Catalog of Star Names (2024).xlsx", sheet_name="Reference", usecols="A", header=0)
        return df.iloc[:,0].tolist()  # Converts the column to a list
    except Exception as e:
        print(f"Failed to read the Excel file: {e}")
        return []

# List of star names from Excel file
star_names = load_star_names()

# Function to simulate the spinning of the wheel
def spin():
    """Simulates spinning the wheel by cycling through star names and slowing down to stop at a random name."""
    # Create a cyclic iterator over the star names
    cyclic_names = cycle(star_names)
    # Determine a random stop point, allowing for potential full list cycles
    stop_after = random.randint(0, len(star_names) + 100)
    count = 0
    # Continuously update the label with star names to simulate spinning
    while count < stop_after:
        root.update()  # Necessary to keep the GUI responsive
        name = next(cyclic_names)  # Get the next star name
        name_label.config(text=name)  # Display the star name
        sleep(max(0.01, 0.05 - (count / 1000)))  # Decrease sleep time to simulate slowing down
        count += 1

# Set up the main window for the GUI
root = tk.Tk()
root.title("Star Name Spinner - Fusion Space")

# Label widget to display the star name
name_label = tk.Label(root, text="", font=('Helvetica', 24), width=20)
name_label.pack(pady=20)

# Button widget to start the spinning
spin_button = tk.Button(root, text="Spin the Wheel", command=spin, font=('Helvetica', 16))
spin_button.pack(pady=20)

# Start the GUI event loop
root.mainloop()