import time
import os

# Define the path to your Excel file
excel_path = "/Users/cynical/OneDrive/Mario Kart Wii/Documents/Wednesday Night Sidepots_MacroCopy.xlsx"  # Synced local path

# Close the Excel application on macOS
os.system("pkill -x 'Microsoft Excel'")  # Terminates the Excel process

# Simulate a brief delay before reopening
time.sleep(2)

# Reopen the Excel file
os.system(f"open '{excel_path}'")  # Opens the Excel file with the default application