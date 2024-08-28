import pdfplumber
import re
import openpyxl
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# Path to the JSON key file you just downloaded
json_keyfile_path = "/Users/cynical/Documents/GitHub/ChristiansburgBowling/its-gametime-at-the-superbowl-5f69e9b89124.json"

def extract_handicaps(pdf_path):
    """
    Extracts player names, their averages, and corresponding handicaps from the 'Team Rosters' section of a bowling league PDF.

    Args:
    pdf_path (str): The file path to the PDF document.

    Returns:
    list: A list of tuples where each tuple contains (name, average, handicap).
    """
    player_handicaps = []

    with pdfplumber.open(pdf_path) as pdf:
        # Initialize a counter for the pages
        page_number = 1

        # Loop through each page in the PDF
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.splitlines()

            team_roster_found = False

            for line in lines:
                # Stop processing if we reach the 'Temporary Substitutes' section
                if 'Temporary Substitutes' in line:
                    # Sort the list alphabetically by the first name
                    player_handicaps.sort(key=lambda x: x[0])  # Sort by name
                    return player_handicaps

                # Skip team headers like "11 - WYTHEVILLE MOOSE Lane: 22"
                if '-' in line and 'Lane' in line:
                    continue

                # Check for 'Team Rosters' on the first page or 'of' on subsequent pages
                if 'Team Rosters' in line or (page_number > 1 and 'of' in line):
                    team_roster_found = True

                # Once in the 'Team Rosters' section, look for player names and their HDCP
                if team_roster_found:
                    if 'Name' in line and 'Avg HDCP' in line:
                        # Skip header lines, move to the actual player data
                        continue

                    # Look for lines that have player data
                    if any(char.isdigit() for char in line):  # Detect lines with numeric data
                        # Use regex to extract the name, average, and handicap
                        match = re.match(r"([^\d]+)\s+(\d+|bk\d+)\s+(\d+)", line)
                        if match:
                            name = match.group(1).strip()
                            average = match.group(2).strip()
                            handicap = match.group(3).strip()
                            player_handicaps.append((name, average, handicap))

            # Increment the page number counter after processing each page
            page_number += 1

    # Sort the list alphabetically by the first name
    player_handicaps.sort(key=lambda x: x[0])  # Sort by name
    return player_handicaps

def update_google_sheet(handicaps, sheet_name):
    """
    Updates the player data in the specified Google Sheet.

    Args:
    data (list): A list of tuples containing (name, average, handicap).
    sheet_name (str): The name of the worksheet to update.
    """
    # Define the scope and authenticate with the Google Sheets API
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile_path, scope)
    client = gspread.authorize(creds)
    
    google_sheet_name = "Wednesday Bowling League"

    # Open the Google Sheet by name
    sheet = client.open(google_sheet_name).worksheet(sheet_name)

    # Update the data in the sheet starting at row 2
    for i, (name, average, handicap) in enumerate(handicaps, start=2):
        sheet.update_cell(i, 1, name)
        sheet.update_cell(i, 2, average)
        sheet.update_cell(i, 3, handicap)

    # Get the current date and format it
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Update the "last updated" cell, let's say F4
    sheet.update_acell("F4", f"LAST UPDATED: {current_date} by the Code")

# Paths and worksheet name
pdf_path = "/Users/cynical/Documents/GitHub/ChristiansburgBowling/MenWed.pdf"
excel_path = "/Users/cynical/OneDrive/Mario Kart Wii/Documents/Wednesday Night Sidepots_MacroCopy.xlsx"  # Synced local path
sheet_name = "Men's Handicap Bank | WEDNESDAY"

# Extract the handicaps
handicaps = extract_handicaps(pdf_path)

# Call the function with your data and worksheet name
update_google_sheet(handicaps, sheet_name);

# Optional: Print the results for verification
for entry in handicaps:
    print(entry)
    
print("Reached the end of the script")