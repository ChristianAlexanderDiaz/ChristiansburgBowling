import pdfplumber
import re
import openpyxl
from datetime import datetime

def process_line_for_players(line):
    """
    Processes a single line to extract one or more players' data.
    Args:
    line (str): The text line containing player data.
    Returns:
    list: A list of tuples containing (name, average, handicap).
    """
    # Use regex to match valid bowling data (name, average, handicap)
    matches = re.findall(r"([^\d]+)\s+\d+\s+\d+\s+(bk\d+|\d+)\s+(\d+)", line)
    player_handicaps = []
    
    for match in matches:
        name = match[0].strip()
        average = match[1].strip()
        handicap = match[2].strip()
        player_handicaps.append((name, average, handicap))
    
    return player_handicaps

def extract_text_from_team_rosters(pdf_path):
    """
    Extracts text from the PDF starting from the 'Team Rosters' section, handling lines that may contain multiple players.
    Args:
    pdf_path (str): The file path to the PDF document.
    Returns:
    list: A list of tuples where each tuple contains (name, average, handicap).
    """
    player_handicaps = []
    start_extracting = False
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = text.splitlines()

            for line in lines:
                if 'Team Rosters' in line:
                    start_extracting = True
                    continue
                
                if start_extracting:
                    # Process the line to extract player data
                    players_data = process_line_for_players(line)
                    if players_data:
                        player_handicaps.extend(players_data)
    
    # Sort the list alphabetically by the player's name
    player_handicaps.sort(key=lambda x: x[0])
    
    return player_handicaps

def update_excel(data, excel_path, sheet_name):
    """
    Updates the player data in the specified Excel sheet and updates F4 with the last update date.

    Args:
    data (list): A list of tuples containing (name, average, handicap).
    excel_path (str): The file path to the Excel document.
    sheet_name (str): The name of the worksheet to update.
    """
    # Load the workbook and select the specified worksheet
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook[sheet_name]

    # Write the data to the sheet starting at A2
    for i, (name, average, handicap) in enumerate(data, start=2):
        sheet[f"A{i}"] = name
        sheet[f"B{i}"] = average
        sheet[f"C{i}"] = handicap

    # Get the current date and format it
    current_date = datetime.now().strftime("%m/%d %I:%M %p")

    # Update cell F4 with the last updated information
    sheet["F4"] = f"LAST UPDATED: {current_date}"

    # Save the workbook
    workbook.save(excel_path)


# Path to the PDF file
pdf_path = "/Users/cynical/Documents/GitHub/ChristiansburgBowling/OpenWed.pdf"
excel_path = "/Users/cynical/OneDrive/Mario Kart Wii/Documents/Wednesday Night Sidepots_MacroCopy.xlsx"  # Synced local path
sheet_name = "Open Handicap Bank | WEDNESDAY"


# Extract the text starting from "Team Rosters"
handicaps = extract_text_from_team_rosters(pdf_path)

# Update the Excel sheet
update_excel(handicaps, excel_path, sheet_name)

# Print the extracted player names, averages, and handicaps
for entry in handicaps:
    print(entry)
        
print("Reached the end of the script")