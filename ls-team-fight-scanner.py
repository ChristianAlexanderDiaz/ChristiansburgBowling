import pdfplumber
import re

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

# Path to the PDF file
pdf_path = "/Users/cynical/Documents/GitHub/BowlingHDCPScanner/4950308212024F202401STANDG00.pdf"

# Extract the text starting from "Team Rosters"
handicaps = extract_text_from_team_rosters(pdf_path)

# Print the extracted player names, averages, and handicaps
for entry in handicaps:
    print(entry)