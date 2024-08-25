import pdfplumber
import re
import openpyxl

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

def update_excel(data, excel_path, sheet_name):
    """
    Updates the player data in the specified Excel sheet.

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

    # Save the workbook
    workbook.save(excel_path)

# Paths and worksheet name
pdf_path = "/Users/cynical/Documents/GitHub/ChristiansburgBowling/4950308212024F202401STANDG00.pdf"
excel_path = "/Users/cynical/Documents/GitHub/ChristiansburgBowling/Wednesday Night Sidepots - MASTER COPY DO NOT EDIT.xlsx"
sheet_name = "Handicap Bank | WEDNESDAY"

# Extract the handicaps
handicaps = extract_handicaps(pdf_path)

# Update the Excel sheet
update_excel(handicaps, excel_path, sheet_name)

# Optional: Print the results for verification
for entry in handicaps:
    print(entry)