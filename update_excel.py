import json
from openpyxl import load_workbook

def update_excel_from_json(json_file, excel_file):
    with open(json_file, 'r') as f:
        scraped_data = json.load(f)

    workbook = load_workbook(excel_file)
    sheet = workbook["Handicap Sidepot"]

    for player in scraped_data:
        name = player['name']
        scores = player['scoresArray']

        for row in range(2, 52):  # Rows 2 to 51
            if sheet.cell(row=row, column=2).value == name:
                for i, score in enumerate(scores):
                    cell = sheet.cell(row=row, column=4 + i)  # Columns D, E, F
                    if cell.value is None:  # Only write if the cell is empty
                        cell.value = score
                        break
                break

    workbook.save(excel_file)
    print(f"Excel file {excel_file} updated successfully.")

# Usage
update_excel_from_json('scraped_data.json', '/Users/cynical/OneDrive/Mario Kart Wii/Documents/Wednesday Night Sidepots_MacroCopy.xlsx')