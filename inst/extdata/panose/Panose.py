# Created with ChatGPT 4
# pip install fonttools
from fontTools.ttLib import TTFont
import csv
import os

def get_font_info(font_path):
    try:
        font = TTFont(font_path)
        name = font['name']
        family_name = ""
        subfamily_name = ""
        for record in name.names:
            if record.nameID == 1 and record.platformID == 3 and record.langID == 0x409:  # Family name for Windows, English
                family_name = record.toUnicode()
            elif record.nameID == 2 and record.platformID == 3 and record.langID == 0x409:  # Subfamily name for Windows, English
                subfamily_name = record.toUnicode()
                
        os2 = font['OS/2']
        panose = os2.panose
        panose_bytes = bytes([
            panose.bFamilyType,
            panose.bSerifStyle,
            panose.bWeight,
            panose.bProportion,
            panose.bContrast,
            panose.bStrokeVariation,
            panose.bArmStyle,
            panose.bLetterForm,
            panose.bMidline,
            panose.bXHeight
        ])
        panose_hex_string = panose_bytes.hex().upper()
        
        return family_name, subfamily_name, panose_hex_string
    except Exception as e:
        print(f"Error processing {font_path}: {e}")
        return None, None, None

def extract_fonts_panose(font_directories, output_csv):
    font_data = []  # List to collect font information

    for font_directory in font_directories:
        if os.path.exists(font_directory):
            for font_file in os.listdir(font_directory):
                if font_file.lower().endswith('.ttf'):
                    font_path = os.path.join(font_directory, font_file)
                    family_name, subfamily_name, panose_hex_string = get_font_info(font_path)
                    if family_name and subfamily_name and panose_hex_string:
                        font_data.append((family_name, subfamily_name, panose_hex_string))
    
    # Sort the font data by family name and then subfamily name
    font_data.sort(key=lambda x: (x[0], x[1]))

    # Write the sorted data to the CSV file
    with open(output_csv, 'w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["family", "type", "panose"])  # Updated column names
        
        for entry in font_data:
            writer.writerow(entry)

# Resolving the %AppData% environment variable to include the local Microsoft Windows Fonts directory
appdata_fonts_path = os.path.join(os.environ['APPDATA'], '..', 'Local', 'Microsoft', 'Windows', 'Fonts')

# List of font directories to scan
font_directories = [
    'C:\\Windows\\Fonts',
    appdata_fonts_path,  # Dynamically resolved AppData fonts path
    # Add more directories as needed
]

output_csv = 'panose.csv'

extract_fonts_panose(font_directories, output_csv)
