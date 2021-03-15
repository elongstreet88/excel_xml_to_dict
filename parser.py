import xml.etree.cElementTree as ET

ns = '{urn:schemas-microsoft-com:office:spreadsheet}'

def excel_xml_to_dict(path = "/master.xml", worksheet_name=None):
    # Parse top level
    
    tree = ET.parse(path)
    root = tree.getroot()

    worksheets = {}
    # Get Worksheets
    for worksheet in root.iter(f"{ns}Worksheet"):
        name = worksheet.attrib[f"{ns}Name"]
        data = get_data_from_worksheet(worksheet)
        worksheets[name] = data

    # Return all if none specified
    if not worksheet_name:
        return worksheets

    # Return specific sheet
    return worksheets[worksheet_name]

def get_data_from_worksheet(worksheet):
     # Get Headers
    headers = []
    for row in worksheet.iter(f"{ns}Row"):
        for cell in row.iter(f"{ns}Cell"):
            value = cell.find(f"{ns}Data")
            headers.append(value.text)
        # Break after first line only
        break

    data = []
    for row in worksheet.iter(f"{ns}Row"):
        entry = {}
        for i, cell in enumerate(row.iter(f"{ns}Cell")):
            value = cell.find(f"{ns}Data")
            entry[headers[i]]=value.text
        data.append(entry)

    # Remove first element from list (the headers)
    data.pop(0)

    return data

print(excel_xml_to_dict(worksheet_name="test"))
