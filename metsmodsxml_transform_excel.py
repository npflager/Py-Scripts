# MODS/METS XML to XLSX Crosswalk
# Project ECHO migration from ResCarta to CONTENTdm
# Nate Pflager, ILS Administrator, Winding Rivers Library System

import xml.etree.ElementTree as ET
import pandas as pd

def parse_xml(xml_file):
    # Create an ElementTree object and parse the XML file
    tree = ET.parse(xml_file)
    root = tree.getroot()

    # Define namespaces
    namespaces = {
        "mods": "http://www.loc.gov/mods/v3",
        "mets": "http://www.loc.gov/METS/"
    }

    # Extract the necessary data from the XML structure
    data = []
    for dmdSec in root.iterfind(".//mets:dmdSec", namespaces):
        record_data = {}
        mods = dmdSec.find(".//mods:mods", namespaces)
        
        title = mods.find(".//mods:title", namespaces)
        if title is not None:
            record_data["Title"] = title.text

        
        for child in mods:
            if child.tag.endswith("title"):
                record_data["Title"] = child.text
            elif child.tag.endswith("abstract"):
                record_data["Abstract"] = child.text
            elif child.tag.endswith("identifier"):
                record_data["Identifier"] = child.text
            elif child.tag.endswith("typeOfResource"):
                record_data["Type of Resource"] = child.text
            elif child.tag.endswith("genre"):
                record_data["Genre"] = child.text
            elif child.tag.endswith("subject"):
                if child.find("mods:hierarchicalGeographic", namespaces) is not None:
                    hier_geo = child.find("mods:hierarchicalGeographic", namespaces)
                    country = hier_geo.find(".//mods:country", namespaces)
                    state = hier_geo.find(".//mods:state", namespaces)
                    county = hier_geo.find(".//mods:county", namespaces)
                    city = hier_geo.find(".//mods:city", namespaces)
                    if country is not None:
                        record_data["Country"] = country.text
                    if state is not None:
                        record_data["State"] = state.text
                    if county is not None:
                        record_data["County"] = county.text
                    if city is not None:
                        record_data["City"] = city.text
                if child.find("mods:geographic", namespaces) is not None:
                    geographic = child.find("mods:geographic", namespaces).text
                    record_data.setdefault("Geographic", []).append(geographic)
                elif child.find("mods:temporal", namespaces) is not None:
                    temporal = child.find("mods:temporal", namespaces).text
                    record_data.setdefault("Temporal", []).append(temporal)
                elif child.find("mods:topic", namespaces) is not None:
                    topic = child.find("mods:topic", namespaces).text
                    record_data.setdefault("Subject", []).append(topic)
            elif child.tag.endswith("originInfo"):
                for info in child:
                    if info.tag.endswith("place"):
                        record_data["Place"] = info.find("mods:placeTerm", namespaces).text
                    elif info.tag.endswith("publisher"):
                        record_data["Publisher"] = info.text
                    elif info.tag.endswith("dateIssued"):
                        record_data["Date Issued"] = info.text
                    elif info.tag.endswith("dateCaptured"):
                        record_data["Date Captured"] = info.text
            elif child.tag.endswith("part"):
                part_number = child.find(".//mods:number", namespaces).text
                record_data["Date Published"] = part_number
            elif child.tag.endswith("titleInfo"):
                if child.attrib.get("type") == "alternative":
                    alternative_title = child.find("mods:title", namespaces).text
                    record_data["Alternative Title"] = alternative_title
            elif child.tag.endswith("name"):
                if child.attrib.get("type") == "corporate":
                    corporate_name = child.find("mods:namePart", namespaces).text
                    record_data["Corporate Name"] = corporate_name

        # Additional fields from the provided XML structure
            elif child.tag.endswith("language"):
                language_term = child.find("mods:languageTerm", namespaces).text
                record_data["Language"] = language_term
            elif child.tag.endswith("physicalDescription"):
                form = child.find(".//mods:form", namespaces).text
                extent = child.find(".//mods:extent", namespaces).text
                record_data["Form"] = form
                record_data["Dimensions"] = extent

        
        access_condition = mods.find(".//mods:accessCondition", namespaces)
        if access_condition is not None:
            record_data["Access Condition"] = access_condition.text

        data.append(record_data)

    return data

def xml_to_excel(xml_file, excel_file):
    # Parse XML file
    data = parse_xml(xml_file)

    # Convert data to a DataFrame
    df = pd.DataFrame(data)

    # Reorder columns to have "Title" as the first column
    if "Title" in df.columns:
        columns = df.columns.tolist()
        columns.remove("Title")
        columns = ["Title"] + columns
        df = df[columns]

    # Export to Excel
    df.to_excel(excel_file, index=False)

    print(f"Data exported to {excel_file} successfully!")

#File Management
xml_file = "wrls_ECHO_metadata.xml"  # Replace with your XML file path
excel_file = "wrls_ECHO_output.xlsx"  # Replace with desired output Excel file path

xml_to_excel(xml_file, excel_file)
