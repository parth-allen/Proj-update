import os
import json
import xml.etree.ElementTree as ET

# Namespace for parsing XML
NAMESPACE = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}

def parse_xml_to_dict(xml_file):
    """Parses an XML file and converts it into a nested dictionary."""
    if not os.path.exists(xml_file):
        return None

    tree = ET.parse(xml_file)
    root = tree.getroot()

    def parse_element(element):
        parsed = {"tag": element.tag.split("}")[-1], "attributes": element.attrib, "children": []}
        for child in element:
            parsed["children"].append(parse_element(child))
        if element.text and element.text.strip():
            parsed["text"] = element.text.strip()
        return parsed

    return parse_element(root)

def parse_rels(rels_file):
    """Parses a .rels file and extracts linked media, layouts, and masters."""
    relationships = {"media": [], "layout": None, "master": None}

    if os.path.exists(rels_file):
        tree = ET.parse(rels_file)
        root = tree.getroot()

        for rel in root.findall("r:Relationship", NAMESPACE):
            target = rel.attrib.get("Target")
            rel_type = rel.attrib.get("Type")

            if "image" in rel_type or "video" in rel_type or "audio" in rel_type:
                relationships["media"].append(os.path.basename(target))
            elif "slideLayout" in rel_type:
                relationships["layout"] = os.path.basename(target)
            elif "slideMaster" in rel_type:
                relationships["master"] = os.path.basename(target)

    return relationships

def convert_pptx_to_json(extracted_folder, output_json):
    """Converts extracted PPTX files into structured JSON with parsed XML content."""
    
    json_data = {"slides": [], "global_assets": {"themes": [], "layouts": [], "masters": []}}

    slide_folders = [f for f in os.listdir(extracted_folder) if f.startswith("slide_")]

    for slide_folder in slide_folders:
        slide_number = int(slide_folder.replace("slide_", ""))
        slide_path = os.path.join(extracted_folder, slide_folder)

        slide_content_file = None
        rels_file = None
        notes_file = None

        for file in os.listdir(slide_path):
            if file.startswith("slide") and file.endswith(".xml"):
                slide_content_file = file
            elif file.startswith("slide") and file.endswith(".xml.rels"):
                rels_file = file
            elif file.startswith("notesSlide") and file.endswith(".xml"):
                notes_file = file

        slide_content = parse_xml_to_dict(os.path.join(slide_path, slide_content_file)) if slide_content_file else None
        rels_data = parse_rels(os.path.join(slide_path, rels_file)) if rels_file else None
        notes_content = parse_xml_to_dict(os.path.join(slide_path, notes_file)) if notes_file else None

        json_data["slides"].append({
            "slide_number": slide_number,
            "content": slide_content,
            "relationships": rels_data,
            "notes": notes_content
        })

    # Sort slides by slide_number
    json_data["slides"] = sorted(json_data["slides"], key=lambda x: x["slide_number"])

    # Process global assets
    for asset_type in ["theme", "layouts", "masters"]:
        asset_path = os.path.join(extracted_folder, asset_type)
        if os.path.exists(asset_path):
            json_data["global_assets"][asset_type] = os.listdir(asset_path)

    # Write JSON file
    with open(output_json, "w", encoding="utf-8") as json_file:
        json.dump(json_data, json_file, indent=4)

    print(f"✅ JSON file created and sorted: {output_json}")

# Run the script
extracted_folder = "extracted_pptx"
output_json = "pptx_data.json"

convert_pptx_to_json(extracted_folder, output_json)
