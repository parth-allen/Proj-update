import csv
import json
from collections import defaultdict

def read_csv_to_dict(file_path):
    with open(file_path, mode='r', newline='') as csv_file:
        return list(csv.DictReader(csv_file))

def replace_rid_with_path(value, base_path):
    # Assuming the base path is where media files are stored
    return f"{base_path}/{value}"

def organize_data_by_slide(animation_data, asset_data, presentation_data, layout_data, base_path):
    slides_data = defaultdict(lambda: {
        'animations': [],
        'assets': [],
        'presentation_details': [],
        'layout_details': []
    })

    for row in animation_data:
        slide = row['Slide']
        # Create an animation dictionary with descriptive fields
        animation_info = {
            'target_id': row['Target ID'],
            'animation_type': row['Animation'],
            'property_name': row['Property'],
            'property_value': row['Value']
        }
        slides_data[slide]['animations'].append(animation_info)

    for row in asset_data:
        slide = row['Slide']
        # Replace 'rid' with folder path for media files
        if row['Type'] == 'Media':
            row['Value'] = replace_rid_with_path(row['Value'], base_path)
        # Create an asset dictionary with descriptive fields
        asset_info = {
            'asset_name': row['Name'],
            'asset_type': row['Type'],
            'asset_value': row['Value'],
            'related_media': []  # Initialize related media list for hierarchical structure
        }
        slides_data[slide]['assets'].append(asset_info)

    for row in presentation_data:
        slide = row['Slide']
        # Create a presentation dictionary with descriptive fields
        presentation_info = {
            'element_name': row['Name'],
            'element_type': row['Type'],
            'element_value': row['Value']
        }
        slides_data[slide]['presentation_details'].append(presentation_info)

    for row in layout_data:
        slide = row['Slide']
        # Create a layout dictionary with descriptive fields
        layout_info = {
            'layout_name': row['Name'],
            'layout_type': row['Type'],
            'layout_value': row['Value']
        }
        slides_data[slide]['layout_details'].append(layout_info)

    # Example of creating a parent-child relationship
    for slide, data in slides_data.items():
        for asset in data['assets']:
            if asset['asset_type'] == 'Image':
                # Example: Add a video as a child to an image
                for child_asset in data['assets']:
                    if child_asset['asset_type'] == 'Video':
                        asset['related_media'].append(child_asset)

    return slides_data

def write_json(data, output_file):
    with open(output_file, 'w') as json_file:
        json.dump(data, json_file, indent=4)

# Read data from CSV files
animation_data = read_csv_to_dict('animation.csv')
asset_data = read_csv_to_dict('asset.csv')
presentation_data = read_csv_to_dict('presentation.csv')
layout_data = read_csv_to_dict('layout.csv')

# Define the base path for media files
base_path = 'media_files'

# Organize data by slide
slides_data = organize_data_by_slide(animation_data, asset_data, presentation_data, layout_data, base_path)

# Write organized data to JSON
write_json(slides_data, 'descriptive_slides.json')