import csv
import json
import os
from collections import defaultdict

def read_csv_to_dict(file_path):
    with open(file_path, mode='r', newline='') as csv_file:
        return list(csv.DictReader(csv_file))

def replace_rid_with_path(value, base_path):
    # Normalize the path to avoid duplicates
    return os.path.normpath(f"{base_path}/{value}")

def organize_data_by_slide(asset_data, base_path):
    slides_data = defaultdict(lambda: {
        'assets': {}
    })

    for row in asset_data:
        slide = row['Slide']
        # Replace 'rid' with folder path for media files
        if row['Type'] == 'Media':
            row['Value'] = replace_rid_with_path(row['Value'], base_path)
        
        asset_info = {
            'asset_id': row['Asset'],
            'asset_name': row['Name'],
            'asset_type': row['Type'],
            'asset_value': row['Value'],
            'children': {}  # Initialize children dictionary for hierarchical structure
        }

        # If the asset has a parent, nest it under the parent
        parent_id = row['Parent']
        if parent_id and parent_id in slides_data[slide]['assets']:
            slides_data[slide]['assets'][parent_id]['children'][row['Asset']] = asset_info
        else:
            # Otherwise, add it as a top-level asset
            slides_data[slide]['assets'][row['Asset']] = asset_info

    return slides_data

def write_json(data, output_file):
    with open(output_file, 'w') as json_file:
        json.dump(data, json_file, indent=4)

# Read data from CSV files
asset_data = read_csv_to_dict('asset.csv')

# Define the base path for media files
base_path = 'media_files'

# Organize data by slide
slides_data = organize_data_by_slide(asset_data, base_path)

# Write organized data to JSON
write_json(slides_data, 'parent_child_format.json')