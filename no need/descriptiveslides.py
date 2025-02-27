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
        'assets': []
    })

    seen_assets = set()  # Track seen assets to avoid duplicates

    for row in asset_data:
        slide = row['Slide']
        # Replace 'rid' with folder path for media files
        if row['Type'] == 'Media':
            row['Value'] = replace_rid_with_path(row['Value'], base_path)
        
        asset_key = (row['Asset'], row['Value'])  # Unique key for each asset
        if asset_key not in seen_assets:
            seen_assets.add(asset_key)
            asset_info = {
                'asset_id': row['Asset'],
                'parent_id': row['Parent'],
                'asset_name': row['Name'],
                'asset_type': row['Type'],
                'asset_value': row['Value'],
                'related_media': []  # Initialize related media list for hierarchical structure
            }
            slides_data[slide]['assets'].append(asset_info)

    # Example of creating a parent-child relationship
    for slide, data in slides_data.items():
        for asset in data['assets']:
            if asset['asset_type'] == 'Image':
                # Example: Add a video as a child to an image
                for child_asset in data['assets']:
                    if child_asset['asset_type'] == 'Video' and child_asset['parent_id'] == asset['asset_id']:
                        asset['related_media'].append(child_asset)

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
write_json(slides_data, 'descriptiveslides.json')