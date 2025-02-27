import csv
import json
from collections import defaultdict

def read_csv_to_dict(file_path):
    with open(file_path, mode='r', newline='') as csv_file:
        return list(csv.DictReader(csv_file))

def replace_rid_with_path(value, base_path):
    # Assuming the base path is where media files are stored
    return f"{base_path}/{value}"

def combine_data_by_slide(animation_data, asset_data, presentation_data, layout_data, base_path):
    combined_data = defaultdict(lambda: {'animations': [], 'assets': [], 'presentation': [], 'layout': []})

    for row in animation_data:
        slide = row['Slide']
        # Filter out unnecessary fields and add to animations
        combined_data[slide]['animations'].append({
            'target_id': row['Target ID'],
            'animation': row['Animation'],
            'property': row['Property'],
            'value': row['Value']
        })

    for row in asset_data:
        slide = row['Slide']
        # Replace 'rid' with folder path for media files
        if row['Type'] == 'Media':
            row['Value'] = replace_rid_with_path(row['Value'], base_path)
        # Filter out unnecessary fields and add to assets
        combined_data[slide]['assets'].append({
            'name': row['Name'],
            'type': row['Type'],
            'value': row['Value']
        })

    for row in presentation_data:
        slide = row['Slide']
        # Filter out unnecessary fields and add to presentation
        combined_data[slide]['presentation'].append({
            'name': row['Name'],
            'type': row['Type'],
            'value': row['Value']
        })

    for row in layout_data:
        slide = row['Slide']
        # Filter out unnecessary fields and add to layout
        combined_data[slide]['layout'].append({
            'name': row['Name'],
            'type': row['Type'],
            'value': row['Value']
        })

    return combined_data

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

# Combine data by slide
combined_data = combine_data_by_slide(animation_data, asset_data, presentation_data, layout_data, base_path)

# Write combined data to JSON
write_json(combined_data, 'updated.json')