import csv
import json
from collections import defaultdict

def build_tree(rows):
    # Create a dictionary of nodes where each node will include a "children" array
    nodes = {}
    for row in rows:
        asset_id = row.get("Asset", "").strip()
        node = {key: value for key, value in row.items() if key not in ["Pptx", "Slide"]}  # Remove "Pptx" and "Slide" keys from children
        node["children"] = []
        nodes[asset_id] = node
    
    # Build the hierarchy by connecting child nodes to their parent node
    tree = {}
    for asset_id, node in nodes.items():
        parent = node.get("Parent", "").strip()
        if parent.lower() == "root" or parent == "":
            # If node has "Root" (or empty) as parent, add it as a top-level node
            tree[asset_id] = node
        else:
            # Otherwise, nest it under the parent's "children" if the parent exists
            parent_node = nodes.get(parent)
            if parent_node:
                parent_node["children"].append(node)
            else:
                # If parent not found, insert it at the top level
                tree[asset_id] = node
    return tree

def main():
    input_csv = "asset.csv"
    output_json = "asset.json"
    result = {}
    # Group rows by Slide for separate trees per slide.
    slides = defaultdict(list)

    with open(input_csv, newline='', encoding='utf-8') as fp:
        reader = csv.DictReader(fp)
        for row in reader:
            slide = row.get("Slide", "").strip()
            slides[slide].append({key: value for key, value in row.items() if key != "Pptx"})  # Remove "Pptx" key
    
    # Build a nested structure: Slide -> tree of assets
    for slide, rows in slides.items():
        tree = build_tree(rows)
        for node in tree.values():
            node.pop("Slide", None)  # Remove "Slide" key from top-level nodes
        result[slide] = tree

    # Write the JSON result to the output file with indentation for readability.
    with open(output_json, "w", encoding="utf-8") as fp:
        json.dump(result, fp, indent=4)

    print(f"JSON file created: {output_json}")

if __name__ == "__main__":
    main()
