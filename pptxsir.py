import os
import sys
import zipfile
import xml.etree.ElementTree as ET
import csv
import traceback
import re
import random

# Define namespaces
namespaces = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p14': "http://schemas.microsoft.com/office/powerpoint/2010/main",
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
}

def get_auto_id():
    return "Auto_"+str(random.randint(1001, 4999))

def get_asset_tag(asset):
    return re.sub(r'\{[^}]*\}', '', asset.tag),

def get_asset(element, assets, parentId, relation):
    for txBody in element.findall('./p:txBody', namespaces):
        texts = txBody.findall('./*/a:r/a:t', namespaces)
        if texts is not None and len(texts) > 0:
            asset_info = {
                'id': get_auto_id(),
                'parentId': parentId,
                'name': "Text",
                'type': "Text",
                'value': texts[0].text if texts is not None and len(texts) > 0 else "None"
            }
            assets.append(asset_info)
    
    for spPr in element.findall('./p:spPr', namespaces):
        geom = spPr.find('./a:prstGeom', namespaces)
        if geom is None:
            geom = spPr.find('./a:custGeom', namespaces)
            if geom is None:
                asset_info = {
                    'id': get_auto_id(),
                    'parentId': parentId,
                    'name': "Custom Geometry",
                    'type': "Shape",
                    'value': "None"
                }
                assets.append(asset_info)
            else:
                asset_info = {
                    'id': get_auto_id(),
                    'parentId': parentId,
                    'name': "Unknown Geometry",
                    'type': "Shape",
                    'value': "None"
                }
                assets.append(asset_info)
        else:
            asset_info = {
                'id': get_auto_id(),
                'parentId': parentId,
                'name': "Preset Geometry",
                'type': "Shape",
                'value': geom.get("prst")
            }
            assets.append(asset_info)


    for blip in element.findall('./*/a:blipFill/a:blip', namespaces):
        embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        asset_info = {
            'id': get_auto_id(),
            'parentId': parentId,
            'name': "Blip",
            'type': "Image",
            'value': relation.get(embed, "Rel not found")
        }
        assets.append(asset_info)
    
    for video in element.findall('./*/*/a:videoFile', namespaces):
        embed = video.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link')
        asset_info = {
            'id': get_auto_id(),
            'parentId': parentId,
            'name': "Video",
            'type': "Video",
            'value': relation.get(embed, "Rel not found")
        }
        assets.append(asset_info)
    
    for video in element.findall('./*/a:videoFile', namespaces):
        embed = video.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link')
        asset_info = {
            'id': get_auto_id(),
            'parentId': parentId,
            'name': "Video",
            'type': "Video",
            'value': relation.get(embed, "Rel not found")
        }
        assets.append(asset_info)
    
    for media in element.findall('./*/*/p:extLst/*/p14:media', namespaces):
        embed = media.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        asset_info = {
            'id': get_auto_id(),
            'parentId': parentId,
            'name': "Media",
            'type': "Media",
            'value': relation.get(embed, "Rel not found")
        }
        assets.append(asset_info)

    for graphic_data in element.findall('./a:graphicData', namespaces):
        for child in graphic_data:
            embed = media.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            asset_info = {
                'id': get_auto_id(),
                'parentId': parentId,
                'name': get_asset_tag(child),
                'type': "Graphic Data",
                'value': relation.get(embed, "Rel not found")
            }
            assets.append(asset_info)

    for layout_data in element.findall('./p:sldLayoutId', namespaces):
        id = layout_data.get('id')
        layoutId = layout_data.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        asset_info = {
            'id': id,
            'parentId': parentId,
            'name': "Layout data",
            'type': "Layout data",
            'value': relation.get(layoutId, "Rel not found")
        }
        assets.append(asset_info)

    for clrMap in element.findall('./p:clrMap', namespaces):
        for key, value in clrMap.attrib.items():
            asset_info = {
                'id': get_auto_id(),
                'parentId': parentId,
                'name': key,
                'type': "ColorMap",
                'value': value
            }
            assets.append(asset_info)

    for slide in element.findall('./p:sldId', namespaces):
        id = slide.get('id')
        layoutId = slide.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        asset_info = {
            'id': id,
            'parentId': parentId,
            'name': "Slide",
            'type': "Slide",
            'value': relation.get(layoutId, "Rel not found")
        }
        assets.append(asset_info)


# Function to extract all assets and shapes
def get_assets_and_shapes(element, assets, parentId, relation):
    get_asset(element, assets, parentId, relation)

    parentElements = {
        './p:cSld': "Slide", 
        './p:bg': "Background", 
        './p:pic': "Picture", 
        './p:spTree': "Tree", 
        './p:grpSp': "Group", 
        './p:sp': "Shape", 
        './p:graphicFrame': "Graphic", 
        './p:cxnSp': "Connector",
        './p:sldLayoutIdLst': "LayoutList",
        './p:sldIdLst': "SlideList"
    }

    for xpath in parentElements:
        for asset in element.findall(xpath, namespaces):
            idElement = asset.find('./*/p:cNvPr', namespaces)
            id = idElement.get('id') if idElement is not None else "Auto_"+str(random.randint(1001, 4999))
            name = idElement.get('name') if idElement is not None else "None"
            asset_info = {
                'id': id,
                'parentId': parentId,
                'name': name,
                'type': parentElements[xpath],
                'value': "None"
            }
            assets.append(asset_info)
            get_assets_and_shapes(asset, assets, id, relation)


# Function to extract animations and behaviors
def total_anim_count(element):
    count = 0
    for anim in element.findall('.//p:set', namespaces):
        count = count + 1
    for anim in element.findall('.//p:cmd', namespaces):
        count = count + 1
    for anim in element.findall('.//p:animEffect', namespaces):
        count = count + 1
    for anim in element.findall('.//p:anim', namespaces):
        count = count + 1
    for anim in element.findall('.//p:animClr', namespaces):
        count = count + 1
    for anim in element.findall('.//p:animMotion', namespaces):
        count = count + 1
    for anim in element.findall('.//p:animRot', namespaces):
        count = count + 1
    for anim in element.findall('.//p:animScale', namespaces):
        count = count + 1
    for anim in element.findall('.//p:transition', namespaces):
        count = count + 1
    return count

# Function to extract animations and behaviors
def total_bhvr_count(element):
    count = 0
    for anim in element.findall('.//p:cBhvr', namespaces):
        count = count + 1
    for anim in element.findall('.//p:transition', namespaces):
        count = count + 1
    return count


# Function to extract animations and behaviors
def extract_animations_and_behaviors(element, parent_map, animations):
    # for anim in element.findall('.//p:cTn', namespaces):
    #     anim_id = anim.get('id')
        
    # Check for behaviors within animations
    for behavior in element.findall('.//p:cBhvr', namespaces):
        target_element = behavior.find('./p:tgtEl/p:spTgt', namespaces)
        if target_element is not None:
            spid = target_element.get('spid')
            
            # Find the parent node of the current behavior
            parent_node = parent_map[behavior]
            if parent_node is not None:
                parent_type = parent_node.tag if parent_node is not None else None
                parent_type = parent_type.split('}')[-1] if parent_type else None  # Extract local name
                if parent_type == 'set':
                    attr_name_elem = behavior.find('./p:attrNameLst/p:attrName', namespaces)
                    if attr_name_elem is not None:
                        property_name = attr_name_elem.text  # Get the property name
                        value_elem = parent_node.find('./p:to/p:strVal', namespaces)  # Get the value being set
                        value = value_elem.get('val') if value_elem is not None else None
                        animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type,
                            'property': property_name,
                            'value': value
                        })
                    else:
                        # Append details including parent type if it's not a "set"
                        animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type
                        })
                elif parent_type == 'animEffect':
                    transition = parent_node.get('transition')
                    if transition is None:
                        transition = "None"
                    
                    filter = parent_node.get('filter')
                    if filter is None:
                        filter = "None"

                    animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type,
                            'property': 'transition_'+transition,
                            'value': 'filter_'+filter
                        })
                elif parent_type == 'anim':
                    calcmode = parent_node.get('calcmode')
                    if calcmode is None:
                        calcmode = "None"
                    
                    valueType = parent_node.get('valueType')
                    if valueType is None:
                        valueType = "None"

                    animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type,
                            'property': 'calcmode_'+calcmode,
                            'value': 'valueType_'+valueType
                        })
                elif parent_type == 'animClr':
                    clrSpc = parent_node.get('clrSpc')
                    if clrSpc is None:
                        clrSpc = "None"
                    
                    dir = parent_node.get('dir')
                    if dir is None:
                        dir = "None"

                    animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type,
                            'property': 'clrSpc_'+clrSpc,
                            'value': 'dir_'+dir
                        })
                elif parent_type == 'animMotion':
                    origin = parent_node.get('origin')
                    if origin is None:
                        origin = "None"
                    
                    pathEditMode = parent_node.get('pathEditMode')
                    if pathEditMode is None:
                        pathEditMode = "None"

                    animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type,
                            'property': 'origin_'+origin,
                            'value': 'pathEditMode_'+pathEditMode
                        })
                elif parent_type == 'animRot':
                    animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type,
                            'property': 'None',
                            'value': 'None'
                        })
                elif parent_type == 'animScale':
                    animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type,
                            'property': 'None',
                            'value': 'None'
                        })
                elif parent_type == 'cmd':
                    type = parent_node.get('type')
                    if type is None:
                        type = "None"

                    cmd = parent_node.get('cmd')
                    if cmd is None:
                        cmd = "None"

                    animations.append({
                            # 'id': anim_id,
                            'target': spid,
                            'parent_type': parent_type,
                            'property': 'type_'+type,
                            'value': 'cmd_'+cmd
                        })
                else:
                    # Append details including parent type if it's not any of the above
                    animations.append({
                        # 'id': anim_id,
                        'target': spid,
                        'parent_type': parent_type,
                        'property': "Unknown",
                        'value': "None"
                    })


# Function to extract transitions
def extract_transitions(element, transitions):
    for transition in element.findall('.//p:transition', namespaces):
        # print("Transition found")
        transition_type = transition.get('type')
        duration = transition.get('dur')
        transitions.append({
            'type': transition_type,
            'duration': duration
        })

def analyze_xml(ppt_file, slide_file_name, root, anim_writer, asset_writer, relations):
    animations = []
    transitions = []
    parent_map = {c:p for p in root.iter() for c in p}
    anim_count = total_anim_count(root)
    bhvr_count = total_bhvr_count(root)
    extract_animations_and_behaviors(root, parent_map, animations)
    extract_transitions(root, transitions)
    assets = []
    get_assets_and_shapes(root, assets, "Root", relations.get(slide_file_name))
    if (bhvr_count != anim_count):
        print(f"Animation count mismatch {len(animations)}, {len(transitions)}, {anim_count}, {bhvr_count}")

    # # print("Animations and Behaviors:")
    for anim in animations:
        output = [
            ppt_file[1], 
            slide_file_name, 
            # anim["id"], 
            anim["target"],
            anim["parent_type"], 
            anim["property"], 
            anim["value"]
        ]
        anim_writer.writerow(output)

    # print("\nTransitions:")
    for trans in transitions:
        output = [
            ppt_file[1], 
            slide_file_name, 
            # "None",
            "None", 
            "Transition",
            trans["type"], 
            trans["duration"]
        ]
        anim_writer.writerow(output)

    for asset in assets:
        output = [
            ppt_file[1], 
            slide_file_name,
            asset["id"],
            asset["parentId"],
            asset["name"],
            asset["type"],
            asset["value"] 
        ]
        asset_writer.writerow(output)

def analyze_file(ppt_file, path, typename, ppt_zip, anim_writer, asset_writer):
    relations = {}
    rel_files = [f for f in ppt_zip.namelist() if f.startswith(path+'_rels/'+typename) and f.endswith('.xml.rels')]
    for rel_file in rel_files:
        with ppt_zip.open(rel_file) as file:
            relTree = ET.parse(file)
            relRoot = relTree.getroot()
            rel_file_name = rel_file.replace(path+'_rels/', '').replace('.rels', '')
            relations[rel_file_name] = {}
            for element in relRoot.findall('./rel:Relationship', namespaces):
                relations[rel_file_name][element.get('Id')] = element.get('Target')

    # List all slide XML files in the ppt/slides directory
    type_files = [f for f in ppt_zip.namelist() if f.startswith(path+typename) and f.endswith('.xml')]

    for type_file in type_files:
        with ppt_zip.open(type_file) as file:
            type_file_name = type_file.replace(path, '')
            tree = ET.parse(file)
            root = tree.getroot()
            analyze_xml(ppt_file, type_file_name, root, anim_writer, asset_writer, relations)


def analyze_files(ppt_file, ppt_zip, anim_writer, asset_writer, presentation_writer, layout_writer):
    analyze_file(ppt_file, 'ppt/', 'presentation', ppt_zip, anim_writer, presentation_writer)
    analyze_file(ppt_file, 'ppt/slideMasters/', 'slideMaster', ppt_zip, anim_writer, presentation_writer)
    analyze_file(ppt_file, 'ppt/slideLayouts/', 'slideLayout', ppt_zip, anim_writer, layout_writer)
    analyze_file(ppt_file, 'ppt/slides/', 'slide', ppt_zip, anim_writer, asset_writer)
    

def unzip_pptx(ppt_file, anim_writer, asset_writer, presentation_writer, layout_writer):
    """Extracts and prints contents of slides from a PowerPoint file."""
    try:
        with zipfile.ZipFile(ppt_file[0], 'r') as ppt_zip:
            # print(f"Processing {ppt_file[0]}")
            analyze_files(ppt_file, ppt_zip, anim_writer, asset_writer, presentation_writer, layout_writer)
    except Exception as e:
        print(f"Error processing {ppt_file[0]}: {e}")
        traceback.print_exc()

def find_ppt_files(directory):
    """Finds all .pptx files in the given directory."""
    ppt_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.pptx') and not file.startswith('~$'):
                ppt_files.append([os.path.join(root, file), file])
    return ppt_files

def process_presentations(directory, anim_writer, asset_writer, presentation_writer, layout_writer):
    """Processes all PowerPoint presentations in the specified directory."""
    ppt_files = find_ppt_files(directory)
    if not ppt_files:
        print("No PowerPoint (.pptx) files found in the directory.")
        return
    
    for ppt_file in ppt_files:
        # print(f"Processing {ppt_file}...")
        unzip_pptx(ppt_file, anim_writer, asset_writer, presentation_writer, layout_writer)

if __name__ == "__main__":
    # Input directory path
    directory_path = str(sys.argv[1])
    
    if os.path.isdir(directory_path):
        # Writing to a CSV file
        with open('layout.csv', mode='w', newline='') as layout_file:
            with open('presentation.csv', mode='w', newline='') as presentation_file:
                with open('animation.csv', mode='w', newline='') as anim_file:
                    anim_writer = csv.writer(anim_file)
                    layout_writer = csv.writer(layout_file)
                    presentation_writer = csv.writer(presentation_file)
                    # writer = csv.writer(sys.stdout)
                    with open('asset.csv', mode='w', newline='') as asset_file:
                        asset_writer = csv.writer(asset_file)

                        output = [
                            "PPTX", 
                            "Slide", 
                            # "Anim ID",
                            "Target ID", 
                            "Animation", 
                            "Property", 
                            "Value"
                        ]
                        anim_writer.writerow(output)

                        output = [
                            "Pptx", 
                            "Slide", 
                            "Asset",
                            "Parent", 
                            "Name", 
                            "Type", 
                            "Value"
                        ]
                        asset_writer.writerow(output)
                        presentation_writer.writerow(output)
                        layout_writer.writerow(output)


                        process_presentations(directory_path, anim_writer, asset_writer, presentation_writer, layout_writer)
        #             xml_data = """\                                                                                                                                         </p:sld>
        # """
                # ppt_file = ["File", "file"]
                # root = ET.fromstring(xml_data)
                # analyze_xml(ppt_file, "slide1.xml", root, writer)
    else:
        print("Invalid directory path. Please try again.")


# python pptx.py ../../PNCF/ 