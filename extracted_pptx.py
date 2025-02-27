import os
import zipfile
import shutil

def extract_pptx(pptx_file, output_folder):
    """Extracts a .pptx file into a structured directory with slide-wise folders."""
    
    # Ensure the file exists
    if not os.path.exists(pptx_file):
        print(f"🚨 Error: File '{pptx_file}' not found.")
        return

    # Remove old output folder if exists
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder)

    # Unzip PPTX (PPTX is just a ZIP file)
    with zipfile.ZipFile(pptx_file, 'r') as zip_ref:
        zip_ref.extractall(output_folder)

    # Define paths inside extracted folder
    ppt_folder = os.path.join(output_folder, "ppt")
    slides_folder = os.path.join(ppt_folder, "slides")
    rels_folder = os.path.join(slides_folder, "_rels")
    media_folder = os.path.join(ppt_folder, "media")
    theme_folder = os.path.join(ppt_folder, "theme")
    layout_folder = os.path.join(ppt_folder, "slideLayouts")
    master_folder = os.path.join(ppt_folder, "slideMasters")

    # Ensure paths exist
    for folder in [slides_folder, rels_folder, media_folder, theme_folder, layout_folder, master_folder]:
        os.makedirs(folder, exist_ok=True)

    # Process slides and organize them
    slide_files = [f for f in os.listdir(slides_folder) if f.startswith("slide") and f.endswith(".xml")]

    for slide_file in slide_files:
        slide_number = slide_file.replace("slide", "").replace(".xml", "")
        slide_output = os.path.join(output_folder, f"slide_{slide_number}")
        os.makedirs(slide_output, exist_ok=True)

        # Move slide XML
        shutil.move(os.path.join(slides_folder, slide_file), os.path.join(slide_output, slide_file))

        # Move .rels file (if exists)
        rels_file = f"slide{slide_number}.xml.rels"
        rels_path = os.path.join(rels_folder, rels_file)
        if os.path.exists(rels_path):
            shutil.move(rels_path, os.path.join(slide_output, rels_file))

        # Move related media (images/videos)
        if os.path.exists(rels_path):
            with open(rels_path, 'r', encoding='utf-8') as f:
                rels_content = f.read()
                for media_file in os.listdir(media_folder):
                    if media_file in rels_content:
                        shutil.move(os.path.join(media_folder, media_file), os.path.join(slide_output, media_file))

    # Move global assets
    shutil.move(theme_folder, os.path.join(output_folder, "theme"))
    shutil.move(layout_folder, os.path.join(output_folder, "layouts"))
    shutil.move(master_folder, os.path.join(output_folder, "masters"))

    print(f"✅ PPTX extracted and organized in '{output_folder}'")


# ✅ Correct File Path
pptx_file = "03 Relative Motion.pptx"  # Replace with actual file path
output_folder = "extracted_pptx"

# Run the function
extract_pptx(pptx_file, output_folder)
