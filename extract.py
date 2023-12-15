import os
import shutil
import zipfile

# Input and output folder paths
input_folder = "input_folder"
output_folder = "output_folder"

# Loop through all .docx files in input folder
for filename in os.listdir(input_folder):
    if filename.endswith(".docx"):
        # Open the file as a zip archive
        with zipfile.ZipFile(os.path.join(input_folder, filename)) as archive:
            # Loop through all files in the archive
            for fileinfo in archive.infolist():
                # Check if the file is a picture
                if fileinfo.filename.startswith("word/media/") and fileinfo.filename.endswith((".jpg", ".jpeg", ".png", ".gif", ".bmp")):
                    # Extract the picture file to output folder
                    archive.extract(fileinfo.filename, path=output_folder)

print("Pictures extracted successfully!")
