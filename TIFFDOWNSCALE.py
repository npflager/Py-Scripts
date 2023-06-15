import os
from PIL import Image
from pathlib import Path

def convert_tif_to_jpg(tif_path, output_dir, dpi=150, error_list=None):
    try:
        # Check if the JPG file already exists
        output_filename = os.path.splitext(os.path.basename(tif_path))[0] + '.jpg'
        output_subdir = os.path.dirname(os.path.relpath(tif_path, directory))
        output_path = os.path.join(output_dir, output_subdir, output_filename)
        
        if os.path.exists(output_path):
            print(f"Skipping {tif_path} - Already converted")
            return

        # Open the TIF image
        image = Image.open(tif_path)

        # Calculate the dimensions for the desired resolution
        width = int(image.width * dpi / 72)
        height = int(image.height * dpi / 72)

        # Resize the image to the desired resolution
        image = image.resize((width, height))

        # Convert the image to RGB if it's not already
        if image.mode != 'RGB':
            image = image.convert('RGB')

        # Set the DPI (dots per inch) resolution
        image.info['dpi'] = (dpi, dpi)

        # Create the subdirectories within the output directory
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # Save the image as a JPG with the specified DPI
        image.save(output_path, dpi=(dpi, dpi))

        print(f"Conversion complete: {output_path}")

    except Exception as e:
        if error_list is not None:
            error_list.append((tif_path, str(e)))
        else:
            print(f"Error processing image: {tif_path}")
            print(f"Error message: {str(e)}")

def convert_directory(directory, output_dir, dpi=150):
    # Use pathlib.Path for directory traversal
    directory_path = Path(directory)
    error_list = []

    for file_path in directory_path.glob('**/*'):
        if file_path.suffix.lower() in ['.tif', '.tiff']:
            # Convert TIF files to JPG
            convert_tif_to_jpg(str(file_path), output_dir, dpi, error_list)

    # Print the list of errors
    if error_list:
        print("\n\nErrors occurred during conversion:\n")
        for tif_path, error_message in error_list:
            print(f"Image: {tif_path}")
            print(f"Error: {error_message}")
            print()

# Example usage
directory = 'D:\\RESCARTA\\RCDATA01'
output_dir = 'RCDATA01'
convert_directory(directory, output_dir)






