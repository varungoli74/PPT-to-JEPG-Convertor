import os
from pptx import Presentation

def convert_ppt_to_images(input_ppt_path, output_image_dir):
    # Load the PowerPoint presentation
    presentation = Presentation(input_ppt_path)

    # Create output directory if it doesn't exist
    os.makedirs(output_image_dir, exist_ok=True)

    # Iterate through each slide and save it as an image
    for i, slide in enumerate(presentation.slides):
        image_path = os.path.join(output_image_dir, f'slide_{i + 1}.png')
        slide_image = slide.image
        with open(image_path, 'wb') as img_file:
            img_file.write(slide_image.blob)

if __name__ == "__main__":
    input_ppt_path = "/Users/varungoli/Downloads/py_ppt_compressor-master/edc-present.pptx"
    output_image_dir = 'output_images'
    convert_ppt_to_images(input_ppt_path, output_image_dir)
