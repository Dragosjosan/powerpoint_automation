import os
import sys
from pathlib import Path

from loguru import logger
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


logger.add(sys.stderr, format="{time} {level} {message}", colorize=True, filter="main.py", level="DEBUG")


class PowerPointTemplate:
    def __init__(self, template_path):
        self.template_path = template_path
        self.prs = Presentation(template_path)
        self.slide_map = self._create_slide_map()
        
    def _create_slide_map(self):
        slide_map = {}
        for i, slide in enumerate(self.prs.slides):
            if slide.shapes.title:
                title = slide.shapes.title.text
                slide_map[title] = {"index": i, "slide": slide}
        return slide_map
        
    def apply_data(self, data_dict):
        for slide_name, slide_data in data_dict.items():
            logger.debug(f"Try to apply data to slide: '{slide_name}'")
            if slide_name in self.slide_map:
                logger.debug(f"Slide name found in input data")
                slide = self.slide_map[slide_name]["slide"]
                
                # Process text replacements
                if "text" in slide_data:
                    logger.debug(f"Input data contains text data for slide: '{slide_name}'")
                    self._replace_text_placeholders(slide, slide_data["text"])
                
                # Process table data
                if "tables" in slide_data:
                    logger.debug(f"Input data contains table data for slide: '{slide_name}'")
                    self._update_tables(slide, slide_data["tables"])
                
                # Process images
                if "images" in slide_data:
                    logger.debug(f"Input data contains image data for slide: '{slide_name}'")
                    self._replace_images(slide, slide_data["images"])
            else:
                logger.warning(f"Slide '{slide_name}' not found in template")
    
    def _replace_text_placeholders(self, slide, text_data):
        for shape in slide.shapes:
            logger.debug(f"Processing shape: {shape.name}")

            if hasattr(shape, "text"):
                text = shape.text
                logger.debug(f"Shape containes text: '{text}'")
                for key, value in text_data.items():
                    placeholder = "{{" + key + "}}"
                    logger.debug(f"Placeholder: '{placeholder}'")
                    if placeholder in text:
                        logger.debug(f"Replacing placeholder: '{placeholder}' with '{value}'")
                        text = text.replace(placeholder, str(value))
                        
                # Update the shape text if it changed
                if text != shape.text:
                    logger.debug(f"Updating shape text with new text: '{text}'")
                    shape.text = text
    
    def _update_tables(self, slide, tables_data):
        logger.debug(f"Trying to update the tables")
        table_shapes = [shape for shape in slide.shapes if shape.has_table]
        if len(table_shapes) == 0:
            logger.warning(f"No tables found in slide")
            return
        logger.debug(f"Found {len(table_shapes)} tables in slide")

        for table_index, table_data in tables_data.items():
            # Table index can be numeric or a name identifier
            table_index = int(table_index)
            if table_index < len(table_shapes):
                table = table_shapes[table_index].table
                logger.debug(f"Found table at index {table_index}")
            else:
                # Try to find table by name attribute or placeholder text
                table = None
                for shape in table_shapes:
                    if (hasattr(shape, "name") and shape.name == table_index) or \
                       (table_data.get("identifier") and table_data["identifier"] in shape.text):
                        table = shape.table
                        break
            
            if table and "data" in table_data:
                rows = min(len(table_data["data"]), len(table.rows))
                for r in range(rows):
                    row_data = table_data["data"][r]
                    cols = min(len(row_data), len(table.columns))
                    for c in range(cols):
                        table.cell(r, c).text = str(row_data[c])
    
    def _replace_images(self, slide, images_data):
        for image_name, image_path in images_data.items():
            if not os.path.exists(image_path):
                logger.warning(f"Image file '{image_path}' not found")
                continue

            # Try to find by index (image or placeholder)
            if image_name.isdigit():
                image_index = int(image_name)
                # Try images first
                images = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]
                if image_index < len(images):
                    self._replace_single_image(slide, images[image_index], image_path)
                    continue
                # Try placeholders if no image found
                placeholders = [shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER]
                if image_index < len(placeholders):
                    self._replace_single_image(slide, placeholders[image_index], image_path)
                    continue
            else:
                # Try to find by name or alt text
                for shape in slide.shapes:
                    if (shape.shape_type == 13 or shape.shape_type == 14) and (
                        (hasattr(shape, "name") and shape.name == image_name) or
                        (hasattr(shape, "alt_text") and shape.alt_text == image_name)
                    ):
                        self._replace_single_image(slide, shape, image_path)
                        break
    
    def _replace_single_image(self, slide, shape, image_path):
        """Replace a single image or placeholder with an image, preserving position and aspect ratio."""
        from PIL import Image

        # Get placeholder position and size
        left, top, box_width, box_height = shape.left, shape.top, shape.width, shape.height

        # Remove the old shape (image or placeholder)
        sp = shape._element
        sp.getparent().remove(sp)

        # Load the new image to get its dimensions
        with Image.open(image_path) as img:
            img_width, img_height = img.size

        # Calculate scaling factor to fit image in box, preserving aspect ratio
        width_ratio = box_width / img_width
        height_ratio = box_height / img_height
        scale = min(width_ratio, height_ratio)

        new_width = int(img_width * scale)
        new_height = int(img_height * scale)

        # Center the image in the placeholder box
        new_left = left + int((box_width - new_width) / 2)
        new_top = top + int((box_height - new_height) / 2)

        # Add the new image
        slide.shapes.add_picture(image_path, new_left, new_top, new_width, new_height)
    
    def save(self, output_path):
        """Save the modified presentation"""
        self.prs.save(output_path)
        print(f"Presentation saved as {output_path}")


# Usage example
if __name__ == "__main__":
    # Load template
    template_path = "example-presentation.pptx"
    template = PowerPointTemplate(template_path)
    
    # Print all slide titles to help identify the correct names
    logger.debug("Available slides in template:")
    for title in template.slide_map.keys():
        logger.debug(f"{title}")
    
    # Define your data structure based on the template's slide titles
    # This is just an example - update with your actual slide names and desired content
    data = {
        "{{presentation_name}}": {
            "text": {
                "presentation_name": "Presentation name from python",
            }
        },
        "Content": {
            "text": {
                "text_content_name": "Content from python",
                "table_content_name": "Table from python",
                "plot_content_name": "Plot from python",
            }
        },
        "{{text_content_title}}": {
            "text": {
                "text_content_title": "Python Programming Language",
                "text_content_data": "Python is the most popular programming language in the world"
            }
        },
        "{{table_content_title}}": {
            "text": {
                "table_content_title": "Our company bitcoin profit",
            },
            "tables": {
                "0": {  # Using index 0 for the first table on the slide
                    "data": [
                        ["Quarter", "Price expected", "Price actual", "Profit"],
                        ["Q1 2025", "$0.1M", "$0.07M", "$-0.03M"],
                        ["Q2 2025", "$0.2M", "$0.07M", "$-1.03M"],
                        ["Q3 2025", "$0.3M", "$1.0M", "$0.7M"],
                        ["Q4 2025", "$0.4M", "$2.0M", "$1.6M"],
                    ]
                }
            }
        },
        "{{plot_content_title}}": {
            "text": {
                "plot_content_title": "Bitcoin price 2025",
            },
            "images": {
                "0": "my_plot.png"
            }
        }
    }
    
    logger.info("Applying data to template")
    # Apply data to template
    template.apply_data(data)
    
    # Save the result
    template.save("updated_presentation.pptx")