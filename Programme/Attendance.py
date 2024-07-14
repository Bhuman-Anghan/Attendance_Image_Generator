import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageOps
import os
from datetime import datetime
import random

# Excel data fetch
data = pd.read_excel("F:\\Study\\Git Hub\\Zoom Attendance\\Book1.xlsm")
column_name = "Final"
column_data = data[column_name].dropna()

# Convert to numeric and remove NaN
column_data_numeric = pd.to_numeric(column_data, errors='coerce').dropna()

# Base image path
base_image_path = "F:\\Study\\Git Hub\\Zoom Attendance\\Programme\\img.jpg"
base_image = Image.open(base_image_path)

# Check for existing images
image_folder_path = "F:\\Study\\Git Hub\\Zoom Attendance\\Programme\\img\\"
matched_images = []

for item in column_data:
    possible_image_path = os.path.join(image_folder_path, f"{item}.jpg")
    if os.path.exists(possible_image_path):
        matched_images.append(possible_image_path)


random.shuffle(matched_images)

# Calculate total images
total_images = len(matched_images)

# Determine grid layout based on total images
if total_images <= 10:
    columns = 5
    rows = 3
    target_width = 300
    target_height = 300
    horizontal_spacing = 60
    vertical_spacing = 100
elif 10 < total_images <= 18:
    columns = 6
    rows = 3
    target_width = 250
    target_height = 250
    horizontal_spacing = 55
    vertical_spacing = 60
else:
    columns = 7
    rows = 4    
    target_width = 200
    target_height = 200
    horizontal_spacing = 60
    vertical_spacing = 27

border_color = "white"
border_width = 5

x, y = 80, 175

for i, img_path in enumerate(matched_images):
    matching_image = Image.open(img_path)
    resized_image = matching_image.resize((target_width, target_height), Image.LANCZOS)
    bordered_image = ImageOps.expand(resized_image, border=border_width, fill=border_color)
    
    row = i // columns
    col = i % columns
    
    x_pos = x + col * (target_width + horizontal_spacing)
    y_pos = y + row * (target_height + vertical_spacing)
    
    base_image.paste(bordered_image, (x_pos, y_pos))

# Draw text for date and total count
current_date = datetime.now().strftime("%d-%m-%Y")
total_count = len(column_data)
draw = ImageDraw.Draw(base_image)
font_path = "verdanab.ttf"
font = ImageFont.truetype(font_path, 30)
text_color = "white"

draw.text((40, 30), f"Date: {current_date}", fill=text_color, font=font)
draw.text((40, 80), f"Total Count: {total_count}", fill=text_color, font=font)

# Save the output image
output_path = f"F:\\Study\\Git Hub\\Zoom Attendance\\Final\\{current_date}.png"
base_image.save(output_path, "PNG")

print(f"Final image saved at {output_path}")
