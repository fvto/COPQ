from PIL import Image

img = Image.open("legend_Picture 88.png")
print("Image size:", img.size)
# inspect top-left pixels (header)
print("Header color (x=10, y=5):", img.getpixel((10, 5)))
print("Header text color (x=50, y=5):", img.getpixel((50, 5)))
# inspect row 1 color box
print("Row 1 color box (x=10, y=25):", img.getpixel((10, 25)))
print("Row 1 text area (x=50, y=25):", img.getpixel((50, 25)))
