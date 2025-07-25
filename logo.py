from PIL import Image, ImageDraw, ImageFont

# Intel blue RGB
intel_blue = (0, 113, 197)
background = (255, 255, 255)
text = "PWQ"
sizes = [64, 128, 256, 512]
font_path = "arialbd.ttf"  # Use a bold font, or specify a path to a .ttf file

for size in sizes:
    img = Image.new("RGB", (size, size), background)
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype(font_path, int(size * 0.4))  # Reduced font size to 40% of image size
    except IOError:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((size-w)/2, (size-h)/2), text, font=font, fill=intel_blue)
    img.save(f"PWQ_logo_{size}.png")

# Generate multi-size ICO file for app icon
ico_sizes = [(64, 64), (128, 128), (256, 256)]
icons = []
for size in [s[0] for s in ico_sizes]:
    img = Image.open(f"PWQ_logo_{size}.png")
    icons.append(img)
icons[0].save("PWQ_logo.ico", format="ICO", sizes=ico_sizes)
