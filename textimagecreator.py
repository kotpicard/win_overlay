from PIL import Image, ImageFont, ImageDraw

def create_text_image(width, height, text, text_color, font_path, font_size):
    """Create image with transparent background and opaque text in specified color and font."""
    print("Drawing")
    img = create_text_image_mask(width, height, text, font_path, font_size)
    img = recolor_image_with_alpha(img, text_color)
    return img


def create_text_image_mask(width, height, text, font_path, font_size):
    """This function creates a transparent image with fully opaque white text that can be recolored for the overlay."""
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))  # create transparent image
    draw = ImageDraw.Draw(img)

    try:
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        font = ImageFont.load_default()

    x = 0
    bbox = draw.textbbox((0, 0), text + "Ag", font=font)
    # in line above Ag is added cause otherwise words with no ascenders/descenders would be positioned
    # at a different height as their bounding box is smaller that words with ascenders/descenders
    # this causes text to jump around vertically and is ugly
    text_height = bbox[3] - bbox[1]
    y = height - int(text_height*1.5)
    draw.text((x, y), text, font=font, fill=(255,255,255,255))
    return img


def recolor_image_with_alpha(img, rgb_color):
    """This function creates an image with premultiplied alpha based on an image mask."""
    r, g, b = rgb_color # color that the mask should be recolored to
    pixels = img.load()

    for y in range(img.height):
        for x in range(img.width):
            red, green, blue, alpha = pixels[x, y]
            if alpha == 0:
                continue
            alpha_factor = alpha / 255
            premultiplied_red = int(r * alpha_factor)
            premultiplied_green = int(g * alpha_factor)
            premultiplied_blue = int(b * alpha_factor)
            pixels[x, y] = (premultiplied_red, premultiplied_green, premultiplied_blue, alpha)
    return img
