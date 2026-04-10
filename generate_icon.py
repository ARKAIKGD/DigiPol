import os

from PIL import Image, ImageDraw


def create_icon(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)

    size = 256
    image = Image.new("RGBA", (size, size), (46, 95, 255, 255))
    draw = ImageDraw.Draw(image)

    draw.rounded_rectangle((28, 52, 228, 204), radius=34, fill=(255, 255, 255, 255))
    draw.rectangle((72, 28, 184, 76), fill=(255, 255, 255, 255))
    draw.rounded_rectangle((94, 98, 162, 168), radius=34, fill=(46, 95, 255, 255))
    draw.rounded_rectangle((168, 138, 228, 198), radius=22, fill=(46, 95, 255, 255))

    image.save(path, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


if __name__ == "__main__":
    create_icon(os.path.join("assets", "snipit.ico"))
