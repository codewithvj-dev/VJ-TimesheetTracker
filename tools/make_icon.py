"""Draws assets/icon.ico (clock + check mark). Build-time only; needs Pillow."""
import os
from PIL import Image, ImageDraw

NAVY, GREEN, WHITE = (22, 47, 106, 255), (27, 107, 58, 255), (255, 255, 255, 255)
S = 512
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)
d.rounded_rectangle([16, 16, S - 16, S - 16], radius=110, fill=NAVY)
c, r = (230, 250), 150                                     # clock face
d.ellipse([c[0] - r, c[1] - r, c[0] + r, c[1] + r], outline=WHITE, width=30)
d.line([c, (c[0], c[1] - 95)], fill=WHITE, width=30)        # hands
d.line([c, (c[0] + 75, c[1] + 40)], fill=WHITE, width=30)
d.ellipse([c[0] - 22, c[1] - 22, c[0] + 22, c[1] + 22], fill=WHITE)
b = (380, 385)                                              # check badge
d.ellipse([b[0] - 95, b[1] - 95, b[0] + 95, b[1] + 95], fill=GREEN, outline=NAVY, width=14)
d.line([(b[0] - 45, b[1] + 2), (b[0] - 10, b[1] + 38), (b[0] + 50, b[1] - 35)], fill=WHITE, width=26, joint="curve")

out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "icon.ico")
os.makedirs(os.path.dirname(out), exist_ok=True)
img.save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print(out)
