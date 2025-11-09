from PIL import Image
from pathlib import Path

src = Path("assets/balanza.png")
dst = Path("assets/balanza.png")   # sobrescribe el grande

im = Image.open(src)
im.thumbnail((64, 64), Image.Resampling.LANCZOS)
im.save(dst, optimize=True)

print(dst, "->", round(dst.stat().st_size/1024, 1), "KB")
