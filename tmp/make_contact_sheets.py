from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

src = Path("tmp/rendered_paper")
out = src / "contact_sheets"
out.mkdir(exist_ok=True)
font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 26)
pages = sorted(src.glob("page-*.png"), key=lambda p: int(p.stem.split("-")[-1]))

for i in range(0, len(pages), 2):
    group = pages[i:i+2]
    images = [Image.open(p).convert("RGB") for p in group]
    width = sum(img.width for img in images) + 30 * (len(images) + 1)
    height = max(img.height for img in images) + 80
    sheet = Image.new("RGB", (width, height), "#D9DDE3")
    draw = ImageDraw.Draw(sheet)
    x = 30
    for path, img in zip(group, images):
        draw.text((x, 12), path.stem.replace("-", " ").title(), font=font, fill="#202124")
        sheet.paste(img, (x, 60))
        x += img.width + 30
    sheet.save(out / f"pages_{i+1:02d}_{i+len(group):02d}.png")
