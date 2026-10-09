# Obrazek podglądu repozytorium na GitHubie (Settings → Social preview), 1280×640, ze zrzutów w docs/.
# Uruchom po odświeżeniu zrzutów: python build/social_preview.py  →  docs/social-preview.png
import os
from PIL import Image, ImageDraw, ImageFilter, ImageFont

root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
W, H = 1280, 640
NAVY, YELLOW, WHITE, LIGHT = (31, 56, 100), (255, 242, 204), (255, 255, 255), (217, 225, 242)
F = "C:/Windows/Fonts/"
bold = lambda s: ImageFont.truetype(F + "arialbd.ttf", s)
reg = lambda s: ImageFont.truetype(F + "arial.ttf", s)

img = Image.new("RGB", (W, H), NAVY)
d = ImageDraw.Draw(img)


def card(path, width, pos, angle=0):
    shot = Image.open(path).convert("RGB")
    shot = shot.resize((width, round(shot.height * width / shot.width)), Image.LANCZOS)
    pad = 10
    c = Image.new("RGBA", (shot.width + 2 * pad, shot.height + 2 * pad), WHITE + (255,))
    c.paste(shot, (pad, pad))
    if angle:
        c = c.rotate(angle, expand=True, resample=Image.BICUBIC)
    sh = Image.new("RGBA", (c.width + 40, c.height + 40), (0, 0, 0, 0))
    sh.paste((0, 0, 0, 110), (20, 26, 20 + c.width, 26 + c.height), c)
    sh = sh.filter(ImageFilter.GaussianBlur(12))
    img.paste(sh, (pos[0] - 20, pos[1] - 20), sh)
    img.paste(c, pos, c)


# zrzuty: druk delegacji z tyłu, kalkulator z przodu
card(f"{root}/docs/delegacja-2.png", 250, (1000, 70), angle=-4)
card(f"{root}/docs/kalkulator.png", 640, (560, 170))

x = 64
d.rectangle((x, 92, x + 90, 100), fill=YELLOW)
d.text((x, 122), "Kalkulator diety", font=bold(58), fill=WHITE)
d.text((x, 190), "zagranicznej", font=bold(58), fill=WHITE)
d.text((x, 268), "w Excelu – z drukiem delegacji", font=reg(28), fill=LIGHT)
y = 340
for t in ("Bez makr, wszystko liczą formuły", "Kursy NBP pobierane automatycznie",
          "120 krajów – stawki z rozporządzenia", "Gotowy druk polecenia wyjazdu"):
    d.ellipse((x, y + 9, x + 12, y + 21), fill=YELLOW)
    d.text((x + 26, y), t, font=reg(25), fill=WHITE)
    y += 44
d.text((x, 572), "Licencja MIT  ·  Microsoft 365 / Excel 2024 na Windows", font=reg(19), fill=LIGHT)

img.save(f"{root}/docs/social-preview.png", optimize=True)
print("ok", img.size)
