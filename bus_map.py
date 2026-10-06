"""Draw a map of Toei buses coloured by how late they are (from a bus_delays_*.csv file).

Setup once:  python -m pip install staticmap
Run:         python bus_map.py api-samples/bus_delays_20260930-072256.csv
"""
import csv, sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from staticmap import StaticMap, CircleMarker

src = Path(sys.argv[1])
rows = list(csv.DictReader(open(src, encoding="utf-8")))
far = [r for r in rows if float(r["lon"]) < 139.6]  # Toei also runs a few buses in Ome, far west
shown = [r for r in rows if r not in far]

# Light grey base map (Esri World Light Gray Canvas, free with credit)
m = StaticMap(1600, 1150, url_template="https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}")
GREY, BLUE, NAVY = "#9b9b9b", "#2f7bd6", "#0b2e6b"

# On-time buses first, so late ones are drawn on top
for r in sorted(shown, key=lambda r: float(r["delay_min"])):
    d, pos = float(r["delay_min"]), (float(r["lon"]), float(r["lat"]))
    if d >= 5:
        m.add_marker(CircleMarker(pos, "white", 22)); m.add_marker(CircleMarker(pos, NAVY, 17))
    elif d >= 3:
        m.add_marker(CircleMarker(pos, "white", 18)); m.add_marker(CircleMarker(pos, BLUE, 13))
    else:
        m.add_marker(CircleMarker(pos, GREY, 9))
img = m.render()

# Title, legend and credits around the map
font = lambda size, bold=False: ImageFont.truetype(f"C:/Windows/Fonts/segoeui{'b' if bold else ''}.ttf", size)
late3 = sum(float(r["delay_min"]) >= 3 for r in rows)
late5 = sum(float(r["delay_min"]) >= 5 for r in rows)
page = Image.new("RGB", (1600, 1150 + 170 + 120), "white")
page.paste(img, (0, 170))
d = ImageDraw.Draw(page)
d.text((40, 30), "Where Toei buses were running late, 07:22 Tokyo time, Wed 30 Sep 2026", font=font(44, True), fill="#111")
d.text((40, 95), f"{late3} of {len(rows)} buses at least 3 min late ({round(late3 * 100 / len(rows))}%), {late5} at least 5 min late",
       font=font(32), fill="#555")
legend = [(GREY, 9, "under 3 min"), (BLUE, 13, "3 to 5 min late"), (NAVY, 17, "5+ min late")]
x = 40
for colour, size, label in legend:
    d.ellipse((x, 1150 + 170 + 22, x + size * 1.6, 1150 + 170 + 22 + size * 1.6), fill=colour)
    d.text((x + 40, 1150 + 170 + 14), label, font=font(30), fill="#333")
    x += 330
d.text((40, 1150 + 170 + 72), f"Central Tokyo shown; {len(far)} buses in Ome (western Tokyo) not shown. Data: ODPT / Toei, CC BY 4.0. "
       "Base map: Esri, HERE, Garmin, (c) OpenStreetMap contributors", font=font(22), fill="#777")
out = src.with_name(src.stem.replace("bus_delays", "bus_map") + ".png")
page.save(out)
print("saved", out)
