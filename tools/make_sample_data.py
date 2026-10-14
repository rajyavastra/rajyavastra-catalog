"""Generates the sample Excel file and illustrated saree images for RajyaVastra.
Only needed once to create demo data. Replace with your real photos/Excel."""
import math, os
import numpy as np
from PIL import Image, ImageDraw, ImageFilter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG_DIR = os.path.join(ROOT, "public", "assets", "products")
GOLD, SILVER = "#D4A93C", "#C9CED6"

# id, name, category, fabric, price, description, in_stock, style
P = [
 ("P001", "Kanjivaram Silk Saree", "Silk", "Pure silk", 389.00,
  "Rich maroon silk with a gold zari temple border and woven paisley pallu. 5.5 m saree with 0.8 m unstitched blouse piece. Dry clean only.", "yes",
  dict(body="#8E1B34", band="temple", bc="#7A1530", bz=GOLD, buti="paisley", bt=GOLD, pallu="paisley", pc=GOLD)),
 ("P002", "Banarasi Silk Saree", "Silk", "Banarasi silk", 329.00,
  "Royal blue Banarasi silk with an all-over floral jaal and a broad gold border. Includes 0.8 m blouse piece. Dry clean only.", "yes",
  dict(body="#1E3A8A", band="zari", bc="#16296A", bz=GOLD, buti="flower", bt=GOLD, pallu="flower", pc=GOLD)),
 ("P003", "Mysore Crepe Silk Saree", "Silk", "Crepe silk", 215.00,
  "Soft emerald crepe silk with a slim gold zari border. Light to drape and easy to carry for weddings and puja. Blouse piece included.", "no",
  dict(body="#0F6B4F", band="zari", bc="#0B4F3A", bz=GOLD, buti="none", bt=GOLD, pallu="stripe", pc=GOLD)),
 ("P004", "Bandhani Cotton Silk Saree", "Festive", "Cotton silk", 158.00,
  "Hand-tied bandhani dots in mustard, red and green on a cotton silk base. A bright pick for Navratri, Diwali and family functions.", "yes",
  dict(body="#E8A317", band="stripe", bc="#B3202F", bz="#FFFFFF", buti="bandhani", bt="#B3202F", pallu="bandhani", pc="#1F6B4F")),
 ("P005", "Chanderi Cotton Saree", "Cotton", "Chanderi cotton", 96.00,
  "Sheer, lightweight chanderi in pastel pink with tiny silver motifs and a subtle border. Comfortable for long days.", "yes",
  dict(body="#F1CFDA", band="zari", bc="#E3B4C4", bz=SILVER, buti="dot", bt="#A9A9B5", pallu="stripe", pc=SILVER)),
 ("P006", "Tant Cotton Saree", "Cotton", "Handloom cotton", 84.00,
  "Crisp handloom tant cotton in ivory with a bold red border and small woven motifs. Breathable and ideal for everyday wear.", "yes",
  dict(body="#F7F1E3", band="stripe", bc="#B3202F", bz="#1B1B1B", buti="flower", bt="#B3202F", pallu="stripe", pc="#B3202F")),
 ("P007", "Georgette Sequin Saree", "Party Wear", "Georgette", 142.00,
  "Flowing plum georgette with a sparkling sequin border. Lightweight and easy to drape for receptions and parties.", "yes",
  dict(body="#5B1E4D", band="sequin", bc="#3F1136", bz="#F0CF6A", buti="dot", bt="#F0CF6A", pallu="sequin", pc="#F0CF6A")),
 ("P008", "Organza Embroidered Saree", "Party Wear", "Organza", 176.00,
  "Peach organza with hand-finished floral embroidery along the border and pallu. Crisp fall with a graceful shimmer.", "yes",
  dict(body="#F4BFA9", band="floral", bc="#E9A68E", bz="#FFFFFF", buti="flower", bt="#FFFFFF", pallu="flower", pc="#FFFFFF")),
 ("P009", "Patola Style Silk Saree", "Festive", "Silk blend", 248.00,
  "Geometric double-ikat inspired weave in red, blue and gold. A statement saree for weddings and festive evenings.", "no",
  dict(body="#9B1B30", band="ikat", bc="#1E3A8A", bz=GOLD, buti="ikat", bt="#E8B93C", pallu="ikat", pc="#1E3A8A")),
 ("P010", "Printed Chiffon Saree", "Party Wear", "Chiffon", 68.00,
  "Airy aqua chiffon with a multicolour floral print. Pairs easily with a plain blouse for brunches and casual events.", "yes",
  dict(body="#BDE3EA", band="stripe", bc="#2E8FA3", bz="#FFFFFF", buti="print", bt="#E86A8D", pallu="print", pc="#E86A8D")),
]

def rgb(h):
    h = h.lstrip("#"); return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def paisley(d, cx, cy, s, fill, rot=0.0, inner=None):
    pts = []
    for t in np.linspace(0, 2*math.pi, 70):
        x = math.cos(t); y = math.sin(t) * math.sin(t/2) ** 1.6
        y += 0.28 * (1 - x) ** 2 * 0.5
        pts.append((x, y))
    def put(scale, dx, color):
        c, sn = math.cos(rot), math.sin(rot)
        poly = [(cx + scale*((x+dx)*c - y*sn), cy + scale*((x+dx)*sn + y*c)) for x, y in pts]
        d.polygon(poly, fill=color)
    put(s, 0, fill)
    if inner:
        put(s*0.55, -0.22, inner)
        put(s*0.25, -0.35, fill)

def flower(d, cx, cy, r, petal, center=None, n=6):
    for i in range(n):
        a = 2*math.pi*i/n
        px, py = cx + r*0.58*math.cos(a), cy + r*0.58*math.sin(a)
        d.ellipse((px-r*0.42, py-r*0.42, px+r*0.42, py+r*0.42), fill=petal)
    rc = r*0.28
    d.ellipse((cx-rc, cy-rc, cx+rc, cy+rc), fill=center or "#FFFFFF")

def diamond(d, cx, cy, r, fill):
    d.polygon([(cx, cy-r), (cx+r, cy), (cx, cy+r), (cx-r, cy)], fill=fill)

def motif(d, kind, cx, cy, s, col, body):
    if kind == "paisley": paisley(d, cx, cy, s*0.5, col, rot=-0.9, inner=body)
    elif kind == "flower": flower(d, cx, cy, s*0.42, col, center=body)
    elif kind == "dot": d.ellipse((cx-s*0.1, cy-s*0.1, cx+s*0.1, cy+s*0.1), fill=col)
    elif kind == "diamond": diamond(d, cx, cy, s*0.3, col)
    elif kind == "bandhani":
        for i, (dx, dy) in enumerate([(0, -1), (1, 0), (0, 1), (-1, 0), (0, 0)]):
            r = s*0.07
            c = col if i < 4 else "#FFFFFF"
            d.ellipse((cx+dx*s*0.16-r, cy+dy*s*0.16-r, cx+dx*s*0.16+r, cy+dy*s*0.16+r), fill=c)
    elif kind == "ikat":
        diamond(d, cx, cy, s*0.34, col); diamond(d, cx, cy, s*0.18, body)
        for dx in (-1, 1): d.rectangle((cx+dx*s*0.4-s*0.04, cy-s*0.04, cx+dx*s*0.4+s*0.04, cy+s*0.04), fill=col)
    elif kind == "print":
        cols = ["#E86A8D", "#F2B544", "#7A5CC2", "#FFFFFF"]
        flower(d, cx, cy, s*0.4, cols[(int(cx+cy) // 7) % 3], center=cols[3], n=5)

def band(w, h, st, k, kind=None):
    """Horizontal border band (RGBA). Vertical bands are made by rotating it."""
    kind = kind or st["band"]
    im = Image.new("RGB", (int(w), int(h)), st["bc"]); d = ImageDraw.Draw(im)
    z, hh = st["bz"], h
    if kind == "temple":
        d.rectangle((0, hh*0.42, w, hh*0.58), fill=z)
        step = hh*0.5
        for i in range(int(w/step)+2):
            x = i*step
            d.polygon([(x, hh*0.42), (x+step/2, hh*0.08), (x+step, hh*0.42)], fill=z)
            d.polygon([(x, hh*0.58), (x+step/2, hh*0.92), (x+step, hh*0.58)], fill=z)
    elif kind == "zari":
        d.rectangle((0, 0, w, hh*0.12), fill=z); d.rectangle((0, hh*0.88, w, hh), fill=z)
        d.rectangle((0, hh*0.3, w, hh*0.36), fill=z); d.rectangle((0, hh*0.64, w, hh*0.7), fill=z)
        step = hh*0.5
        for i in range(int(w/step)+2): diamond(d, i*step+step/2, hh/2, hh*0.12, z)
    elif kind == "sequin":
        step = hh*0.28
        for row in range(3):
            for i in range(int(w/step)+2):
                x, y = i*step + (step/2 if row % 2 else 0), hh*(0.25+0.25*row)
                d.ellipse((x-step*0.36, y-step*0.36, x+step*0.36, y+step*0.36), fill=z)
    elif kind == "floral":
        step = hh*0.7
        for i in range(int(w/step)+2): flower(d, i*step+step/2, hh/2, hh*0.3, z, center=st["bt"] if st["bt"] != z else "#F0CF6A")
        d.rectangle((0, 0, w, hh*0.06), fill=z); d.rectangle((0, hh*0.94, w, hh), fill=z)
    elif kind == "ikat":
        step = hh*0.5
        for i in range(int(w/step)+2):
            diamond(d, i*step+step/2, hh/2, hh*0.26, z); diamond(d, i*step+step/2, hh/2, hh*0.12, st["body"])
        d.rectangle((0, 0, w, hh*0.1), fill=z); d.rectangle((0, hh*0.9, w, hh), fill=z)
    else:  # stripe
        d.rectangle((0, hh*0.12, w, hh*0.24), fill=z); d.rectangle((0, hh*0.76, w, hh*0.88), fill=z)
        d.rectangle((0, hh*0.42, w, hh*0.58), fill=st["body"] if st["body"] != st["bc"] else z)
    return im

def saree_layer(st, k):
    W, H = int(560*k), int(800*k)
    body = rgb(st["body"])
    im = Image.new("RGB", (W, H), body); d = ImageDraw.Draw(im)
    bw = int(74*k)
    pallu_y = int(H*0.62)
    # body motifs
    step_x, step_y = 92*k, 104*k
    if st["buti"] != "none":
        row = 0; y = step_y*0.6
        while y < pallu_y - step_y*0.4:
            x = bw + step_x*0.55 + (step_x/2 if row % 2 else 0)
            while x < W - bw - step_x*0.3:
                motif(d, st["buti"], x, y, 70*k, st["bt"], st["body"]); x += step_x
            y += step_y; row += 1
    # pallu
    pb = band(W - 2*bw, int(64*k), st, k)
    im.paste(pb, (bw, pallu_y))
    y = pallu_y + int(64*k)
    kind = st["pallu"]
    if kind in ("stripe",):
        for i in range(7):
            yy = y + (14 + i*24)*k
            d.rectangle((bw, yy, W-bw, yy + (6 if i % 2 else 12)*k), fill=st["pc"] if i % 2 == 0 else st["bz"])
    else:
        row = 0; yy = y + 50*k
        while yy < H - 50*k:
            xx = bw + 56*k + (56*k if row % 2 else 0)
            while xx < W - bw - 20*k:
                if kind == "paisley": paisley(d, xx, yy, 44*k, st["pc"], rot=-0.9 if row % 2 == 0 else -2.2, inner=st["body"])
                else: motif(d, kind, xx, yy, 100*k, st["pc"], st["body"])
                xx += 112*k
            yy += 78*k; row += 1
    # side borders
    vb = band(H, bw, st, k).rotate(90, expand=True)
    im.paste(vb, (0, 0)); im.paste(vb.transpose(Image.FLIP_LEFT_RIGHT), (W-bw, 0))
    # fringe
    f = Image.new("RGBA", (W, int(26*k)), (0, 0, 0, 0)); fd = ImageDraw.Draw(f)
    light = tuple(min(255, c+40) for c in body)
    for x in range(0, W, max(3, int(5*k))): fd.line((x, 0, x + (x % 7 - 3)*0.4*k, 26*k), fill=light + (255,), width=max(1, int(2*k)))
    out = Image.new("RGBA", (W, H + int(26*k)), (0, 0, 0, 0))
    out.paste(im, (0, 0)); out.alpha_composite(f, (0, H - 2))
    return out

def fold_shade(arr, ph=0.0, diag=0.0):
    h, w = arr.shape[:2]
    xs = np.arange(w)[None, :]; ys = np.arange(h)[:, None]
    f = np.sin(xs/34.0 + diag*ys/260.0 + ph) * 0.6 + np.sin(xs/13.0 + ys/90.0) * 0.25 + np.sin(xs/71.0 + 1.3) * 0.3
    m = 0.93 + 0.11 * f
    arr[..., :3] = np.clip(arr[..., :3] * m[..., None], 0, 255)
    return arr

def hanging(st, idx):
    S = 900
    lay = saree_layer(st, 1.0)
    canvas = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    canvas.paste(lay, ((S - lay.width)//2, 60))
    a = np.array(canvas)
    for y in range(S):  # gentle flutter of the hanging cloth
        a[y] = np.roll(a[y], int(9*math.sin(y/95 + idx) + 4*math.sin(y/37)), axis=0)
    a = fold_shade(a, ph=idx, diag=1.0).astype(np.uint8)
    cloth = Image.fromarray(a)
    # wall
    wall = rgb(st["body"]); mix = tuple(int(235*0.78 + c*0.22) for c in wall)
    bg = Image.new("RGB", (S, S), mix); g = ImageDraw.Draw(bg)
    for y in range(S):
        t = y / S; g.line((0, y, S, y), fill=tuple(int(c*(1.03 - 0.1*t)) for c in mix))
    bg = bg.convert("RGBA")
    sh = Image.new("RGBA", (S, S), (0, 0, 0, 0)); sh.putalpha(cloth.split()[3].filter(ImageFilter.GaussianBlur(14)).point(lambda v: int(v*0.35)))
    sh = ImageChops_offset(sh, 16, 14)
    bg.alpha_composite(sh); bg.alpha_composite(cloth)
    d = ImageDraw.Draw(bg)  # wooden display rod
    d.rounded_rectangle((120, 40, 780, 58), 9, fill=(104, 70, 44)); d.ellipse((104, 36, 134, 62), fill=(84, 54, 33)); d.ellipse((766, 36, 796, 62), fill=(84, 54, 33))
    return bg.convert("RGB")

def ImageChops_offset(im, dx, dy):
    out = Image.new("RGBA", im.size, (0, 0, 0, 0)); out.paste(im, (dx, dy)); return out

def closeup(st, idx):
    k = 2.3
    lay = saree_layer(st, k)
    H = lay.height
    y0 = int(H*0.62) - 330
    crop = lay.crop((0, y0, 900, y0 + 900)).convert("RGB")
    a = fold_shade(np.array(crop).astype(float), ph=idx + 1.7, diag=-1.4).astype(np.uint8)
    img = Image.fromarray(a)
    v = Image.new("L", (900, 900), 0); ImageDraw.Draw(v).ellipse((-250, -250, 1150, 1150), fill=255)
    v = v.filter(ImageFilter.GaussianBlur(120))
    dark = Image.new("RGB", (900, 900), (20, 10, 12))
    return Image.composite(img, Image.blend(img, dark, 0.35), v)

os.makedirs(IMG_DIR, exist_ok=True)
for i, (pid, *_rest, st) in enumerate(P):
    folder = os.path.join(IMG_DIR, pid); os.makedirs(folder, exist_ok=True)
    hanging(st, i).save(os.path.join(folder, "1.jpg"), quality=86)
    closeup(st, i).save(os.path.join(folder, "2.jpg"), quality=86)

# ---------- Excel ----------
wb = Workbook(); ws = wb.active; ws.title = "Products"
ws.append(["id", "name", "category", "fabric", "price", "description", "in_stock", "images"])
for pid, name, cat, fab, price, desc, stock, _ in P:
    ws.append([pid, name, cat, fab, price, desc, stock, "1.jpg, 2.jpg"])
hfill = PatternFill("solid", fgColor="7A1F3D")
for c in ws[1]:
    c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = hfill; c.alignment = Alignment(vertical="center")
for row in ws.iter_rows(min_row=2):
    for c in row:
        c.font = Font(name="Arial", size=10); c.alignment = Alignment(vertical="top", wrap_text=(c.column == 6))
    row[4].number_format = '"$"#,##0.00'
for col, w in zip("ABCDEFGH", [9, 30, 14, 18, 10, 70, 10, 16]): ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"
dv = DataValidation(type="list", formula1='"yes,no"', allow_blank=False); ws.add_data_validation(dv); dv.add("G2:G500")

h = wb.create_sheet("How to use")
lines = [
 ("How to update the RajyaVastra catalog", True),
 ("Edit only the 'Products' sheet. One row per saree.", False),
 ("id: unique code, e.g. P011. Also the name of the image folder: public/assets/products/P011/", False),
 ("name, category, description: plain text shown on the site. Categories become the filter buttons.", False),
 ("fabric: e.g. Pure silk, Georgette, Handloom cotton (optional).", False),
 ("price: number only (no $ sign). Shown in the store currency.", False),
 ("in_stock: yes or no. 'no' shows a Sold out label on the site.", False),
 ("images: file names inside the product's folder, separated by commas, e.g. 1.jpg, 2.jpg. The first is the cover photo.", False),
 ("After saving, run: npm run sync   (or just npm start / npm run build, which sync automatically).", False),
 ("Rows P001 to P010 are sample data. Delete or overwrite them with your real sarees.", False),
]
for i, (t, b) in enumerate(lines, 1):
    c = h.cell(row=i, column=1, value=t); c.font = Font(name="Arial", size=12 if b else 10, bold=b)
h.column_dimensions["A"].width = 120
os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
wb.save(os.path.join(ROOT, "data", "products.xlsx")); print("done")
