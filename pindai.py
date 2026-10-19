"""Pindai: aplikasi scan ringan untuk HP Scanjet Enterprise Flow 7000 s3 (lewat driver WIA).

Hasil scan ditampilkan dulu untuk diperiksa, diputar, atau dihapus, lalu baru disimpan sebagai JPG.

Kebutuhan: Windows, Python 3.9+, pywin32, Pillow, dan pytwain (pip install pywin32 pillow pytwain),
driver HP Scanjet terpasang. Bila driver TWAIN HP ada, pengolahan gambar (deteksi ukuran, luruskan,
orientasi, warna otomatis, buang halaman kosong) dikerjakan driver HP seperti di HP Scan.
"""

import ctypes
import json
import sys
import os
import queue
import re
import shutil
import tempfile
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pythoncom
    import win32com.client
    from pywintypes import com_error
    HAS_WIA = True
except ImportError:
    HAS_WIA = False
    com_error = OSError

try:
    import twain
    HAS_TWAIN = True
except Exception:
    HAS_TWAIN = False

try:
    from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageFilter, ImageStat, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

APP_NAME = "Pindai"
# Identitas aplikasi untuk taskbar. Tanpa ini Windows menganggap jendela milik Python,
# sehingga ikon yang disematkan bercampur dengan aplikasi Python lain.
APP_ID = "PTSPSerdang.Pindai.Scanner"
ICON_FILE = Path(__file__).resolve().with_name("pindai.ico")
SETTINGS_DIR = Path(os.environ.get("APPDATA", Path.home())) / APP_NAME
SETTINGS_FILE = SETTINGS_DIR / "settings.json"
WORK_DIR = Path(tempfile.gettempdir()) / "pindai_kerja"

FORMAT_BMP = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"

# ID properti WIA (wiadef.h)
DPS_DOCUMENT_HANDLING_SELECT = 3088
DPS_PAGES = 3096
FEEDER = 0x001
DUPLEX = 0x004
IPS_CUR_INTENT = 6146
IPS_XRES = 6147
IPS_YRES = 6148
IPS_XPOS = 6149
IPS_YPOS = 6150
IPS_XEXTENT = 6151
IPS_YEXTENT = 6152
IPS_PAGE_SIZE = 3097
PAGE_AUTO = 100

# Kode TWAIN (spesifikasi TWAIN 2.x). Dipakai untuk menyalakan pengolahan bawaan driver HP.
CAP_XFERCOUNT = 1
CAP_FEEDERENABLED = 4098
CAP_AUTOFEED = 4103
CAP_INDICATORS = 4107
CAP_DUPLEXENABLED = 4115
ICAP_PIXELTYPE = 257
ICAP_UNITS = 258
ICAP_XRESOLUTION = 4376
ICAP_YRESOLUTION = 4377
ICAP_SUPPORTEDSIZES = 4386
ICAP_AUTODISCARDBLANKPAGES = 4404
ICAP_AUTOMATICBORDERDETECTION = 4432
ICAP_AUTOMATICDESKEW = 4433
ICAP_AUTOMATICROTATE = 4434
ICAP_AUTOMATICCOLORENABLED = 4441
ICAP_AUTOMATICCOLORNONCOLORPIXELTYPE = 4442
ICAP_BRIGHTNESS = 4353
ICAP_CONTRAST = 4355
TWTY_INT16, TWTY_INT32, TWTY_UINT16, TWTY_BOOL, TWTY_FIX32 = 1, 2, 4, 6, 7
TWPT_BW, TWPT_GRAY, TWPT_RGB = 0, 1, 2
TWBP_AUTO, TWBP_DISABLE = -1, -2
TWSS_NONE = 0
TWSS_BY_PAPER = {"A4 (210 x 297 mm)": 1, "Letter (8,5 x 11 in)": 3, "Legal (8,5 x 14 in)": 4}

WIA_ERRORS = {
    0x80210002: "Kertas macet. Buka penutup feeder, keluarkan kertas, lalu scan ulang dari halaman yang belum masuk.",
    0x80210005: "Scanner tidak merespons. Periksa daya dan kabel, lalu tekan Muat ulang.",
    0x80210006: "Scanner sedang dipakai aplikasi lain. Tutup aplikasi HP Scan atau sejenisnya lalu coba lagi.",
    0x80210009: "Scanner sedang pemanasan. Tunggu beberapa detik lalu coba lagi.",
    0x80210016: "Penutup scanner terbuka. Tutup lalu coba lagi.",
}
PAPER_EMPTY = 0x80210003

AUTO_COLOR = "Otomatis (warna / abu-abu)"
BW = "Hitam putih"
COLOR_MODES = {AUTO_COLOR: 1, "Berwarna": 1, "Abu-abu": 2, BW: 4}
RESOLUTIONS = ["150", "200", "300", "400", "600"]
AUTO_PAPER = "Otomatis (deteksi ukuran)"
PAPER_SIZES_IN = {
    AUTO_PAPER: None,
    "A4 (210 x 297 mm)": (8.27, 11.69),
    "F4 / Folio (215 x 330 mm)": (8.46, 12.99),
    "Legal (8,5 x 14 in)": (8.5, 14.0),
    "Letter (8,5 x 11 in)": (8.5, 11.0),
}
# Bingkai scan saat mode otomatis: selebar dan sepanjang kertas umum terbesar (Legal).
AUTO_FRAME_IN = (8.5, 14.0)
KNOWN_SIZES_IN = {
    "A4": (8.27, 11.69),
    "F4": (8.46, 12.99),
    "Legal": (8.5, 14.0),
    "Letter": (8.5, 11.0),
    "A5": (5.83, 8.27),
}
SAME_SIZE = "Sama dengan input"
OUTPUT_SIZES = {
    SAME_SIZE: None,
    "A4": KNOWN_SIZES_IN["A4"],
    "F4 / Folio": KNOWN_SIZES_IN["F4"],
    "Legal": KNOWN_SIZES_IN["Legal"],
    "Letter": KNOWN_SIZES_IN["Letter"],
}
SIDES = {"Satu sisi": FEEDER, "Bolak-balik": FEEDER | DUPLEX}
MAX_SKEW_DEG = 5.0
# Bagian halaman (setelah tinta hitam disisihkan) yang harus berwarna agar halaman disimpan berwarna.
COLOR_SHARE_MIN = 0.00004
TEMP_QUALITY = 97
ANALYSIS_DPI = 150

DEFAULTS = {
    "device_id": "",
    "side": "Bolak-balik",
    "color": AUTO_COLOR,
    "dpi": "200",
    "paper": AUTO_PAPER,
    "quality": 85,
    "skip_blank": False,
    "auto_rotate": True,
    "deskew": True,
    "out_size": SAME_SIZE,
    "brightness": 0,
    "contrast": 0,
    "folder": str(Path.home() / "Documents" / "Hasil Scan"),
    "save_format": "jpg",
}

ACCENT = "#1F5FA8"
ACCENT_ACTIVE = "#174A84"
SURFACE = "#FFFFFF"
MUTED = "#555B63"
PRIMARY_BTN = dict(bg=ACCENT, fg="white", activebackground=ACCENT_ACTIVE, activeforeground="white",
                   disabledforeground="#DDE4EE")
SECONDARY_BTN = dict(bg="#E9EAEC", fg="#1B1F24", activebackground="#DADCE0", activeforeground="#1B1F24",
                     disabledforeground="#80868E")
THUMB_W, THUMB_H = 130, 180


class ScanError(Exception):
    pass


def hresult_of(err):
    candidates = [getattr(err, "hresult", None)]
    info = getattr(err, "excepinfo", None)
    if info and len(info) > 5:
        candidates.append(info[5])
    for code in candidates:
        if isinstance(code, int) and code:
            code &= 0xFFFFFFFF
            if code & 0xFFFF0000 == 0x80210000:
                return code
    return (candidates[0] or 0) & 0xFFFFFFFF


def set_prop(props, pid, value, clamp=False):
    for i in range(1, props.Count + 1):
        p = props.Item(i)
        if p.PropertyID != pid:
            continue
        if clamp:
            try:
                value = max(p.SubTypeMin, min(value, p.SubTypeMax))
            except Exception:
                pass
        try:
            p.Value = value
            return True
        except Exception:
            return False
    return False


def find_paper(gray, tol=24):
    """Cari batas kertas terhadap latar scanner. Latar diambil dari baris paling bawah bingkai,
    yang kosong karena kertas lebih pendek dari bingkai. Hasil None berarti tidak ada tepi yang jelas."""
    h, w = len(gray), len(gray[0]) if gray else 0
    if h < 20 or w < 20:
        return None
    tail = sorted(v for row in gray[-max(3, h // 40):] for v in row)
    bg = tail[len(tail) // 2]
    if sum(1 for v in tail if abs(v - bg) <= tol) / len(tail) < 0.9:
        return None
    rows = [y for y, row in enumerate(gray) if sum(1 for v in row if abs(v - bg) > tol) > w * 0.3]
    if not rows:
        return None
    top, bottom = rows[0], rows[-1]
    band = gray[top:bottom + 1]
    cols = [x for x in range(w) if sum(1 for row in band if abs(row[x] - bg) > tol) > len(band) * 0.3]
    if not cols or bottom - top < h * 0.15:
        return None
    inside = sorted(v for row in band for v in row[cols[0]:cols[-1] + 1])
    if abs(inside[len(inside) // 2] - bg) <= tol:
        return None
    return (cols[0], top, cols[-1], bottom), bg


def size_label(w_in, h_in):
    w_in, h_in = sorted((w_in, h_in))
    for name, (kw, kh) in KNOWN_SIZES_IN.items():
        if abs(w_in - kw) <= 0.3 and abs(h_in - kh) <= 0.3:
            return name
    return f"{round(w_in * 25.4)} x {round(h_in * 25.4)} mm"


def locate_paper(gray):
    """Batas kertas di gambar penuh (left, top, right, bottom) dan tingkat abu latar scanner, atau None."""
    small = gray.copy()
    small.thumbnail((200, 400), Image.Resampling.BOX)
    w = small.width
    data = list(small.getdata())
    found = find_paper([data[i:i + w] for i in range(0, len(data), w)])
    if not found:
        return None
    (left, top, right, bottom), bg = found
    sx, sy = gray.width / small.width, gray.height / small.height
    box = (int((left + 0.5) * sx), int((top + 0.5) * sy),
           min(gray.width, int((right + 0.5) * sx)), min(gray.height, int((bottom + 0.5) * sy)))
    return box, bg


def paper_level(gray):
    hist = gray.histogram()
    total, acc = sum(hist), 0
    for v in range(255, -1, -1):
        acc += hist[v]
        if acc >= total * 0.5:
            return v
    return 255


def profile_var(values):
    mean = sum(values) / len(values)
    return sum((v - mean) ** 2 for v in values) / len(values)


def detect_skew(gray):
    """Sudut (derajat, berlawanan jarum jam) yang membuat baris teks paling lurus, atau 0 bila tidak yakin.
    Dicari dengan memutar masker tinta dan memilih sudut yang profil barisnya paling tajam."""
    mask = ink_mask(gray)
    if mask.histogram()[255] < mask.width * mask.height * 0.002:
        return 0.0

    def score(angle):
        r = mask.rotate(angle, resample=Image.Resampling.NEAREST, fillcolor=0) if angle else mask
        rows = list(r.resize((1, r.height), Image.Resampling.BOX).getdata())
        cols = list(r.resize((r.width, 1), Image.Resampling.BOX).getdata())
        return profile_var(rows) + profile_var(cols)

    base = score(0)
    best_angle, best = 0.0, base
    steps = int(MAX_SKEW_DEG / 0.5)
    for i in range(-steps, steps + 1):
        angle = i * 0.5
        value = score(angle)
        if value > best:
            best_angle, best = angle, value
    for i in range(-4, 5):
        angle = round(best_angle + i * 0.1, 2)
        value = score(angle)
        if value > best:
            best_angle, best = angle, value
    if best < base * 1.03 or abs(best_angle) < 0.15:
        return 0.0
    return best_angle


def color_share(im, dpi):
    """Porsi halaman yang benar-benar berwarna. Warna kertas dinetralkan dulu, dan pinggiran warna
    di sekitar tinta hitam (efek sensor scanner) disisihkan agar teks hitam tidak terbaca berwarna."""
    small = im.convert("RGB")
    k = min(1.0, ANALYSIS_DPI / dpi)
    small = small.resize((max(1, int(small.width * k)), max(1, int(small.height * k))), Image.Resampling.BOX)
    w, h = small.size
    small = small.crop((int(w * .02), int(h * .02), int(w * .98), int(h * .98)))
    bands = [b.point(lambda v, g=max(0.8, min(1.5, 245 / max(m, 1))): min(255, int(v * g)))
             for b, m in zip(small.split(), ImageStat.Stat(small).median)]
    r, g, b = bands
    hi = ImageChops.lighter(ImageChops.lighter(r, g), b)
    lo = ImageChops.darker(ImageChops.darker(r, g), b)
    chroma = ImageChops.difference(hi, lo)
    black_ink = ImageChops.multiply(hi.point(lambda v: 255 if v < 120 else 0),
                                    chroma.point(lambda v: 255 if v < 40 else 0))
    colored = ImageChops.subtract(chroma.point(lambda v: 255 if v >= 40 else 0),
                                  black_ink.filter(ImageFilter.MaxFilter(7)))
    return colored.histogram()[255] / max(1, small.width * small.height)


def process_page(path, opts, crop, driver_auto):
    """Olah satu halaman mentah: luruskan, potong ke tepi kertas, atur pencahayaan, pilih warna/abu-abu.
    Menimpa berkas di path dan mengembalikan keterangan untuk ditampilkan."""
    dpi = opts["dpi"]
    with Image.open(path) as src:
        im = src.convert("L" if src.mode in ("1", "L", "P") else "RGB")
    gray = im if im.mode == "L" else im.convert("L")
    found = locate_paper(gray) if crop else None

    skew = 0.0
    if opts["deskew"]:
        region = gray.crop(found[0]) if found else gray
        w, h = region.size
        region = region.crop((int(w * .04), int(h * .04), int(w * .96), int(h * .96)))
        k = 100 / dpi
        if k < 1:
            region = region.resize((max(1, int(region.width * k)), max(1, int(region.height * k))),
                                   Image.Resampling.BOX)
        skew = detect_skew(region)
    if skew:
        if found:
            level = found[1]
            fill = level if im.mode == "L" else (level, level, level)
            im = im.rotate(skew, resample=Image.Resampling.BICUBIC, expand=True, fillcolor=fill)
            found = locate_paper(im if im.mode == "L" else im.convert("L"))
        else:
            level = paper_level(gray)
            fill = level if im.mode == "L" else (level, level, level)
            im = im.rotate(skew, resample=Image.Resampling.BICUBIC, expand=False, fillcolor=fill)
    if found:
        # Pangkas 1 mm per sisi agar sisa garis latar scanner di tepi kertas tidak ikut terlihat.
        inset = max(2, round(dpi / 25.4))
        left, top, right, bottom = found[0]
        im = im.crop((left + inset, top + inset, max(left + inset + 1, right - inset), max(top + inset + 1, bottom - inset)))
        if skew:
            im = clean_border(im, found[1])

    size = None
    if opts["paper"] is None:
        if found or driver_auto:
            size = size_label(im.width / dpi, im.height / dpi)
        else:
            size = "ukuran tidak terdeteksi"

    if not opts["bw"] and (opts["brightness"] or opts["contrast"]):
        if opts["brightness"]:
            im = ImageEnhance.Brightness(im).enhance(1 + opts["brightness"] / 100)
        if opts["contrast"]:
            im = ImageEnhance.Contrast(im).enhance(1 + opts["contrast"] / 100)

    tone = None
    if opts["auto_color"] and im.mode == "RGB":
        if color_share(im, dpi) > COLOR_SHARE_MIN:
            tone = "berwarna"
        else:
            im = im.convert("L")
            tone = "abu-abu"

    im.save(path, "JPEG", quality=TEMP_QUALITY, dpi=(dpi, dpi))
    return {"size": size, "skew": skew, "tone": tone}


def clean_border(im, backing, tol=28):
    """Isi sisa latar scanner yang menempel di tepi gambar dengan warna kertas. Sisa ini muncul bila
    pojok kertas miring sempat keluar dari bidang pindai lalu halaman diluruskan."""
    gray = im if im.mode == "L" else im.convert("L")
    paper = paper_level(gray)
    if abs(paper - backing) <= tol:
        return im
    fill = paper if im.mode == "L" else (paper, paper, paper)
    w, h = im.size
    step = max(8, min(w, h) // 120)
    seeds = [(x, y) for x in range(0, w, step) for y in (0, h - 1)]
    seeds += [(x, y) for y in range(0, h, step) for x in (0, w - 1)]
    for seed in seeds:
        v = gray.getpixel(seed)
        if abs(v - backing) <= tol and abs(v - paper) > tol:
            ImageDraw.floodfill(gray, seed, paper, thresh=tol)
            if im is not gray:
                ImageDraw.floodfill(im, seed, fill, thresh=tol * 3)
    return im


def fit_page(im, size_in, dpi):
    """Letakkan gambar di tengah halaman berukuran tetap. Gambar yang lebih besar diperkecil, tidak diperbesar."""
    pw, ph = round(size_in[0] * dpi), round(size_in[1] * dpi)
    if im.width > im.height:
        pw, ph = ph, pw
    k = min(1.0, pw / im.width, ph / im.height)
    if k < 0.995:
        im = im.resize((max(1, round(im.width * k)), max(1, round(im.height * k))), Image.Resampling.LANCZOS)
    page = Image.new(im.mode, (pw, ph), 255 if im.mode == "L" else (255, 255, 255))
    page.paste(im, ((pw - im.width) // 2, (ph - im.height) // 2))
    return page


def ink_mask(gray):
    hist = gray.histogram()
    total, acc, paper = sum(hist), 0, 255
    for v in range(255, -1, -1):
        acc += hist[v]
        if acc >= total * 0.5:
            paper = v
            break
    limit = int(paper * 0.6)
    return gray.point(lambda v: 255 if v < limit else 0)


def strip_profiles(mask, strips=10):
    """Profil tinta per baris, dihitung di jalur vertikal sempit supaya kertas yang sedikit miring
    tidak mengaburkan celah antarbaris."""
    w, h = mask.size
    for s in range(strips):
        part = mask.crop((w * s // strips, 0, w * (s + 1) // strips, h))
        yield list(part.resize((1, h), Image.Resampling.BOX).getdata())


def gap_ratio(mask):
    gaps = rows = 0
    for prof in strip_profiles(mask):
        ink = [i for i, v in enumerate(prof) if v > 3]
        if len(ink) < 10:
            continue
        span = prof[ink[0]:ink[-1] + 1]
        gaps += sum(1 for v in span if v <= 3)
        rows += len(span)
    return gaps / rows if rows else 0.0


def line_balance(mask):
    """Positif bila tinta di atas badan huruf (huruf tinggi, titik i) lebih banyak dari di bawahnya
    (ekor g, p, y). Pada teks Latin yang tegak, bagian atas selalu lebih padat."""
    above = below = lines = 0
    for prof in strip_profiles(mask):
        vals = sorted(v for v in prof if v > 0)
        if len(vals) < 20:
            continue
        limit = vals[int(len(vals) * 0.9)] * 0.4
        cores, y, n = [], 0, len(prof)
        while y < n:
            if prof[y] < limit:
                y += 1
                continue
            a = y
            while y < n and prof[y] >= limit:
                y += 1
            if y - a >= 3:
                cores.append((a, y - 1))
        for i, (a, b) in enumerate(cores):
            reach = max(2, int((b - a + 1) * 0.6))
            up_end = (cores[i - 1][1] + a) // 2 if i else 0
            down_end = (b + cores[i + 1][0]) // 2 if i + 1 < len(cores) else n - 1
            above += sum(prof[max(a - 1 - reach, up_end):max(a - 1, 0)])
            below += sum(prof[b + 2:min(b + 2 + reach, down_end + 1)])
            lines += 1
    if lines < 6 or above + below == 0:
        return 0.0
    return (above - below) / (above + below)


def detect_rotation(path, dpi):
    """Sudut putar searah jarum jam (0, 90, 180, 270) agar teks tegak. 0 bila tidak yakin."""
    with Image.open(path) as im:
        w, h = im.size
        k = ANALYSIS_DPI / dpi
        im.draft("L", (int(w * k), int(h * k)))
        gray = im.convert("L")
    if gray.width > w * k * 1.2:
        gray = gray.resize((max(1, int(w * k)), max(1, int(h * k))), Image.Resampling.BOX)
    mask = ink_mask(gray)
    rows_gap = gap_ratio(mask)
    cols_gap = gap_ratio(mask.transpose(Image.Transpose.ROTATE_90))
    if cols_gap > rows_gap + 0.08:
        score = line_balance(mask.transpose(Image.Transpose.ROTATE_270))
        if abs(score) < 0.05:
            return 0
        return 90 if score > 0 else 270
    return 180 if line_balance(mask) < -0.1 else 0


def is_blank(thumb):
    """Perkiraan: halaman dianggap kosong kalau hampir tidak ada piksel yang lebih gelap dari kertasnya."""
    g = thumb.convert("L")
    w, h = g.size
    g = g.crop((int(w * 0.06), int(h * 0.06), int(w * 0.94), int(h * 0.94)))
    hist = g.histogram()
    total = sum(hist)
    if not total:
        return False
    acc, paper = 0, 255
    for v in range(256):
        acc += hist[v]
        if acc >= total / 2:
            paper = v
            break
    dark = sum(hist[:max(0, paper - 40)])
    return dark / total < 0.004


TRANSPOSE = {}
if HAS_PIL:
    TRANSPOSE = {90: Image.Transpose.ROTATE_270, 180: Image.Transpose.ROTATE_180, 270: Image.Transpose.ROTATE_90}


def rotated(im, angle):
    return im.transpose(TRANSPOSE[angle]) if angle else im


def make_thumb(path):
    with Image.open(path) as im:
        im.draft("RGB", (THUMB_W * 2, THUMB_H * 2))
        im = im.convert("RGB")
        im.thumbnail((THUMB_W, THUMB_H), Image.Resampling.LANCZOS)
        return im.copy()


def load_settings():
    data = dict(DEFAULTS)
    try:
        data.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    for key, allowed in (("paper", PAPER_SIZES_IN), ("color", COLOR_MODES), ("out_size", OUTPUT_SIZES)):
        if data.get(key) not in allowed:
            data[key] = DEFAULTS[key]
    return data


def save_settings(data):
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        SETTINGS_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError:
        pass


def clean_name(text):
    return re.sub(r'[\\/:*?"<>|]+', "_", text.strip()).strip(". ")


def list_scanners(root):
    """Daftar (nama tampilan, id). Driver HP lewat TWAIN didahulukan karena pengolahan gambarnya
    sama dengan aplikasi HP Scan; WIA menjadi cadangan."""
    found = []
    if HAS_TWAIN:
        try:
            sm = twain.SourceManager(root)
            try:
                for name in sm.source_list:
                    if "hp" in name.lower() or "scanjet" in name.lower():
                        found.append((f"{name} (driver HP)", f"twain:{name}"))
            finally:
                sm.close()
        except Exception:
            pass
    if HAS_WIA:
        try:
            dm = win32com.client.Dispatch("WIA.DeviceManager")
            wia = []
            for i in range(1, dm.DeviceInfos.Count + 1):
                info = dm.DeviceInfos.Item(i)
                if info.Type == 1:
                    wia.append((f"{info.Properties.Item('Name').Value} (WIA)", f"wia:{info.DeviceID}"))
            wia.sort(key=lambda d: "7000" not in d[0])
            found += wia
        except com_error:
            pass
    return found


def hidden_window():
    """Jendela Windows tak terlihat milik thread scan, tempat driver TWAIN mengirim pesannya."""
    try:
        from ctypes import wintypes
        user32 = ctypes.windll.user32
    except (AttributeError, OSError):
        return None
    user32.CreateWindowExW.restype = wintypes.HWND
    user32.CreateWindowExW.argtypes = [wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
                                       ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                       wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID]
    return user32.CreateWindowExW(0, "STATIC", "Pindai TWAIN", 0, 0, 0, 0, 0, None, None, None, None)


def destroy_window(hwnd):
    if hwnd:
        try:
            ctypes.windll.user32.DestroyWindow(hwnd)
        except (AttributeError, OSError):
            pass


def configure_twain(src, opts):
    """Atur driver HP agar mengolah gambar seperti profil HP Scan. Mengembalikan daftar fitur yang ditolak driver."""
    refused = []

    def cap(label, cap_id, type_id, value):
        try:
            src.set_capability(cap_id, type_id, value)
        except Exception:
            refused.append(label)

    cap("indikator", CAP_INDICATORS, TWTY_BOOL, False)
    cap("feeder", CAP_FEEDERENABLED, TWTY_BOOL, True)
    cap("umpan otomatis", CAP_AUTOFEED, TWTY_BOOL, True)
    cap("bolak-balik", CAP_DUPLEXENABLED, TWTY_BOOL, bool(opts["handling"] & DUPLEX))
    cap("jumlah halaman", CAP_XFERCOUNT, TWTY_INT16, -1)
    pixel = TWPT_BW if opts["bw"] else TWPT_GRAY if opts["intent"] == 2 else TWPT_RGB
    cap("mode warna", ICAP_PIXELTYPE, TWTY_UINT16, pixel)
    cap("resolusi", ICAP_XRESOLUTION, TWTY_FIX32, float(opts["dpi"]))
    cap("resolusi", ICAP_YRESOLUTION, TWTY_FIX32, float(opts["dpi"]))
    if opts["auto_color"]:
        cap("warna otomatis", ICAP_AUTOMATICCOLORENABLED, TWTY_BOOL, True)
        cap("warna otomatis", ICAP_AUTOMATICCOLORNONCOLORPIXELTYPE, TWTY_UINT16, TWPT_GRAY)
    else:
        cap("warna otomatis", ICAP_AUTOMATICCOLORENABLED, TWTY_BOOL, False)

    if opts["paper"] is None:
        cap("ukuran kertas", ICAP_SUPPORTEDSIZES, TWTY_UINT16, TWSS_NONE)
        cap("deteksi ukuran", ICAP_AUTOMATICBORDERDETECTION, TWTY_BOOL, True)
    else:
        cap("deteksi ukuran", ICAP_AUTOMATICBORDERDETECTION, TWTY_BOOL, False)
        code = TWSS_BY_PAPER.get(opts["paper_name"])
        if code is not None:
            cap("ukuran kertas", ICAP_SUPPORTEDSIZES, TWTY_UINT16, code)
        else:
            cap("ukuran kertas", ICAP_SUPPORTEDSIZES, TWTY_UINT16, TWSS_NONE)
            cap("ukuran kertas", ICAP_UNITS, TWTY_UINT16, 0)
            try:
                w_in, h_in = opts["paper"]
                src.set_image_layout((0.0, 0.0, float(w_in), float(h_in)))
            except Exception:
                refused.append("ukuran kertas")

    cap("luruskan", ICAP_AUTOMATICDESKEW, TWTY_BOOL, bool(opts["deskew"]))
    cap("orientasi otomatis", ICAP_AUTOMATICROTATE, TWTY_BOOL, bool(opts["auto_rotate"]))
    cap("buang halaman kosong", ICAP_AUTODISCARDBLANKPAGES, TWTY_INT32,
        TWBP_AUTO if opts["skip_blank"] else TWBP_DISABLE)
    if not opts["bw"]:
        # Rentang driver HP -1000..1000; slider Pindai -50..50.
        cap("kecerahan", ICAP_BRIGHTNESS, TWTY_FIX32, float(opts["brightness"] * 20))
        cap("kontras", ICAP_CONTRAST, TWTY_FIX32, float(opts["contrast"] * 20))
    return sorted(set(refused))


def twain_error_text(err):
    name = type(err).__name__
    known = {
        "PaperJam": "Kertas macet. Buka penutup feeder, keluarkan kertas, lalu scan ulang dari halaman yang belum masuk.",
        "PaperDoubleFeedError": "Dua lembar tertarik sekaligus. Rapikan tumpukan kertas lalu scan ulang halaman terakhir.",
        "CheckDeviceOnlineError": "Scanner tidak merespons. Periksa daya dan kabel USB, lalu tekan Muat ulang.",
        "DeniedError": "Scanner sedang dipakai aplikasi lain. Tutup HP Scan atau aplikasi scan lain lalu coba lagi.",
        "MaxConnectionsError": "Scanner sedang dipakai aplikasi lain. Tutup HP Scan atau aplikasi scan lain lalu coba lagi.",
    }
    return known.get(name, f"Driver HP menolak perintah ({name}). Matikan lalu nyalakan scanner, kemudian coba lagi.")


def finish_twain_page(raw, temp, opts):
    """Ubah halaman BMP dari driver HP menjadi JPG sementara. Pengolahan sudah dikerjakan driver."""
    dpi = opts["dpi"]
    with Image.open(raw) as src:
        if src.mode == "RGB":
            im = src.copy()
        elif src.mode in ("1", "L"):
            im = src.convert("L")
        else:
            im = src.convert("RGB")
    if im.mode == "RGB":
        # Driver kadang mengirim halaman abu-abu dalam format 24-bit; simpan sebagai abu-abu agar ringan.
        r, g, b = im.split()
        if ImageChops.difference(r, g).getbbox() is None and ImageChops.difference(g, b).getbbox() is None:
            im = im.convert("L")
    im.save(temp, "JPEG", quality=TEMP_QUALITY, dpi=(dpi, dpi))
    raw.unlink(missing_ok=True)
    size = size_label(im.width / dpi, im.height / dpi) if opts["paper"] is None else None
    tone = None
    if opts["auto_color"]:
        tone = "abu-abu" if im.mode == "L" else "berwarna"
    return {"size": size, "skew": 0.0, "tone": tone, "engine": "hp"}


def twain_worker(source_name, opts, out_q, stop_event):
    """Scan lewat driver TWAIN HP di thread sendiri. Pesan antrean sama dengan scan_worker."""
    count = 0
    hwnd = hidden_window()
    sm = src = None
    try:
        WORK_DIR.mkdir(parents=True, exist_ok=True)
        sm = twain.SourceManager(hwnd)
        src = sm.open_source(source_name)
        if src is None:
            raise ScanError("Driver HP tidak bisa dibuka. Tekan Muat ulang lalu coba lagi.")
        refused = configure_twain(src, opts)
        if refused:
            out_q.put(("note", "Driver HP tidak menerima pengaturan: " + ", ".join(refused) + "."))
        pending = {}

        def before(_info):
            if stop_event.is_set():
                raise twain.exceptions.CancelAll()
            pending["raw"] = WORK_DIR / f"{opts['token']}_{len(pending.get('done', [])) + 1:03d}.bmp"
            return str(pending["raw"])

        def after(more):
            nonlocal count
            count += 1
            raw = pending.pop("raw")
            temp = raw.with_suffix(".jpg")
            info = finish_twain_page(raw, temp, opts)
            pending.setdefault("done", []).append(temp)
            try:
                thumb = make_thumb(temp)
            except OSError:
                thumb = None
            out_q.put(("page", {"temp": temp, "thumb": thumb, "angle": 0, "auto": 0,
                                "dpi": opts["dpi"], **info}))
            if stop_event.is_set() and more:
                raise twain.exceptions.CancelAll()

        src.acquire_file(before=before, after=after, show_ui=False, modal=False)
        if count == 0 and not stop_event.is_set():
            raise ScanError("Tidak ada halaman yang masuk. Pastikan dokumen ada di feeder, lalu tekan Scan lagi. "
                             "Bila Buang halaman kosong aktif, semua halaman mungkin terdeteksi kosong.")
    except ScanError as err:
        out_q.put(("error", str(err)))
    except Exception as err:
        if HAS_TWAIN and isinstance(err, twain.exceptions.CancelAll):
            pass
        else:
            out_q.put(("error", twain_error_text(err)))
    finally:
        for obj in (src, sm):
            try:
                if obj is not None:
                    obj.close()
            except Exception:
                pass
        destroy_window(hwnd)
        out_q.put(("done", count))


def scan_worker(device_id, opts, out_q, stop_event):
    """Jalan di thread terpisah. Mengirim ('page', dict), ('blank',), ('error', pesan), ('done', jumlah).
    Halaman hanya disimpan sementara di WORK_DIR; penyimpanan akhir terjadi saat pengguna menekan Simpan."""
    pythoncom.CoInitialize()
    count = 0
    try:
        dm = win32com.client.Dispatch("WIA.DeviceManager")
        dev = None
        for i in range(1, dm.DeviceInfos.Count + 1):
            info = dm.DeviceInfos.Item(i)
            if info.DeviceID == device_id:
                dev = info.Connect()
                break
        if dev is None:
            raise ScanError("Scanner yang dipilih tidak ditemukan. Tekan Muat ulang.")

        set_prop(dev.Properties, DPS_DOCUMENT_HANDLING_SELECT, opts["handling"])
        set_prop(dev.Properties, DPS_PAGES, 1)

        item = dev.Items.Item(1)
        # Intent diset dulu karena driver bisa mereset resolusi saat intent berubah.
        set_prop(item.Properties, IPS_CUR_INTENT, opts["intent"])
        set_prop(item.Properties, IPS_XRES, opts["dpi"])
        set_prop(item.Properties, IPS_YRES, opts["dpi"])
        set_prop(item.Properties, IPS_XPOS, 0)
        set_prop(item.Properties, IPS_YPOS, 0)
        auto = opts["paper"] is None
        driver_auto = False
        if auto:
            driver_auto = (set_prop(item.Properties, IPS_PAGE_SIZE, PAGE_AUTO)
                           or set_prop(dev.Properties, IPS_PAGE_SIZE, PAGE_AUTO))
        if not driver_auto:
            w_in, h_in = AUTO_FRAME_IN if auto else opts["paper"]
            set_prop(item.Properties, IPS_XEXTENT, int(w_in * opts["dpi"]), clamp=True)
            set_prop(item.Properties, IPS_YEXTENT, int(h_in * opts["dpi"]), clamp=True)

        WORK_DIR.mkdir(parents=True, exist_ok=True)
        while not stop_event.is_set():
            try:
                img = item.Transfer(FORMAT_BMP)
            except com_error as err:
                code = hresult_of(err)
                if code == PAPER_EMPTY:
                    if count == 0:
                        raise ScanError("Feeder kosong. Masukkan dokumen ke feeder lalu tekan Scan lagi.")
                    break
                if code in WIA_ERRORS:
                    raise ScanError(WIA_ERRORS[code])
                raise ScanError(f"Scan gagal (kode {code:#010x}). Matikan lalu nyalakan scanner, kemudian coba lagi.")

            count += 1
            temp = WORK_DIR / f"{opts['token']}_{count:03d}.jpg"
            to_jpg = win32com.client.Dispatch("WIA.ImageProcess")
            to_jpg.Filters.Add(to_jpg.FilterInfos.Item("Convert").FilterID)
            to_jpg.Filters.Item(1).Properties.Item("FormatID").Value = FORMAT_JPEG
            to_jpg.Filters.Item(1).Properties.Item("Quality").Value = TEMP_QUALITY
            to_jpg.Apply(img).SaveFile(str(temp))
            info = process_page(temp, opts, crop=auto and not driver_auto, driver_auto=driver_auto)

            try:
                thumb = make_thumb(temp)
            except OSError:
                thumb = None
            if opts["skip_blank"] and thumb is not None and is_blank(thumb):
                temp.unlink(missing_ok=True)
                out_q.put(("blank",))
                continue
            angle = 0
            if opts["auto_rotate"]:
                try:
                    angle = detect_rotation(temp, opts["dpi"])
                except (OSError, ValueError):
                    angle = 0
            out_q.put(("page", {"temp": temp, "thumb": thumb, "angle": angle, "auto": angle,
                                "dpi": opts["dpi"], **info}))
    except ScanError as err:
        out_q.put(("error", str(err)))
    except com_error as err:
        out_q.put(("error", f"Scanner menolak perintah (kode {hresult_of(err):#010x}). Coba resolusi atau ukuran kertas lain."))
    except OSError as err:
        out_q.put(("error", f"Tidak bisa menulis berkas sementara: {err.strerror}."))
    finally:
        out_q.put(("done", count))
        pythoncom.CoUninitialize()


def prepare_page(path, angle, out_size, dpi):
    with Image.open(path) as im:
        out = rotated(im, angle)
        if out.mode not in ("RGB", "L"):
            out = out.convert("RGB")
        if out is im:
            out = im.copy()
    if out_size:
        out = fit_page(out, out_size, dpi)
    return out


def save_worker(pages, folder, base, quality, out_size, fmt, out_q):
    """Tulis halaman ke folder tujuan sesuai urutan dan putaran yang dipilih pengguna, sebagai JPG per
    halaman atau satu PDF. Mengirim ('saved', jumlah) per halaman lalu ('save_done', jumlah, error, nama)."""
    saved, error, name = 0, None, ""
    try:
        folder.mkdir(parents=True, exist_ok=True)
        if fmt == "pdf":
            final = folder / f"{base}.pdf"
            n = 2
            while final.exists():
                final = folder / f"{base}_{n}.pdf"
                n += 1
            part = folder / f"~{final.stem}.pdf.part"
            try:
                for i, page in enumerate(pages):
                    out = prepare_page(page["temp"], page["angle"], out_size, page["dpi"])
                    # Halaman ditambahkan satu per satu supaya memori tetap kecil untuk dokumen tebal.
                    out.save(part, "PDF", append=i > 0, resolution=float(page["dpi"]), quality=quality,
                             title=final.stem, creator=APP_NAME)
                    out_q.put(("saved", i + 1))
                os.replace(part, final)
            except OSError:
                part.unlink(missing_ok=True)
                raise
            saved, name = len(pages), final.name
        else:
            pattern = re.compile(re.escape(base) + r"_(\d+)\.jpg$", re.IGNORECASE)
            used = [int(m.group(1)) for f in os.listdir(folder) if (m := pattern.match(f))]
            number = max(used, default=0) + 1
            first = number
            for page in pages:
                final = folder / f"{base}_{number:03d}.jpg"
                out = prepare_page(page["temp"], page["angle"], out_size, page["dpi"])
                out.save(final, "JPEG", quality=quality, optimize=True, dpi=(page["dpi"], page["dpi"]))
                number += 1
                saved += 1
                out_q.put(("saved", saved))
            name = f"{base}_{first:03d}.jpg" + (f" sampai {base}_{number - 1:03d}.jpg" if saved > 1 else "")
    except OSError as err:
        what = "PDF" if fmt == "pdf" else f"halaman ke-{saved + 1}"
        error = f"Gagal menyimpan {what}: {err.strerror or err}. Periksa folder tujuan lalu tekan Simpan lagi."
        if fmt == "pdf":
            saved = 0
    out_q.put(("save_done", saved, error, name))


class Viewer(tk.Toplevel):
    """Tampilan besar satu halaman untuk diperiksa sebelum disimpan."""

    def __init__(self, app, index):
        super().__init__(app.root)
        self.app = app
        self.index = index
        self.photo = None
        self._pending = None
        sh = self.winfo_screenheight()
        h = int(sh * 0.85)
        self.geometry(f"{max(640, int(h * 0.78))}x{h}")
        self.minsize(640, 480)
        self.transient(app.root)

        self.canvas = tk.Canvas(self, bg="#3A3D42", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        bar = ttk.Frame(self, padding=(10, 6, 10, 10))
        bar.pack(fill="x")
        self.var_info = tk.StringVar()
        ttk.Label(bar, textvariable=self.var_info).pack(anchor="w", pady=(0, 6))
        row = ttk.Frame(bar)
        row.pack(fill="x")
        ttk.Button(row, text="\u2039 Sebelumnya", command=lambda: self.go(-1)).pack(side="left")
        ttk.Button(row, text="Berikutnya \u203a", command=lambda: self.go(1)).pack(side="left", padx=(4, 0))
        ttk.Button(row, text="Tutup (Esc)", command=self.destroy).pack(side="right")
        ttk.Button(row, text="Hapus (Del)", command=self.delete).pack(side="right", padx=(0, 12))
        ttk.Button(row, text="\u21bb Kanan (R)", command=lambda: self.rotate(90)).pack(side="right", padx=(0, 4))
        ttk.Button(row, text="\u21ba Kiri (L)", command=lambda: self.rotate(-90)).pack(side="right", padx=(0, 4))

        self.bind("<Left>", lambda e: self.go(-1))
        self.bind("<Right>", lambda e: self.go(1))
        self.bind("<KeyPress-l>", lambda e: self.rotate(-90))
        self.bind("<KeyPress-r>", lambda e: self.rotate(90))
        self.bind("<Delete>", lambda e: self.delete())
        self.bind("<Escape>", lambda e: self.destroy())
        self.canvas.bind("<Configure>", lambda e: self.schedule())
        self.focus_set()

    def schedule(self):
        if self._pending:
            self.after_cancel(self._pending)
        self._pending = self.after(60, self.render)

    def render(self):
        self._pending = None
        pages = self.app.pages
        if not pages:
            self.destroy()
            return
        self.index = max(0, min(self.index, len(pages) - 1))
        page = pages[self.index]
        cw, ch = max(self.canvas.winfo_width(), 50), max(self.canvas.winfo_height(), 50)
        try:
            with Image.open(page["temp"]) as im:
                im.draft("RGB", (max(cw, ch), max(cw, ch)))
                im = rotated(im.convert("RGB"), page["angle"])
                im.thumbnail((cw - 24, ch - 24), Image.Resampling.LANCZOS)
                self.photo = ImageTk.PhotoImage(im)
        except OSError:
            self.photo = None
        self.canvas.delete("all")
        if self.photo:
            self.canvas.create_image(cw // 2, ch // 2, image=self.photo)
        else:
            self.canvas.create_text(cw // 2, ch // 2, fill="white", text="Gambar halaman ini tidak bisa dibuka.")
        self.var_info.set(f"Halaman {self.index + 1} dari {len(pages)}   {self.app.full_detail(page)}")
        self.title(f"Periksa halaman {self.index + 1} dari {len(pages)}")

    def go(self, step):
        self.index = (self.index + step) % max(1, len(self.app.pages))
        self.render()

    def rotate(self, delta):
        if self.app.pages and self.app.rotate(self.app.pages[self.index], delta):
            self.render()

    def delete(self):
        if self.app.pages and self.app.delete(self.app.pages[self.index]):
            self.render()

    def destroy(self):
        self.app.viewer = None
        super().destroy()
        self.app.root.focus_set()


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_settings()
        self.devices = []
        self.pages = []
        self.queue = queue.Queue()
        self.stop_event = threading.Event()
        self.busy = None
        self.viewer = None
        self.close_after_save = False
        self.scan_note = ""
        self.stats = {"pages": 0, "blank": 0, "rotated": 0, "skewed": 0, "gray": 0}

        shutil.rmtree(WORK_DIR, ignore_errors=True)

        root.title(APP_NAME)
        if ICON_FILE.exists():
            try:
                root.iconbitmap(default=str(ICON_FILE))
            except tk.TclError:
                pass
        root.geometry("1000x690")
        root.minsize(780, 500)
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        root.option_add("*Font", ("Segoe UI", 10))
        style.configure("Muted.TLabel", foreground=MUTED)
        style.configure("Head.TLabel", font=("Segoe UI Semibold", 11))
        style.configure("Tool.TButton", padding=(2, 1))

        self.var_device = tk.StringVar()
        self.var_side = tk.StringVar(value=self.cfg["side"])
        self.var_color = tk.StringVar(value=self.cfg["color"])
        self.var_dpi = tk.StringVar(value=str(self.cfg["dpi"]))
        self.var_paper = tk.StringVar(value=self.cfg["paper"])
        self.var_quality = tk.IntVar(value=int(self.cfg["quality"]))
        self.var_skip = tk.BooleanVar(value=bool(self.cfg["skip_blank"]))
        self.var_rotate = tk.BooleanVar(value=bool(self.cfg.get("auto_rotate", True)))
        self.var_deskew = tk.BooleanVar(value=bool(self.cfg.get("deskew", True)))
        self.var_out_size = tk.StringVar(value=self.cfg["out_size"])
        self.var_brightness = tk.IntVar(value=int(self.cfg.get("brightness", 0)))
        self.var_contrast = tk.IntVar(value=int(self.cfg.get("contrast", 0)))
        self.var_folder = tk.StringVar(value=self.cfg["folder"])
        self.var_name = tk.StringVar()
        self.var_status = tk.StringVar()
        self.var_count = tk.StringVar()

        self.build()
        self.bind_keys()
        self.refresh_devices()
        self.update_view()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    def build(self):
        status = ttk.Frame(self.root, padding=(16, 6, 16, 8))
        status.pack(side="bottom", fill="x")
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=120)
        ttk.Label(status, textvariable=self.var_status, style="Muted.TLabel").pack(side="left")

        outer = ttk.Frame(self.root, padding=(16, 12, 16, 0))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(0, weight=1)

        side = ttk.Frame(outer, width=290)
        side.grid(row=0, column=0, sticky="ns", padx=(0, 18))
        side.columnconfigure(0, weight=1)

        r = 0
        ttk.Label(side, text="Scanner", style="Head.TLabel").grid(row=r, column=0, sticky="w"); r += 1
        dev_row = ttk.Frame(side)
        dev_row.grid(row=r, column=0, sticky="ew", pady=(4, 8)); r += 1
        dev_row.columnconfigure(0, weight=1)
        self.cb_device = ttk.Combobox(dev_row, textvariable=self.var_device, state="readonly", width=26)
        self.cb_device.grid(row=0, column=0, sticky="ew")
        ttk.Button(dev_row, text="Muat ulang", command=self.refresh_devices).grid(row=0, column=1, padx=(6, 0))

        tabs = ttk.Notebook(side)
        tabs.grid(row=r, column=0, sticky="ew", pady=(2, 10)); r += 1
        scan_tab = ttk.Frame(tabs, padding=(10, 8))
        out_tab = ttk.Frame(tabs, padding=(10, 8))
        tabs.add(scan_tab, text="Scan")
        tabs.add(out_tab, text="Hasil")
        for tab in (scan_tab, out_tab):
            tab.columnconfigure(1, weight=1)

        def field(tab, row, label, widget):
            ttk.Label(tab, text=label).grid(row=row, column=0, sticky="w", pady=2, padx=(0, 10))
            widget.grid(row=row, column=1, sticky="ew", pady=2)

        def combo(tab, var, values, width=20):
            return ttk.Combobox(tab, textvariable=var, values=values, state="readonly", width=width)

        def slider(tab, var, lo, hi):
            row = ttk.Frame(tab)
            row.columnconfigure(0, weight=1)
            ttk.Scale(row, from_=lo, to=hi, variable=var,
                      command=lambda v: var.set(round(float(v)))).grid(row=0, column=0, sticky="ew")
            ttk.Label(row, textvariable=var, width=4, anchor="e").grid(row=0, column=1, padx=(6, 0))
            return row

        sides = ttk.Frame(scan_tab)
        for i, name in enumerate(SIDES):
            ttk.Radiobutton(sides, text=name, value=name, variable=self.var_side).grid(row=0, column=i, padx=(0, 10))
        field(scan_tab, 0, "Sisi", sides)
        field(scan_tab, 1, "Mode warna", combo(scan_tab, self.var_color, list(COLOR_MODES)))
        field(scan_tab, 2, "Resolusi (DPI)", combo(scan_tab, self.var_dpi, RESOLUTIONS))
        field(scan_tab, 3, "Ukuran input", combo(scan_tab, self.var_paper, list(PAPER_SIZES_IN)))
        ttk.Label(scan_tab, text="Mode otomatis memindai berwarna, lalu menyimpan abu-abu bila halaman "
                                 "tidak punya warna (cap dan tanda tangan berwarna tetap terjaga).",
                  style="Muted.TLabel", wraplength=255).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        field(out_tab, 0, "Ukuran hasil", combo(out_tab, self.var_out_size, list(OUTPUT_SIZES)))
        field(out_tab, 1, "Kualitas JPG", slider(out_tab, self.var_quality, 50, 95))
        field(out_tab, 2, "Kecerahan", slider(out_tab, self.var_brightness, -50, 50))
        field(out_tab, 3, "Kontras", slider(out_tab, self.var_contrast, -50, 50))
        exp_row = ttk.Frame(out_tab)
        exp_row.grid(row=4, column=0, columnspan=2, sticky="ew")
        ttk.Label(exp_row, text="Pencahayaan untuk warna dan abu-abu.", style="Muted.TLabel").pack(side="left")
        ttk.Button(exp_row, text="Bawaan", style="Tool.TButton", command=self.reset_exposure).pack(side="right")
        checks = (("Orientasi otomatis (putar sesuai arah teks)", self.var_rotate),
                  ("Luruskan halaman miring", self.var_deskew),
                  ("Buang halaman kosong", self.var_skip))
        for i, (text, var) in enumerate(checks):
            ttk.Checkbutton(out_tab, text=text, variable=var).grid(
                row=5 + i, column=0, columnspan=2, sticky="w", pady=(6 if i == 0 else 0, 0))

        ttk.Label(side, text="Simpan ke", style="Head.TLabel").grid(row=r, column=0, sticky="w"); r += 1
        f_row = ttk.Frame(side)
        f_row.grid(row=r, column=0, sticky="ew", pady=(4, 6)); r += 1
        f_row.columnconfigure(0, weight=1)
        ttk.Entry(f_row, textvariable=self.var_folder, state="readonly").grid(row=0, column=0, sticky="ew")
        ttk.Button(f_row, text="Ganti...", command=self.pick_folder).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(side, text="Nama berkas (boleh kosong)").grid(row=r, column=0, sticky="w"); r += 1
        ttk.Entry(side, textvariable=self.var_name).grid(row=r, column=0, sticky="ew", pady=(2, 0)); r += 1
        ttk.Label(side, text="Kosong = SCAN_tanggal_jam. Gambar diberi nomor _001, _002, dst; PDF memakai nama ini.",
                  style="Muted.TLabel", wraplength=270).grid(row=r, column=0, sticky="w", pady=(2, 10)); r += 1

        common = dict(relief="flat", font=("Segoe UI Semibold", 11), padx=10, pady=5, cursor="hand2",
                      highlightthickness=2, highlightcolor="#0B2E57")
        self.btn_scan = tk.Button(side, text="Scan  (F5)", command=self.start_scan, **common)
        self.btn_scan.grid(row=r, column=0, sticky="ew"); r += 1
        save_row = tk.Frame(side)
        save_row.grid(row=r, column=0, sticky="ew", pady=(8, 0)); r += 1
        save_row.columnconfigure((0, 1), weight=1, uniform="save")
        self.btn_save = tk.Button(save_row, text="Simpan gambar", command=lambda: self.start_save("jpg"), **common)
        self.btn_save.grid(row=0, column=0, sticky="ew", padx=(0, 3))
        self.btn_pdf = tk.Button(save_row, text="Simpan PDF", command=lambda: self.start_save("pdf"), **common)
        self.btn_pdf.grid(row=0, column=1, sticky="ew", padx=(3, 0))
        ttk.Label(side, text="Gambar: JPG per halaman (Ctrl+S).  PDF: satu berkas (Ctrl+Shift+S).",
                  style="Muted.TLabel", wraplength=280).grid(row=r, column=0, sticky="w", pady=(3, 0)); r += 1
        actions = ttk.Frame(side)
        actions.grid(row=r, column=0, sticky="ew", pady=(8, 0)); r += 1
        actions.columnconfigure((0, 1), weight=1)
        self.btn_stop = ttk.Button(actions, text="Berhenti (Esc)", command=self.stop_scan)
        self.btn_stop.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(actions, text="Buka folder", command=self.open_folder).grid(row=0, column=1, sticky="ew", padx=(4, 0))

        main = ttk.Frame(outer)
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        head = ttk.Frame(main)
        head.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        head.columnconfigure(0, weight=1)
        ttk.Label(head, textvariable=self.var_count, style="Head.TLabel").grid(row=0, column=0, sticky="w")
        self.btn_discard = ttk.Button(head, text="Buang semua", command=self.discard_all)
        self.btn_discard.grid(row=0, column=1)

        box = tk.Frame(main, bg=SURFACE, highlightthickness=1, highlightbackground="#D5D8DC")
        box.grid(row=1, column=0, sticky="nsew")
        box.columnconfigure(0, weight=1)
        box.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(box, bg=SURFACE, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.canvas.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=sb.set)
        self.grid_frame = tk.Frame(self.canvas, bg=SURFACE)
        self.grid_win = self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.grid_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self.on_canvas_resize)
        self.canvas.bind("<Enter>", lambda e: self.root.bind_all("<MouseWheel>", self.on_wheel))
        self.canvas.bind("<Leave>", lambda e: self.root.unbind_all("<MouseWheel>"))

        self.empty = tk.Label(box, bg=SURFACE, fg=MUTED, justify="center",
                              text="Belum ada halaman.\nTaruh dokumen di feeder, menghadap ke atas, lalu tekan Scan.\n"
                                   "Hasil scan tampil di sini untuk diperiksa sebelum disimpan.")


    def reset_exposure(self):
        self.var_brightness.set(0)
        self.var_contrast.set(0)

    def bind_keys(self):
        self.root.bind("<F5>", lambda e: self.start_scan())
        self.root.bind("<Escape>", lambda e: self.stop_scan())
        self.root.bind("<Control-s>", lambda e: self.start_save("jpg"))
        self.root.bind("<Control-S>", lambda e: self.start_save("pdf"))
        self.root.bind("<Control-o>", lambda e: self.open_folder())

    def set_status(self, text):
        self.var_status.set(text)

    def update_view(self):
        has = bool(self.pages)
        idle = self.busy is None
        can_scan = idle and bool(self.devices)
        self.btn_scan.configure(state="normal" if can_scan else "disabled",
                                **(SECONDARY_BTN if has else PRIMARY_BTN))
        for btn in (self.btn_save, self.btn_pdf):
            btn.configure(state="normal" if has and idle else "disabled", **(PRIMARY_BTN if has else SECONDARY_BTN))
        self.btn_discard.configure(state="normal" if has and idle else "disabled")
        self.btn_stop.configure(state="normal" if self.busy == "scan" and not self.stop_event.is_set() else "disabled")
        n = len(self.pages)
        self.var_count.set(f"Belum disimpan: {n} halaman" if n else "Hasil scan")
        if has:
            self.empty.place_forget()
        else:
            self.empty.place(relx=0.5, rely=0.45, anchor="center")

    def refresh_devices(self):
        if self.busy:
            return
        try:
            self.devices = list_scanners(self.root)
        except Exception:
            self.devices = []
        names = [n for n, _ in self.devices]
        self.cb_device["values"] = names
        if not self.devices:
            self.var_device.set("")
            self.set_status("Scanner tidak ditemukan. Pastikan scanner menyala, kabel terpasang, "
                            "dan driver HP Scanjet sudah diinstal, lalu tekan Muat ulang.")
        else:
            saved = next((n for n, d in self.devices if d == self.cfg.get("device_id")), names[0])
            self.var_device.set(saved)
            self.set_status(f"Siap. {len(self.devices)} scanner terdeteksi.")
        self.update_view()

    def pick_folder(self):
        path = filedialog.askdirectory(initialdir=self.var_folder.get() or str(Path.home()),
                                       title="Pilih folder hasil scan")
        if path:
            self.var_folder.set(os.path.normpath(path))

    def open_folder(self):
        folder = Path(self.var_folder.get())
        if not folder.exists():
            self.set_status("Folder belum ada. Folder dibuat otomatis saat halaman pertama disimpan.")
            return
        os.startfile(folder)

    def current_cfg(self):
        dev_id = next((d for n, d in self.devices if n == self.var_device.get()), self.cfg.get("device_id", ""))
        return {
            "device_id": dev_id,
            "side": self.var_side.get(),
            "color": self.var_color.get(),
            "dpi": self.var_dpi.get(),
            "paper": self.var_paper.get(),
            "quality": int(self.var_quality.get()),
            "skip_blank": bool(self.var_skip.get()),
            "auto_rotate": bool(self.var_rotate.get()),
            "deskew": bool(self.var_deskew.get()),
            "out_size": self.var_out_size.get(),
            "brightness": int(self.var_brightness.get()),
            "contrast": int(self.var_contrast.get()),
            "folder": self.var_folder.get(),
            "save_format": self.cfg.get("save_format", "jpg"),
        }

    def start_busy(self, kind, text):
        self.busy = kind
        self.progress.pack(side="right")
        self.progress.start(12)
        self.set_status(text)
        self.update_view()
        self.root.after(100, self.poll)

    def end_busy(self):
        self.busy = None
        self.progress.stop()
        self.progress.pack_forget()
        self.update_view()

    def start_scan(self):
        if self.busy or not self.devices:
            return
        cfg = self.current_cfg()
        if not cfg["device_id"]:
            self.set_status("Pilih scanner dulu.")
            return
        self.cfg = cfg
        save_settings(cfg)
        opts = {
            "paper_name": cfg["paper"],
            "handling": SIDES[cfg["side"]],
            "intent": COLOR_MODES[cfg["color"]],
            "dpi": int(cfg["dpi"]),
            "paper": PAPER_SIZES_IN[cfg["paper"]],
            "skip_blank": cfg["skip_blank"],
            "auto_rotate": cfg["auto_rotate"],
            "deskew": cfg["deskew"],
            "auto_color": cfg["color"] == AUTO_COLOR,
            "bw": cfg["color"] == BW,
            "brightness": cfg["brightness"],
            "contrast": cfg["contrast"],
            "token": datetime.now().strftime("%H%M%S%f"),
        }
        self.stop_event.clear()
        self.stats = {"pages": 0, "blank": 0, "rotated": 0, "skewed": 0, "gray": 0}
        self.scan_note = ""
        kind, _, dev = cfg["device_id"].partition(":")
        worker = twain_worker if kind == "twain" else scan_worker
        threading.Thread(target=worker, args=(dev, opts, self.queue, self.stop_event), daemon=True).start()
        self.btn_scan.configure(text="Memindai...")
        self.start_busy("scan", "Menarik halaman pertama dari feeder...")

    def stop_scan(self):
        if self.busy == "scan" and not self.stop_event.is_set():
            self.stop_event.set()
            self.update_view()
            self.set_status("Berhenti setelah lembar yang sedang ditarik selesai...")

    def poll(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                kind = msg[0]
                if kind == "page":
                    self.add_page(msg[1])
                elif kind == "note":
                    self.scan_note = msg[1]
                    self.set_status(msg[1])
                elif kind == "blank":
                    self.stats["blank"] += 1
                    self.set_status(f"Halaman kosong dibuang ({self.stats['blank']}).")
                elif kind == "error":
                    messagebox.showerror(APP_NAME, msg[1], parent=self.root)
                elif kind == "saved":
                    self.set_status(f"Menyimpan {msg[1]} dari {self.saving_total} halaman...")
                elif kind == "done":
                    self.finish_scan()
                    return
                elif kind == "save_done":
                    self.finish_save(*msg[1:])
                    return
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def caption_detail(self, page):
        parts = [page["size"]] if page["size"] else []
        if page["angle"]:
            parts.append(f"diputar {page['angle']}\u00b0" + (" otomatis" if page["angle"] == page["auto"] else ""))
        return " \u00b7 ".join(parts)

    def full_detail(self, page):
        parts = [self.caption_detail(page)] if self.caption_detail(page) else []
        if page.get("tone"):
            parts.append(page["tone"])
        if page.get("engine") == "hp":
            parts.append("diolah driver HP")
        if page.get("skew"):
            parts.append(f"diluruskan {abs(page['skew']):.1f}\u00b0".replace(".", ","))
        return " \u00b7 ".join(parts)

    def add_page(self, page):
        self.stats["pages"] += 1
        if page["angle"]:
            self.stats["rotated"] += 1
        if page.get("skew"):
            self.stats["skewed"] += 1
        if page.get("tone") == "abu-abu":
            self.stats["gray"] += 1
        self.pages.append(page)
        self.make_tile(page)
        self.relayout()
        self.update_view()
        self.set_status(f"Halaman {len(self.pages)} masuk." + (f" Diputar otomatis {page['angle']}\u00b0." if page["angle"] else ""))
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(1.0)

    def thumb_photo(self, page):
        if page["thumb"] is None:
            return None
        return ImageTk.PhotoImage(rotated(page["thumb"], page["angle"]))

    def make_tile(self, page):
        tile = tk.Frame(self.grid_frame, bg=SURFACE)
        page["photo"] = self.thumb_photo(page)
        btn = ttk.Button(tile, image=page["photo"] or "", text="" if page["photo"] else "(tanpa pratinjau)",
                         command=lambda p=page: self.open_viewer(p))
        btn.pack()
        btn.bind("<Return>", lambda e, p=page: self.open_viewer(p))
        btn.bind("<Delete>", lambda e, p=page: self.delete(p))
        btn.bind("<KeyPress-r>", lambda e, p=page: self.rotate(p, 90))
        btn.bind("<KeyPress-l>", lambda e, p=page: self.rotate(p, -90))
        caption = tk.Label(tile, bg=SURFACE, fg=MUTED, font=("Segoe UI", 9), wraplength=THUMB_W + 30, justify="center", height=3, anchor="n")
        caption.pack(pady=(3, 2))
        tools = tk.Frame(tile, bg=SURFACE)
        tools.pack()
        ttk.Button(tools, text="\u21ba", width=3, style="Tool.TButton",
                   command=lambda p=page: self.rotate(p, -90)).pack(side="left")
        ttk.Button(tools, text="\u21bb", width=3, style="Tool.TButton",
                   command=lambda p=page: self.rotate(p, 90)).pack(side="left", padx=2)
        ttk.Button(tools, text="Hapus", width=6, style="Tool.TButton",
                   command=lambda p=page: self.delete(p)).pack(side="left")
        page.update(tile=tile, btn=btn, caption=caption)
        self.refresh_caption(page)

    def refresh_caption(self, page):
        idx = self.pages.index(page) + 1
        detail = self.caption_detail(page)
        page["caption"].configure(text=f"Hal. {idx}" + (f"\n{detail}" if detail else ""))

    def rotate(self, page, delta):
        if self.busy == "save":
            return False
        page["angle"] = (page["angle"] + delta) % 360
        page["photo"] = self.thumb_photo(page)
        if page["photo"]:
            page["btn"].configure(image=page["photo"])
        self.refresh_caption(page)
        return True

    def delete(self, page):
        if self.busy == "save" or page not in self.pages:
            return False
        idx = self.pages.index(page)
        self.pages.remove(page)
        page["tile"].destroy()
        page["temp"].unlink(missing_ok=True)
        for p in self.pages:
            self.refresh_caption(p)
        self.relayout()
        self.update_view()
        self.set_status(f"Halaman {idx + 1} dihapus.")
        if self.pages:
            target = self.pages[min(idx, len(self.pages) - 1)]["btn"]
            if self.viewer is None:
                target.focus_set()
        return True

    def discard_all(self, ask=True):
        if self.busy or not self.pages:
            return
        if ask and not messagebox.askyesno(APP_NAME, f"Buang {len(self.pages)} halaman yang belum disimpan?",
                                           parent=self.root):
            return
        if self.viewer:
            self.viewer.destroy()
        for p in self.pages:
            p["tile"].destroy()
            p["temp"].unlink(missing_ok=True)
        self.pages.clear()
        self.update_view()
        self.set_status("Semua halaman dibuang. Tidak ada berkas yang disimpan.")

    def open_viewer(self, page):
        if page not in self.pages:
            return
        if self.viewer is None:
            self.viewer = Viewer(self, self.pages.index(page))
        else:
            self.viewer.index = self.pages.index(page)
            self.viewer.lift()
            self.viewer.focus_set()
        self.viewer.schedule()

    def relayout(self):
        width = max(self.canvas.winfo_width(), 1)
        cols = max(1, (width - 12) // (THUMB_W + 40))
        for i, page in enumerate(self.pages):
            page["tile"].grid(row=i // cols, column=i % cols, padx=14, pady=12, sticky="n")

    def on_canvas_resize(self, event):
        self.canvas.itemconfigure(self.grid_win, width=event.width)
        self.relayout()

    def on_wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def finish_scan(self):
        self.btn_scan.configure(text="Scan  (F5)")
        self.end_busy()
        s = self.stats
        if s["pages"]:
            extra = []
            if s["rotated"]:
                extra.append(f"{s['rotated']} diputar otomatis")
            if s["skewed"]:
                extra.append(f"{s['skewed']} diluruskan")
            if s["gray"]:
                extra.append(f"{s['gray']} disimpan abu-abu")
            if s["blank"]:
                extra.append(f"{s['blank']} halaman kosong dibuang")
            note = f" ({', '.join(extra)})" if extra else ""
            if any(p.get("engine") == "hp" for p in self.pages[-s["pages"]:]):
                note += " lewat driver HP"
            self.set_status(f"{s['pages']} halaman dipindai{note}. Periksa, lalu simpan sebagai gambar atau PDF. "
                            + self.scan_note)
        elif s["blank"]:
            self.set_status(f"Semua {s['blank']} halaman terdeteksi kosong dan dibuang. "
                            "Matikan opsi Buang halaman kosong bila ini keliru.")
        else:
            self.set_status("Tidak ada halaman baru.")

    def start_save(self, fmt=None):
        if self.busy or not self.pages:
            return
        cfg = self.current_cfg()
        fmt = fmt or self.cfg.get("save_format", "jpg")
        cfg["save_format"] = fmt
        self.cfg = cfg
        save_settings(cfg)
        base = clean_name(self.var_name.get()) or datetime.now().strftime("SCAN_%d%m%y_%H%M%S")
        self.save_base = base
        self.saving_total = len(self.pages)
        snapshot = [dict(temp=p["temp"], angle=p["angle"], dpi=p["dpi"]) for p in self.pages]
        threading.Thread(target=save_worker, args=(snapshot, Path(cfg["folder"]), base, cfg["quality"],
                                                    OUTPUT_SIZES[cfg["out_size"]], fmt, self.queue),
                         daemon=True).start()
        self.start_busy("save", f"Menyimpan 0 dari {self.saving_total} halaman...")

    def finish_save(self, saved, error, name):
        done, self.pages = self.pages[:saved], self.pages[saved:]
        if self.viewer and done:
            self.viewer.destroy()
        for p in done:
            p["tile"].destroy()
            p["temp"].unlink(missing_ok=True)
        for p in self.pages:
            self.refresh_caption(p)
        self.relayout()
        self.end_busy()
        if error:
            self.close_after_save = False
            messagebox.showerror(APP_NAME, error, parent=self.root)
            self.set_status(f"{saved} halaman tersimpan, {len(self.pages)} belum. {error}")
            return
        self.var_name.set("")
        self.set_status(f"{saved} halaman tersimpan di {self.cfg['folder']} sebagai {name}.")
        if self.close_after_save:
            self.on_close()

    def on_close(self):
        if self.busy == "scan":
            if not messagebox.askyesno(APP_NAME, "Scan masih berjalan. Tutup aplikasi?", parent=self.root):
                return
            self.stop_event.set()
        elif self.busy == "save":
            self.close_after_save = True
            self.set_status("Aplikasi ditutup setelah penyimpanan selesai...")
            return
        elif self.pages:
            label = "PDF" if self.cfg.get("save_format") == "pdf" else "gambar JPG"
            answer = messagebox.askyesnocancel(
                APP_NAME, f"Ada {len(self.pages)} halaman belum disimpan. Simpan sebagai {label} sebelum menutup?",
                parent=self.root)
            if answer is None:
                return
            if answer:
                self.close_after_save = True
                self.start_save()
                return
        save_settings(self.current_cfg())
        shutil.rmtree(WORK_DIR, ignore_errors=True)
        self.root.destroy()


def create_shortcuts():
    """Buat shortcut Pindai di Desktop dan Start Menu dengan identitas taskbar yang sama dengan aplikasinya."""
    from win32com.shell import shell, shellcon
    from win32com.propsys import propsys, pscon

    script = Path(__file__).resolve()
    exe = Path(sys.executable)
    pythonw = exe.with_name("pythonw.exe")
    target = pythonw if pythonw.exists() else exe
    made = []
    for csidl in (shellcon.CSIDL_DESKTOPDIRECTORY, shellcon.CSIDL_PROGRAMS):
        folder = Path(shell.SHGetFolderPath(0, csidl, None, 0))
        link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER,
                                          shell.IID_IShellLink)
        link.SetPath(str(target))
        link.SetArguments(f'"{script}"')
        link.SetWorkingDirectory(str(script.parent))
        link.SetDescription("Scan dokumen dari HP Scanjet ke JPG")
        if ICON_FILE.exists():
            link.SetIconLocation(str(ICON_FILE), 0)
        store = link.QueryInterface(propsys.IID_IPropertyStore)
        store.SetValue(pscon.PKEY_AppUserModel_ID, propsys.PROPVARIANTType(APP_ID, pythoncom.VT_LPWSTR))
        store.Commit()
        path = folder / f"{APP_NAME}.lnk"
        link.QueryInterface(pythoncom.IID_IPersistFile).Save(str(path), 0)
        made.append(path)
    return made


def main():
    if "--shortcut" in sys.argv:
        root = tk.Tk()
        root.withdraw()
        try:
            made = create_shortcuts()
        except Exception as err:
            messagebox.showerror(APP_NAME, f"Shortcut gagal dibuat: {err}")
        else:
            messagebox.showinfo(APP_NAME, "Shortcut Pindai sudah dibuat di:\n" + "\n".join(str(p) for p in made) +
                                "\n\nUntuk menyematkan ke taskbar: buka Start, ketik Pindai, klik kanan, "
                                "lalu pilih Pin to taskbar.")
        root.destroy()
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except (AttributeError, OSError):
        pass
    root = tk.Tk()
    missing = [name for name, ok in (("pywin32", HAS_WIA), ("pillow", HAS_PIL)) if not ok]
    # pytwain tidak wajib: tanpa itu Pindai memakai WIA dan pengolahan buatannya sendiri.
    if missing:
        root.withdraw()
        messagebox.showerror(APP_NAME, "Modul berikut belum terpasang: " + ", ".join(missing) + ".\n\n"
                             "Buka Command Prompt lalu jalankan:\npip install " + " ".join(missing) + "\n\n"
                             "Atau buka lewat Jalankan Pindai.bat yang memasangnya otomatis.")
        root.destroy()
        return
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
