#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import subprocess
import tempfile

from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt
from lxml import etree

NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _ensure_effect_lst(spPr) -> etree._Element:
    effectLst = spPr.find("a:effectLst", NSMAP)
    if effectLst is None:
        effectLst = etree.SubElement(spPr, "{%s}effectLst" % NSMAP["a"])
    return effectLst


def add_glow_to_shape(shape, glow_color: RGBColor, size: int = 40000) -> None:
    try:
        spPr = shape._element.find(".//a:spPr", NSMAP)
        if spPr is None:
            return
        effectLst = _ensure_effect_lst(spPr)
        glow = etree.SubElement(effectLst, "{%s}glow" % NSMAP["a"])
        glow.set("rad", str(size))
        srgb = etree.SubElement(glow, "{%s}srgbClr" % NSMAP["a"])
        srgb.set("val", "%02X%02X%02X" % (glow_color[0], glow_color[1], glow_color[2]))
        alpha = etree.SubElement(srgb, "{%s}alpha" % NSMAP["a"])
        alpha.set("val", "35000")
    except Exception:
        pass


def add_glow_to_run(run, glow_color: RGBColor, size: int = 50000) -> None:
    try:
        rPr = run._r.find(".//a:rPr", NSMAP)
        if rPr is None:
            return
        effectLst = _ensure_effect_lst(rPr)
        glow = etree.SubElement(effectLst, "{%s}glow" % NSMAP["a"])
        glow.set("rad", str(size))
        srgb = etree.SubElement(glow, "{%s}srgbClr" % NSMAP["a"])
        srgb.set("val", "%02X%02X%02X" % (glow_color[0], glow_color[1], glow_color[2]))
        alpha = etree.SubElement(srgb, "{%s}alpha" % NSMAP["a"])
        alpha.set("val", "40000")
    except Exception:
        pass


def add_outer_shadow(shape, color_rgb: str = "000000",
                     blur_rad: int = 76200, dist: int = 25400,
                     direction: int = 5400000, alpha_pct: int = 40000):
    """Add an outer drop shadow to a shape via OOXML injection."""
    try:
        spPr = shape._element.find(".//a:spPr", NSMAP)
        if spPr is None:
            return
        effectLst = _ensure_effect_lst(spPr)
        outerShdw = etree.SubElement(effectLst, "{%s}outerShdw" % NSMAP["a"])
        outerShdw.set("blurRad", str(blur_rad))
        outerShdw.set("dist", str(dist))
        outerShdw.set("dir", str(direction))
        outerShdw.set("algn", "bl")
        outerShdw.set("rotWithShape", "0")
        srgbClr = etree.SubElement(outerShdw, "{%s}srgbClr" % NSMAP["a"])
        srgbClr.set("val", color_rgb)
        alpha = etree.SubElement(srgbClr, "{%s}alpha" % NSMAP["a"])
        alpha.set("val", str(alpha_pct))
    except Exception:
        pass


def add_accent_line(slide, left_px, top_px, width_px, color_name, thickness=3):
    """Add a thin horizontal accent/separator line with glow."""
    accent = color(color_name)
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(left_px), px(top_px), px(width_px), px(thickness))
    line.fill.solid()
    line.fill.fore_color.rgb = accent
    line.line.fill.background()
    add_glow_to_shape(line, accent, size=18000)


def add_gradient_panel(slide, left_px, top_px, width_px, height_px, accent_name, transparency=0.30):
    """Add a rounded rectangle card with gradient fill and glow border."""
    accent = color(accent_name)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left_px), px(top_px), px(width_px), px(height_px))
    fill = shape.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = RGBColor(10, 10, 18)
    fill.gradient_stops[0].position = 0.0
    fill.gradient_stops[1].color.rgb = RGBColor(18, 18, 32)
    fill.gradient_stops[1].position = 1.0
    shape.fill.transparency = transparency
    shape.line.color.rgb = RGBColor(255, 255, 255)
    shape.line.width = Pt(1.2)
    add_glow_to_shape(shape, accent, size=40000)
    add_outer_shadow(shape, color_rgb="%02X%02X%02X" % (accent[0], accent[1], accent[2]),
                     blur_rad=50000, dist=12700, direction=5400000, alpha_pct=25000)
    return shape


CANVAS_PRESETS = {
    "widescreen": {
        "width": 1920,
        "height": 1080,
        "slide_w": Inches(13.333333),
        "slide_h": Inches(7.5),
    },
    "xhs-vertical": {
        "width": 1080,
        "height": 1440,
        "slide_w": Inches(7.5),
        "slide_h": Inches(10),
    },
    "lecture-vertical": {
        "width": 1080,
        "height": 1920,
        "slide_w": Inches(7.5),
        "slide_h": Inches(13.333333),
    },
}

FONT_PATH_BLACK = "/usr/share/fonts/opentype/noto/NotoSansCJK-Black.ttc"
FONT_PATH_REGULAR = "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
FONT_PATH_MONO = "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"

COLORS = {
    "WHITE": RGBColor(255, 255, 255),
    "MUTED": RGBColor(188, 194, 210),
    "SOFT": RGBColor(120, 132, 154),
    "CARD": RGBColor(10, 10, 10),
    "CARD_2": RGBColor(5, 5, 8),
    "CYAN": RGBColor(0, 255, 255),
    "BLUE": RGBColor(59, 130, 246),
    "ORANGE": RGBColor(249, 115, 22),
    "YELLOW": RGBColor(251, 191, 36),
    "PINK": RGBColor(236, 72, 153),
    "RED": RGBColor(255, 51, 102),
    "PURPLE": RGBColor(139, 92, 246),
    "LIME": RGBColor(16, 185, 129),
    "TEAL": RGBColor(20, 184, 166),
}

# Safe area margins (px) — tag at top, page number at bottom
SLIDE_SAFE = {
    "widescreen": {"max_y": 980, "max_x": 1860, "top_y": 110},
    "xhs-vertical": {"max_y": 1380, "max_x": 1020, "top_y": 110},
    "lecture-vertical": {"max_y": 1860, "max_x": 1020, "top_y": 150},
}


def px(value: float):
    return Inches(value / 144)


def color(name: str) -> RGBColor:
    try:
        return COLORS[name.upper()]
    except KeyError as exc:
        raise ValueError(f"Unsupported color: {name}") from exc


def to_rgb(color_value: RGBColor) -> tuple[int, int, int]:
    return (color_value[0], color_value[1], color_value[2])


def pil_font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_PATH_BLACK, size)


def _clamp(val: int, lo: int, hi: int) -> int:
    return max(lo, min(val, hi))


def _box_height_for_pt(pt_size: int) -> int:
    """Minimum textbox height (px) to fit a given pt font size without overflow."""
    return max(50, int(pt_size * 2.4))


def _line_advance_for_pt(pt_size: int) -> int:
    """Y advance (px) after rendering one title line at the given pt size."""
    return _box_height_for_pt(pt_size) + 8


def _resolve_font_path(font_name: str) -> str:
    if "Mono" in font_name or "mono" in font_name:
        return FONT_PATH_MONO
    return FONT_PATH_BLACK


def measure_text(text: str, font_path: str, font_size_pt: int, max_width_px: int) -> dict:
    """Measure text with Pillow getbbox() and simulate word wrapping.

    Returns dict with lines, num_lines, total_height_px, max_width_px.
    Handles CJK characters (each char is a word boundary) and Latin words.
    """
    font = ImageFont.truetype(font_path, font_size_pt)

    # Split into tokens: CJK chars as individual tokens, Latin words grouped
    tokens: list[str] = []
    for ch in text:
        # CJK Unified Ideographs, Hiragana, Katakana, etc.
        if '一' <= ch <= '鿿' or '぀' <= ch <= 'ヿ' or '가' <= ch <= '힯':
            tokens.append(ch)
        elif ch in (' ', '\t'):
            if tokens and tokens[-1] != ' ':
                tokens.append(' ')
        else:
            if tokens and tokens[-1] not in (' ', '') and '一' <= tokens[-1][-1] <= '鿿':
                tokens.append(ch)
            elif tokens and tokens[-1] not in (' ', ''):
                tokens[-1] += ch
            else:
                tokens.append(ch)

    lines: list[str] = []
    current_line = ""

    for token in tokens:
        if token == ' ':
            current_line += ' '
            continue
        test_line = current_line + token
        bbox = font.getbbox(test_line)
        text_width = bbox[2] - bbox[0]
        if text_width > max_width_px and current_line.strip():
            lines.append(current_line.strip())
            current_line = token
        else:
            current_line = test_line

    if current_line.strip():
        lines.append(current_line.strip())

    ascent, descent = font.getmetrics()
    line_height = int((ascent + descent) * 1.15)
    total_height = len(lines) * line_height
    max_line_width = max(
        font.getbbox(line)[2] - font.getbbox(line)[0]
        for line in lines
    ) if lines else 0

    return {
        "lines": lines,
        "num_lines": len(lines),
        "total_height_px": total_height,
        "max_width_px": max_line_width,
        "line_height_px": line_height,
    }


def fit_text_to_box(
    text: str,
    font_path: str,
    max_width_px: int,
    max_height_px: int,
    max_pt: int = 18,
    min_pt: int = 8,
) -> int:
    """Return the largest font size (pt) that fits text in the box."""
    for pt in range(max_pt, min_pt - 1, -1):
        metrics = measure_text(text, font_path, pt, max_width_px)
        if metrics["total_height_px"] <= max_height_px:
            return pt
    return min_pt


def get_canvas(spec: dict) -> dict:
    canvas_name = spec.get("canvas", "widescreen")
    try:
        return CANVAS_PRESETS[canvas_name]
    except KeyError as exc:
        raise ValueError(f"Unsupported canvas: {canvas_name}") from exc


# ---------------------------------------------------------------------------
# Background generation (unchanged)
# ---------------------------------------------------------------------------

def build_background(idx: int, slide_spec: dict, asset_dir: Path, width: int, height: int) -> Path:
    canvas_name = slide_spec.get("_canvas_name", "widescreen")
    if canvas_name == "lecture-vertical":
        return build_lecture_background(idx, slide_spec, asset_dir, width, height)

    return build_poster_background(idx, slide_spec, asset_dir, width, height)


def build_poster_background(idx: int, slide_spec: dict, asset_dir: Path, width: int, height: int) -> Path:
    asset_dir.mkdir(parents=True, exist_ok=True)
    accent_cycle = [
        (COLORS["RED"], COLORS["YELLOW"], COLORS["CYAN"]),
        (COLORS["CYAN"], COLORS["PURPLE"], COLORS["PINK"]),
        (COLORS["YELLOW"], COLORS["TEAL"], COLORS["ORANGE"]),
        (COLORS["BLUE"], COLORS["PINK"], COLORS["LIME"]),
    ]
    palette = accent_cycle[idx % len(accent_cycle)]
    a1, a2, a3 = to_rgb(palette[0]), to_rgb(palette[1]), to_rgb(palette[2])

    img = Image.new("RGBA", (width, height), (0, 0, 0, 255))

    glow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(glow, "RGBA")
    cx, cy = width // 2, height // 2
    gdraw.ellipse((cx - int(width * 0.22), cy - int(height * 0.28), cx + int(width * 0.22), cy + int(height * 0.28)), fill=a1 + (14,))
    gdraw.ellipse((int(width * 0.72), int(height * 0.08), width + 60, int(height * 0.42)), fill=a2 + (11,))
    gdraw.ellipse((-60, int(height * 0.62), int(width * 0.28), height + 40), fill=a3 + (12,))
    glow = glow.filter(ImageFilter.GaussianBlur(radius=35))
    img = Image.alpha_composite(img, glow)
    draw = ImageDraw.Draw(img, "RGBA")

    grid_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw_grid = ImageDraw.Draw(grid_layer, "RGBA")
    grid_step = max(60, width // 24)
    for x in range(0, width, grid_step):
        gdraw_grid.line((x, 0, x, height), fill=(255, 255, 255, 12), width=1)
    for y in range(0, height, grid_step):
        gdraw_grid.line((0, y, width, y), fill=(255, 255, 255, 10), width=1)
    grid_layer = grid_layer.filter(ImageFilter.GaussianBlur(radius=2))
    img = Image.alpha_composite(img, grid_layer)
    draw = ImageDraw.Draw(img, "RGBA")

    ghost = slide_spec.get("ghost", "")
    if ghost:
        ghost_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        ghost_draw = ImageDraw.Draw(ghost_layer)
        ghost_draw.text((width - 100, int(height * 0.18)), ghost, font=pil_font(max(140, min(width, height) // 5)), fill=a3 + (18,), anchor="ra")
        ghost_layer = ghost_layer.filter(ImageFilter.GaussianBlur(radius=3))
        img = Image.alpha_composite(img, ghost_layer)
        draw = ImageDraw.Draw(img, "RGBA")

    draw.rounded_rectangle((24, 24, width - 24, height - 24), radius=10, outline=(255, 255, 255, 25), width=1)

    output = asset_dir / f"poster_bg_{idx + 1:02d}.jpg"
    img.convert("RGB").save(output, quality=90, optimize=True, progressive=True)
    return output


def add_lecture_scanlines(image: Image.Image) -> Image.Image:
    width, height = image.size
    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay, "RGBA")
    for y in range(0, height, 6):
        alpha = 8 if y % 12 == 0 else 4
        draw.line((0, y, width, y), fill=(255, 255, 255, alpha), width=1)
    return Image.alpha_composite(image, overlay)


def draw_layered_glow(draw: ImageDraw.ImageDraw, center: tuple[int, int], radii: list[int], color_value: tuple[int, int, int], alphas: list[int]) -> None:
    x, y = center
    for radius, alpha in zip(radii, alphas):
        draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=color_value + (alpha,))


def add_lecture_orb(image: Image.Image, palette: list[tuple[int, int, int]]) -> Image.Image:
    width, height = image.size
    layer = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer, "RGBA")
    cx = width // 2
    cy = height - 210
    r = 88

    draw.ellipse((cx - 150, cy - 42, cx + 150, cy + 78), fill=(255, 175, 90, 30))
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(230, 232, 238, 220))
    draw.ellipse((cx - r + 12, cy - r + 16, cx + r - 12, cy + r - 10), fill=(82, 86, 102, 170))
    draw.ellipse((cx - 46, cy - 58, cx + 18, cy - 8), fill=(255, 255, 255, 148))
    for idx, color_value in enumerate(palette):
        rr = 96 + idx * 52
        draw.ellipse((cx - rr, cy - rr, cx + rr, cy + rr), outline=color_value + (36,), width=2)
    return Image.alpha_composite(image, layer)


def build_lecture_background(idx: int, slide_spec: dict, asset_dir: Path, width: int, height: int) -> Path:
    asset_dir.mkdir(parents=True, exist_ok=True)
    accent_cycle = [
        (COLORS["RED"], COLORS["YELLOW"], COLORS["CYAN"]),
        (COLORS["CYAN"], COLORS["PURPLE"], COLORS["PINK"]),
        (COLORS["YELLOW"], COLORS["TEAL"], COLORS["ORANGE"]),
        (COLORS["BLUE"], COLORS["PINK"], COLORS["LIME"]),
    ]
    palette = accent_cycle[idx % len(accent_cycle)]
    a1, a2, a3 = to_rgb(palette[0]), to_rgb(palette[1]), to_rgb(palette[2])

    img = Image.new("RGBA", (width, height), (0, 0, 0, 255))

    glow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw = ImageDraw.Draw(glow, "RGBA")
    draw_layered_glow(gdraw, (width // 2, 320), [180, 130, 80], a1, [14, 10, 7])
    draw_layered_glow(gdraw, (140, 960), [160, 110, 70], a2, [12, 9, 6])
    draw_layered_glow(gdraw, (880, 1240), [150, 100, 60], a3, [12, 8, 5])
    draw_layered_glow(gdraw, (width // 2, 1580), [110, 70], a1, [8, 5])
    glow = glow.filter(ImageFilter.GaussianBlur(radius=30))
    img = Image.alpha_composite(img, glow)

    grid_layer = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    gdraw_grid = ImageDraw.Draw(grid_layer, "RGBA")
    grid_step = max(54, width // 20)
    for x in range(0, width, grid_step):
        gdraw_grid.line((x, 0, x, height), fill=(255, 255, 255, 10), width=1)
    for y in range(0, height, grid_step):
        gdraw_grid.line((0, y, width, y), fill=(255, 255, 255, 8), width=1)
    grid_layer = grid_layer.filter(ImageFilter.GaussianBlur(radius=2))
    img = Image.alpha_composite(img, grid_layer)

    img = add_lecture_scanlines(img)
    img = add_lecture_orb(img, [a1, a2, a3])

    draw = ImageDraw.Draw(img, "RGBA")
    draw.rounded_rectangle((20, 20, width - 20, height - 20), radius=8, outline=(255, 255, 255, 24), width=1)

    output = asset_dir / f"lecture_bg_{idx + 1:02d}.png"
    img.save(output)
    return output


# ---------------------------------------------------------------------------
# Shared shape helpers
# ---------------------------------------------------------------------------

def add_textbox(slide, left, top, width, height, paragraphs, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, auto_fit=False):
    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = valign
    frame.margin_left = Pt(4)
    frame.margin_right = Pt(4)
    frame.margin_top = Pt(2)
    frame.margin_bottom = Pt(2)

    # Auto-fit: measure total text height and shrink if needed
    if auto_fit and paragraphs:
        box_width_px = int(width / 12700)
        box_height_px = int(height / 12700)
        if box_width_px > 20 and box_height_px > 10:
            full_text = "\n".join(p["text"] for p in paragraphs)
            font_name = paragraphs[0].get("font", "Noto Sans CJK SC")
            font_path = _resolve_font_path(font_name)
            max_pt = paragraphs[0].get("size", 18)
            inner_width = box_width_px - 10
            best_pt = fit_text_to_box(full_text, font_path, inner_width, box_height_px - 6, max_pt=max_pt, min_pt=8)
            if best_pt < max_pt:
                for p in paragraphs:
                    p["size"] = best_pt

    for idx, spec in enumerate(paragraphs):
        paragraph = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        paragraph.alignment = align
        paragraph.space_after = Pt(spec.get("space_after", 0))
        paragraph.line_spacing = spec.get("line_spacing", 1.0)
        run = paragraph.add_run()
        run.text = spec["text"]
        font = run.font
        font.name = spec.get("font", "Noto Sans CJK SC")
        font.size = Pt(spec.get("size", 18))
        font.bold = spec.get("bold", True)
        text_color = spec.get("color", COLORS["WHITE"])
        font.color.rgb = text_color
        glow_size = spec.get("glow", 0)
        if glow_size > 0:
            add_glow_to_run(run, text_color, size=glow_size)
    return box


def add_tag(slide, text, canvas_name="widescreen"):
    if canvas_name == "lecture-vertical":
        add_textbox(
            slide,
            px(240),
            px(110),
            px(600),
            px(28),
            [{"text": text, "size": 12, "bold": False, "color": COLORS["MUTED"]}],
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        return

    width_px = 430 if canvas_name == "widescreen" else 300
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(82), px(60), px(width_px), px(40))
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLORS["CARD"]
    shape.fill.transparency = 0.30
    shape.line.color.rgb = RGBColor(255, 255, 255)
    shape.line.width = Pt(0.8)
    add_glow_to_shape(shape, COLORS["CYAN"], size=25000)
    add_textbox(slide, px(108), px(69), px(width_px - 40), px(24), [{"text": text, "size": 12, "bold": False, "color": COLORS["MUTED"]}])


def add_page_no(slide, num, canvas_name="widescreen"):
    if canvas_name == "widescreen":
        add_textbox(slide, px(1760), px(995), px(80), px(32), [{"text": f"{num:02d}", "size": 15, "bold": True, "color": COLORS["SOFT"]}], align=PP_ALIGN.RIGHT)
        add_textbox(slide, px(112), px(995), px(420), px(24), [{"text": "POSTER MODE / CYBERPUNK PPT", "size": 11, "bold": False, "color": COLORS["SOFT"]}])
    elif canvas_name == "lecture-vertical":
        add_textbox(slide, px(930), px(1818), px(60), px(24), [{"text": f"{num:02d}", "size": 12, "bold": False, "color": COLORS["MUTED"]}], align=PP_ALIGN.RIGHT)
    else:
        add_textbox(slide, px(900), px(1350), px(80), px(32), [{"text": f"{num:02d}", "size": 15, "bold": True, "color": COLORS["SOFT"]}], align=PP_ALIGN.RIGHT)
        add_textbox(slide, px(90), px(1350), px(320), px(24), [{"text": "XHS / CYBERPUNK COVER", "size": 11, "bold": False, "color": COLORS["SOFT"]}])


# ---------------------------------------------------------------------------
# Title blocks — fixed sizing to prevent text overflow
# ---------------------------------------------------------------------------

def add_title_block(slide, title_lines, subtitle, left_px=118, top_px=168, width_px=980):
    """Render title lines + subtitle with measured text. Returns Y position after the entire block."""
    y = top_px
    for item in title_lines:
        pixel_size = int(item["size"])
        pt_size = max(26, int(pixel_size * 0.46))
        text_color = color(item["color"])
        # Measure actual text to determine box height
        metrics = measure_text(item["text"], FONT_PATH_BLACK, pt_size, width_px - 10)
        box_h = max(_box_height_for_pt(pt_size), metrics["total_height_px"] + 12)
        add_textbox(
            slide,
            px(left_px),
            px(y),
            px(width_px),
            px(box_h),
            [{"text": item["text"], "size": pt_size, "color": text_color, "glow": 54000}],
        )
        y += box_h + 10
    if subtitle:
        sub_text = " ".join(subtitle)
        sub_metrics = measure_text(sub_text, FONT_PATH_BLACK, 18, min(width_px, 860) - 10)
        sub_h = max(72, sub_metrics["total_height_px"] + 16)
        add_textbox(
            slide,
            px(left_px + 4),
            px(y + 10),
            px(min(width_px, 860)),
            px(sub_h),
            [{"text": sub_text, "size": 18, "bold": False, "color": COLORS["WHITE"], "line_spacing": 1.05}],
        )
        y += sub_h + 14
    return y


def add_title_block_vertical(slide, title_lines, subtitle, left_px=88, top_px=176, width_px=900):
    y = top_px
    for item in title_lines:
        pixel_size = int(item["size"])
        pt_size = max(24, int(pixel_size * 0.38))
        text_color = color(item["color"])
        metrics = measure_text(item["text"], FONT_PATH_BLACK, pt_size, width_px - 10)
        box_h = max(_box_height_for_pt(pt_size), metrics["total_height_px"] + 12)
        add_textbox(
            slide,
            px(left_px),
            px(y),
            px(width_px),
            px(box_h),
            [{"text": item["text"], "size": pt_size, "color": text_color, "glow": 48000}],
        )
        y += box_h + 10
    if subtitle:
        sub_text = " ".join(subtitle)
        sub_metrics = measure_text(sub_text, FONT_PATH_BLACK, 15, width_px - 30)
        sub_h = max(60, sub_metrics["total_height_px"] + 14)
        add_textbox(
            slide,
            px(left_px + 4),
            px(y + 8),
            px(width_px - 20),
            px(sub_h),
            [{"text": sub_text, "size": 15, "bold": False, "color": COLORS["WHITE"], "line_spacing": 1.02}],
        )
        y += sub_h + 10
    return y


def add_title_block_lecture(slide, title_lines, subtitle, top_px=260, width_px=820):
    y = top_px
    center_left = (1080 - width_px) // 2
    for item in title_lines:
        pixel_size = int(item["size"])
        pt_size = max(24, int(pixel_size * 0.36))
        text_color = color(item["color"])
        metrics = measure_text(item["text"], FONT_PATH_BLACK, pt_size, width_px - 10)
        box_h = max(_box_height_for_pt(pt_size), metrics["total_height_px"] + 12)
        add_textbox(
            slide,
            px(center_left),
            px(y),
            px(width_px),
            px(box_h),
            [{"text": item["text"], "size": pt_size, "color": text_color, "glow": 46000}],
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        y += box_h + 10
    if subtitle:
        sub_text = " ".join(subtitle)
        sub_metrics = measure_text(sub_text, FONT_PATH_BLACK, 15, width_px - 30)
        sub_h = max(60, sub_metrics["total_height_px"] + 14)
        add_textbox(
            slide,
            px(center_left + 10),
            px(y + 14),
            px(width_px - 20),
            px(sub_h),
            [{"text": sub_text, "size": 15, "bold": False, "color": COLORS["WHITE"], "line_spacing": 1.05}],
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        y += sub_h + 16
    return y


# ---------------------------------------------------------------------------
# Panel & chip helpers
# ---------------------------------------------------------------------------

def add_panel(slide, left_px, top_px, width_px, height_px, title, lines, accent_name, mono=False, title_size=18, body_size=16, canvas_name="widescreen"):
    accent = color(accent_name)
    safe = SLIDE_SAFE.get(canvas_name, SLIDE_SAFE["widescreen"])
    height_px = min(height_px, safe["max_y"] - top_px)
    if height_px < 60:
        height_px = 60
    add_gradient_panel(slide, left_px, top_px, width_px, height_px, accent_name)
    # Accent line under title
    add_accent_line(slide, left_px + 24, top_px + 44, min(width_px - 48, 200), accent_name, thickness=2)
    add_textbox(slide, px(left_px + 24), px(top_px + 14), px(width_px - 48), px(28), [{"text": title, "size": title_size, "color": accent}])
    body_font = "DejaVu Sans Mono" if mono else "Noto Sans CJK SC"
    effective_body_size = 14 if mono else body_size
    paragraphs = [{"text": line, "size": effective_body_size, "bold": False, "color": COLORS["WHITE"], "font": body_font, "space_after": 4} for line in lines]
    body_top = top_px + 52
    body_height = max(20, height_px - 60)
    add_textbox(slide, px(left_px + 24), px(body_top), px(width_px - 48), px(body_height), paragraphs, auto_fit=True)


def add_chip(slide, left_px, top_px, text, color_name):
    accent = color(color_name)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(left_px), px(top_px), px(230), px(50))
    fill = shape.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = RGBColor(10, 10, 18)
    fill.gradient_stops[0].position = 0.0
    fill.gradient_stops[1].color.rgb = RGBColor(20, 20, 30)
    fill.gradient_stops[1].position = 1.0
    shape.fill.transparency = 0.25
    shape.line.color.rgb = RGBColor(255, 255, 255)
    shape.line.width = Pt(1.0)
    add_glow_to_shape(shape, accent, size=30000)
    add_outer_shadow(shape, color_rgb="%02X%02X%02X" % (accent[0], accent[1], accent[2]),
                     blur_rad=38000, dist=12700, direction=5400000, alpha_pct=20000)
    add_textbox(slide, px(left_px + 16), px(top_px + 12), px(198), px(22), [{"text": text, "size": 13, "color": accent}], align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# Widescreen renderers — dynamic positioning, boundary-aware
# ---------------------------------------------------------------------------

def render_cover(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), left_px=118, top_px=160, width_px=980)
    # Decorative accent line below title
    add_accent_line(slide, 118, bottom + 6, 320, "CYAN", thickness=3)
    # Card on the right, vertically centered relative to title
    cards = spec.get("cards", [])
    if cards:
        title_mid = (160 + bottom) // 2
        card_y = _clamp(title_mid - 100, 200, safe["max_y"] - 250)
        card_h = min(240, safe["max_y"] - card_y - 10)
        add_panel(slide, 1280, card_y, 520, card_h, cards[0]["title"], cards[0].get("lines", []), cards[0]["accent"])
    # Chips below title
    chip_y = _clamp(bottom + 20, 500, safe["max_y"] - 60)
    for i, chip in enumerate(spec.get("chips", [])[:4]):
        add_chip(slide, 120 + i * 255, chip_y, chip["text"], chip["color"])
    ghost = spec.get("ghost", "")
    if ghost:
        ghost_y = _clamp(bottom - 60, 300, safe["max_y"] - 80)
        add_textbox(slide, px(1400), px(ghost_y), px(340), px(100), [{"text": ghost, "size": 36, "color": COLORS["WHITE"], "glow": 30000}], align=PP_ALIGN.CENTER)


def render_poster_cards(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=940)
    add_accent_line(slide, 118, bottom + 6, 280, "PINK", thickness=2)
    cards = spec.get("cards", [])
    base_y = _clamp(bottom + 50, 440, 660)
    card_h = min(220, safe["max_y"] - base_y - 20)
    positions = [(118, base_y + 20), (690, base_y), (1262, base_y + 20)]
    for card, (x, y) in zip(cards[:3], positions):
        add_panel(slide, x, y, 520, card_h, card["title"], card.get("lines", []), card["accent"])


def render_flow(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=940)
    nodes = spec.get("nodes", [])
    base_y = _clamp(bottom + 40, 460, 600)
    node_w = 330
    gap = 40
    total = len(nodes[:4]) * node_w + max(0, len(nodes[:4]) - 1) * gap
    start_x = max(96, (1920 - total) // 2)
    arrows = ["CYAN", "PINK", "YELLOW"]
    node_h = min(148, safe["max_y"] - base_y - 20)
    for i, (node, x_off) in enumerate(zip(nodes[:4], range(len(nodes[:4])))):
        x = start_x + x_off * (node_w + gap)
        add_panel(slide, x, base_y, node_w, node_h, node["title"], [node["body"]], node["accent"])
        if i < min(3, len(nodes[:4]) - 1):
            arrow = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, px(x + node_w + 6), px(base_y + node_h // 2 - 20), px(28), px(40))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = color(arrows[i])
            arrow.line.color.rgb = color(arrows[i])


def render_grid_four(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=980)
    cards = spec.get("cards", [])
    base_y = _clamp(bottom + 40, 460, 620)
    card_w = min(830, (1920 - 118 - 118 - 40) // 2)
    card_h = min(180, (safe["max_y"] - base_y - 20) // 2 - 10)
    col2_x = 118 + card_w + 40
    positions = [(118, base_y), (col2_x, base_y), (118, base_y + card_h + 14), (col2_x, base_y + card_h + 14)]
    for card, (x, y) in zip(cards[:4], positions):
        add_panel(slide, x, y, card_w, card_h, card["title"], card.get("lines", []), card["accent"])


def render_split(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=980)
    base_y = _clamp(bottom + 40, 460, 620)
    panel_w = min(760, (1920 - 118 - 80 - 40) // 2)
    panel_h = min(250, safe["max_y"] - base_y - 20)
    left = spec["left"]
    right = spec["right"]
    add_panel(slide, 118, base_y, panel_w, panel_h, left["title"], left.get("lines", []), left["accent"])
    add_panel(slide, 118 + panel_w + 40, base_y, panel_w, panel_h, right["title"], right.get("lines", []), right["accent"])


def render_code_mix(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=960)
    base_y = _clamp(bottom + 30, 420, 560)
    code_h = min(300, safe["max_y"] - base_y - 20)
    add_panel(slide, 118, base_y, 740, code_h, "目录 / 命令", spec.get("code", []), "CYAN", mono=True)
    cards = spec.get("cards", [])
    card_x = 920
    card_w = min(700, 1920 - card_x - 50)
    card_h = min(110, (safe["max_y"] - base_y - 20) // max(1, len(cards[:3])) - 10)
    for idx, card in enumerate(cards[:3]):
        cy = base_y + idx * (card_h + 12)
        add_panel(slide, card_x, cy, card_w, card_h, card["title"], card.get("lines", []), card["accent"])


def render_timeline(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=960)
    steps = spec.get("steps", [])
    if not steps:
        return
    line_y = _clamp(bottom + 80, 600, 780)
    # Horizontal line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(160), px(line_y), px(1560), px(4))
    line.fill.solid()
    line.fill.fore_color.rgb = COLORS["CYAN"]
    line.fill.transparency = 0.18
    line.line.fill.background()
    n = len(steps[:5])
    gap = 1560 // max(1, n - 1) if n > 1 else 0
    xs = [210 + i * gap for i in range(n)]
    dot_y = line_y - 22
    for si, (step, x) in enumerate(zip(steps[:5], xs)):
        accent = color(step["accent"])
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, px(x - 22), px(dot_y), px(44), px(44))
        dot.fill.solid()
        dot.fill.fore_color.rgb = accent
        dot.line.color.rgb = accent
        add_textbox(slide, px(x - 20), px(dot_y + 10), px(40), px(18), [{"text": step["num"], "size": 11, "color": COLORS["CARD_2"]}], align=PP_ALIGN.CENTER)
        label_y = dot_y - 90 if si % 2 == 0 else line_y + 24
        label_y = _clamp(label_y, bottom + 20, safe["max_y"] - 90)
        add_panel(slide, x - 115, label_y, 230, 78, step["label"], [], step["accent"])


def render_wide_stack(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), width_px=980)
    rows = spec.get("rows", [])
    base_y = _clamp(bottom + 30, 440, 580)
    row_h = min(90, (safe["max_y"] - base_y - 10) // max(1, len(rows[:4])) - 6)
    for i, row in enumerate(rows[:4]):
        add_panel(slide, 118, base_y + i * (row_h + 8), 1680, row_h, row["title"], [row["body"]], row["accent"])


def render_statement(slide, spec):
    safe = SLIDE_SAFE["widescreen"]
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), left_px=150, top_px=220, width_px=930)
    lines = spec.get("lines", [])
    base_y = _clamp(bottom + 50, 500, safe["max_y"] - 80)
    n = len(lines[:4])
    if n == 0:
        return
    item_w = min(300, (1600) // n)
    total_w = n * item_w
    start_x = max(120, (1920 - total_w) // 2)
    for i, item in enumerate(lines[:4]):
        add_textbox(slide, px(start_x + i * item_w), px(base_y), px(item_w - 10), px(70), [{"text": item["text"], "size": 32, "color": color(item["color"])}], align=PP_ALIGN.CENTER)


def render_ending(slide, spec):
    bottom = add_title_block(slide, spec["title"], spec.get("subtitle", []), left_px=150, top_px=250, width_px=980)
    add_accent_line(slide, 150, bottom + 6, 400, "CYAN", thickness=2)
    add_textbox(slide, px(360), px(860), px(1200), px(30), [{"text": spec.get("footer", ""), "size": 12, "bold": False, "color": COLORS["MUTED"]}], align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# XHS vertical renderers
# ---------------------------------------------------------------------------

def render_cover_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=180, width_px=820)
    chips = spec.get("chips", [])
    chip_start_y = _clamp(bottom + 40, 400, safe["max_y"] - 200)
    for idx, chip in enumerate(chips[:4]):
        row = idx // 2
        col = idx % 2
        add_chip(slide, 88 + col * 300, chip_start_y + row * 74, chip["text"], chip["color"])
    cards = spec.get("cards", [])
    if cards:
        card = cards[0]
        card_y = _clamp(chip_start_y + len(chips[:4]) * 74 + 30, bottom + 100, safe["max_y"] - 200)
        card_h = min(180, safe["max_y"] - card_y - 20)
        add_panel(slide, 88, card_y, 900, card_h, card["title"], card.get("lines", []), card["accent"])
    ghost = spec.get("ghost", "")
    if ghost:
        ghost_y = safe["max_y"] - 160
        add_textbox(slide, px(720), px(ghost_y), px(240), px(100), [{"text": ghost, "size": 34, "color": COLORS["WHITE"], "glow": 28000}], align=PP_ALIGN.CENTER)


def render_poster_cards_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = _clamp(bottom + 40, 400, 600)
    card_h = min(150, (safe["max_y"] - start_y - 10) // max(1, len(spec.get("cards", [])[:3])) - 10)
    for idx, card in enumerate(spec.get("cards", [])[:3]):
        add_panel(slide, 88, start_y + idx * (card_h + 14), 900, card_h, card["title"], card.get("lines", []), card["accent"])


def render_flow_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = _clamp(bottom + 40, 380, 580)
    nodes = spec.get("nodes", [])[:4]
    node_h = min(118, (safe["max_y"] - start_y - 10) // max(1, len(nodes)) - 30)
    for idx, node in enumerate(nodes):
        y = start_y + idx * (node_h + 40)
        add_panel(slide, 130, y, 820, node_h, node["title"], [node["body"]], node["accent"])
        if idx < len(nodes) - 1:
            add_textbox(
                slide,
                px(490),
                px(y + node_h + 4),
                px(100),
                px(30),
                [{"text": "▼", "size": 22, "color": color(nodes[idx + 1]["accent"])}],
                align=PP_ALIGN.CENTER,
            )


def render_grid_four_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = _clamp(bottom + 40, 400, 600)
    card_w = (900 - 20) // 2
    card_h = min(180, (safe["max_y"] - start_y - 10) // 2 - 10)
    positions = [(88, start_y), (88 + card_w + 20, start_y), (88, start_y + card_h + 14), (88 + card_w + 20, start_y + card_h + 14)]
    for card, (x, y) in zip(spec.get("cards", [])[:4], positions):
        add_panel(slide, x, y, card_w, card_h, card["title"], card.get("lines", []), card["accent"])


def render_split_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = _clamp(bottom + 40, 420, 650)
    panel_h = min(180, (safe["max_y"] - start_y - 10) // 2 - 10)
    add_panel(slide, 88, start_y, 900, panel_h, spec["left"]["title"], spec["left"].get("lines", []), spec["left"]["accent"])
    add_panel(slide, 88, start_y + panel_h + 14, 900, panel_h, spec["right"]["title"], spec["right"].get("lines", []), spec["right"]["accent"])


def render_code_mix_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    code_y = _clamp(bottom + 30, 380, 560)
    code_h = min(220, safe["max_y"] - code_y - 10)
    add_panel(slide, 88, code_y, 900, code_h, "目录 / 命令", spec.get("code", []), "CYAN", mono=True)
    cards = spec.get("cards", [])
    card_h = min(108, (safe["max_y"] - code_y - code_h - 10) // max(1, len(cards[:3])) - 8)
    for idx, card in enumerate(cards[:3]):
        cy = code_y + code_h + 14 + idx * (card_h + 10)
        add_panel(slide, 88, cy, 900, card_h, card["title"], card.get("lines", []), card["accent"])


def render_timeline_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = _clamp(bottom + 40, 400, 600)
    steps = spec.get("steps", [])[:5]
    step_h = min(96, (safe["max_y"] - start_y - 10) // max(1, len(steps)) - 10)
    for idx, step in enumerate(steps):
        label = f"{step['num']}  {step['label']}"
        add_panel(slide, 140, start_y + idx * (step_h + 12), 800, step_h, label, [], step["accent"])


def render_wide_stack_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=176, width_px=900)
    start_y = _clamp(bottom + 40, 400, 600)
    rows = spec.get("rows", [])[:4]
    row_h = min(112, (safe["max_y"] - start_y - 10) // max(1, len(rows)) - 8)
    for idx, row in enumerate(rows):
        add_panel(slide, 88, start_y + idx * (row_h + 10), 900, row_h, row["title"], [row["body"]], row["accent"])


def render_statement_vertical(slide, spec):
    safe = SLIDE_SAFE["xhs-vertical"]
    bottom = add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=220, width_px=900)
    start_y = _clamp(bottom + 50, 500, safe["max_y"] - 200)
    for idx, item in enumerate(spec.get("lines", [])[:4]):
        add_textbox(
            slide,
            px(120),
            px(start_y + idx * 60),
            px(840),
            px(48),
            [{"text": item["text"], "size": 26, "color": color(item["color"])}],
            align=PP_ALIGN.CENTER,
        )


def render_ending_vertical(slide, spec):
    add_title_block_vertical(slide, spec["title"], spec.get("subtitle", []), left_px=88, top_px=220, width_px=900)
    add_textbox(
        slide,
        px(120),
        px(1200),
        px(840),
        px(42),
        [{"text": spec.get("footer", ""), "size": 14, "bold": False, "color": COLORS["MUTED"]}],
        align=PP_ALIGN.CENTER,
    )


# ---------------------------------------------------------------------------
# Lecture vertical renderers
# ---------------------------------------------------------------------------

def render_cover_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=330, width_px=840)
    chips = spec.get("chips", [])
    chip_y = _clamp(bottom + 30, 600, safe["max_y"] - 200)
    for idx, chip in enumerate(chips[:4]):
        row = idx // 2
        col = idx % 2
        add_chip(slide, 152 + col * 392, chip_y + row * 82, chip["text"], chip["color"])
    cards = spec.get("cards", [])
    if cards:
        card_y = _clamp(chip_y + len(chips[:4]) * 82 + 30, bottom + 100, safe["max_y"] - 200)
        card_h = min(190, safe["max_y"] - card_y - 20)
        add_panel(slide, 120, card_y, 840, card_h, cards[0]["title"], cards[0].get("lines", []), cards[0]["accent"], body_size=15)


def render_poster_cards_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = _clamp(bottom + 50, 600, 900)
    cards = spec.get("cards", [])[:3]
    card_h = min(132, (safe["max_y"] - start_y - 10) // max(1, len(cards)) - 10)
    for idx, card in enumerate(cards):
        add_panel(slide, 118, start_y + idx * (card_h + 14), 844, card_h, card["title"], card.get("lines", []), card["accent"], body_size=15)


def render_grid_four_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = _clamp(bottom + 50, 580, 860)
    card_w = (844 - 20) // 2
    card_h = min(164, (safe["max_y"] - start_y - 10) // 2 - 10)
    positions = [(108, start_y), (108 + card_w + 20, start_y), (108, start_y + card_h + 14), (108 + card_w + 20, start_y + card_h + 14)]
    for card, (x, y) in zip(spec.get("cards", [])[:4], positions):
        add_panel(slide, x, y, card_w, card_h, card["title"], card.get("lines", []), card["accent"], body_size=14)


def render_split_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = _clamp(bottom + 50, 600, 900)
    panel_h = min(184, (safe["max_y"] - start_y - 10) // 2 - 10)
    add_panel(slide, 118, start_y, 844, panel_h, spec["left"]["title"], spec["left"].get("lines", []), spec["left"]["accent"], body_size=14)
    add_panel(slide, 118, start_y + panel_h + 14, 844, panel_h, spec["right"]["title"], spec["right"].get("lines", []), spec["right"]["accent"], body_size=14)


def render_code_mix_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = _clamp(bottom + 40, 560, 820)
    code_h = min(250, safe["max_y"] - start_y - 10)
    add_panel(slide, 118, start_y, 844, code_h, "目录 / 命令", spec.get("code", []), "CYAN", mono=True)
    cards = spec.get("cards", [])
    card_h = min(128, (safe["max_y"] - start_y - code_h - 10) // max(1, len(cards[:3])) - 10)
    for idx, card in enumerate(cards[:3]):
        cy = start_y + code_h + 14 + idx * (card_h + 10)
        add_panel(slide, 118, cy, 844, card_h, card["title"], card.get("lines", []), card["accent"], body_size=13)


def render_flow_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = _clamp(bottom + 50, 600, 900)
    nodes = spec.get("nodes", [])[:4]
    node_h = min(106, (safe["max_y"] - start_y - 10) // max(1, len(nodes)) - 30)
    for idx, node in enumerate(nodes):
        y = start_y + idx * (node_h + 36)
        add_panel(slide, 162, y, 756, node_h, node["title"], [node["body"]], node["accent"], body_size=14)
        if idx < len(nodes) - 1:
            add_textbox(slide, px(500), px(y + node_h + 4), px(80), px(28), [{"text": "▼", "size": 18, "color": color(nodes[idx + 1]["accent"])}], align=PP_ALIGN.CENTER)


def render_timeline_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = _clamp(bottom + 50, 600, 900)
    steps = spec.get("steps", [])[:5]
    step_h = min(92, (safe["max_y"] - start_y - 10) // max(1, len(steps)) - 10)
    for idx, step in enumerate(steps):
        add_panel(slide, 128, start_y + idx * (step_h + 12), 824, step_h, f"{step['num']}  {step['label']}", [], step["accent"], body_size=14)


def render_wide_stack_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=260, width_px=860)
    start_y = _clamp(bottom + 50, 600, 900)
    rows = spec.get("rows", [])[:4]
    row_h = min(112, (safe["max_y"] - start_y - 10) // max(1, len(rows)) - 8)
    for idx, row in enumerate(rows):
        add_panel(slide, 118, start_y + idx * (row_h + 10), 844, row_h, row["title"], [row["body"]], row["accent"], body_size=14)


def render_statement_lecture(slide, spec):
    safe = SLIDE_SAFE["lecture-vertical"]
    bottom = add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=280, width_px=860)
    line_y = _clamp(bottom + 60, 700, safe["max_y"] - 200)
    for idx, item in enumerate(spec.get("lines", [])[:4]):
        add_textbox(
            slide,
            px(140),
            px(line_y + idx * 60),
            px(800),
            px(42),
            [{"text": item["text"], "size": 26, "color": color(item["color"])}],
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )


def render_ending_lecture(slide, spec):
    add_title_block_lecture(slide, spec["title"], spec.get("subtitle", []), top_px=320, width_px=860)
    add_textbox(
        slide,
        px(160),
        px(1180),
        px(760),
        px(80),
        [{"text": spec.get("footer", ""), "size": 14, "bold": False, "color": COLORS["MUTED"], "line_spacing": 1.05}],
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )


# ---------------------------------------------------------------------------
# Renderer registries
# ---------------------------------------------------------------------------

RENDERERS = {
    "cover": render_cover,
    "poster_cards": render_poster_cards,
    "flow": render_flow,
    "grid_four": render_grid_four,
    "split": render_split,
    "code_mix": render_code_mix,
    "timeline": render_timeline,
    "wide_stack": render_wide_stack,
    "statement": render_statement,
    "ending": render_ending,
}

VERTICAL_RENDERERS = {
    "cover": render_cover_vertical,
    "poster_cards": render_poster_cards_vertical,
    "flow": render_flow_vertical,
    "grid_four": render_grid_four_vertical,
    "split": render_split_vertical,
    "code_mix": render_code_mix_vertical,
    "timeline": render_timeline_vertical,
    "wide_stack": render_wide_stack_vertical,
    "statement": render_statement_vertical,
    "ending": render_ending_vertical,
}

LECTURE_VERTICAL_RENDERERS = {
    "cover": render_cover_lecture,
    "poster_cards": render_poster_cards_lecture,
    "flow": render_flow_lecture,
    "grid_four": render_grid_four_lecture,
    "split": render_split_lecture,
    "code_mix": render_code_mix_lecture,
    "timeline": render_timeline_lecture,
    "wide_stack": render_wide_stack_lecture,
    "statement": render_statement_lecture,
    "ending": render_ending_lecture,
}


# ---------------------------------------------------------------------------
# Main generation pipeline
# ---------------------------------------------------------------------------

def make_presentation(spec: dict, output_path: Path, asset_dir: Path) -> None:
    canvas = get_canvas(spec)
    canvas_name = spec.get("canvas", "widescreen")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = canvas["slide_w"]
    prs.slide_height = canvas["slide_h"]

    for idx, slide_spec in enumerate(spec["slides"]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide_spec = dict(slide_spec)
        slide_spec["_canvas_name"] = canvas_name
        bg = build_background(idx, slide_spec, asset_dir, canvas["width"], canvas["height"])
        slide.shapes.add_picture(str(bg), 0, 0, width=canvas["slide_w"], height=canvas["slide_h"])
        add_tag(slide, slide_spec.get("tag", f"CUT {idx + 1:02d}"), canvas_name=canvas_name)
        add_page_no(slide, idx + 1, canvas_name=canvas_name)
        if canvas_name == "xhs-vertical":
            VERTICAL_RENDERERS[slide_spec["layout"]](slide, slide_spec)
        elif canvas_name == "lecture-vertical":
            LECTURE_VERTICAL_RENDERERS[slide_spec["layout"]](slide, slide_spec)
        else:
            RENDERERS[slide_spec["layout"]](slide, slide_spec)

    prs.save(str(output_path))


def load_spec(spec_path: Path) -> dict:
    return json.loads(spec_path.read_text(encoding="utf-8"))


def export_pdf(pptx_path: Path, pdf_output: Path) -> None:
    pdf_output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="cyberpunk-pdf-") as tmpdir:
        tmpdir_path = Path(tmpdir)
        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                str(pptx_path),
                "--outdir",
                str(tmpdir_path),
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        generated_pdf = tmpdir_path / f"{pptx_path.stem}.pdf"
        if not generated_pdf.exists():
            raise FileNotFoundError(f"PDF export failed for {pptx_path}")
        shutil.copy2(generated_pdf, pdf_output)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate editable cyberpunk PPT from JSON spec.")
    parser.add_argument("--spec", required=True, help="Path to JSON spec file.")
    parser.add_argument("--output", required=True, help="Output PPTX path.")
    parser.add_argument("--assets-dir", default="generated_cyberpunk_assets", help="Directory for generated background assets.")
    parser.add_argument("--pdf-output", help="Optional output PDF path.")
    args = parser.parse_args()

    spec_path = Path(args.spec)
    output_path = Path(args.output)
    asset_dir = Path(args.assets_dir)

    spec = load_spec(spec_path)
    make_presentation(spec, output_path, asset_dir)
    if args.pdf_output:
        export_pdf(output_path, Path(args.pdf_output))
    print(f"Generated {output_path} with {len(spec['slides'])} slides")


if __name__ == "__main__":
    main()
