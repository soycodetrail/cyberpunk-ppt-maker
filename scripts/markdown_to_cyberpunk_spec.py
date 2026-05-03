#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
import tempfile

from export_cyberpunk_images import export_images
from generate_cyberpunk_ppt import (
    extract_deck_title,
    export_pdf,
    make_presentation,
    resolve_output_dir,
    sanitize_dirname,
)


SECTION_KEYS = {
    "Title",
    "Subtitle",
    "Body",
    "Chips",
    "Cards",
    "Nodes",
    "Left",
    "Right",
    "Code",
    "Steps",
    "Rows",
    "Lines",
}


def split_parts(value: str, expected: int) -> list[str]:
    parts = [part.strip() for part in value.split("|")]
    if len(parts) < expected:
        raise ValueError(f"Expected at least {expected} pipe-separated parts in: {value}")
    return parts


def parse_card_like(value: str) -> dict:
    title, accent, lines_blob = split_parts(value, 3)[:3]
    lines = [line.strip() for line in lines_blob.split(";") if line.strip()]
    return {"title": title, "accent": accent, "lines": lines}


def parse_body_item(value: str) -> tuple[str, str]:
    if "|" in value:
        left, right = split_parts(value, 2)[:2]
        return left, right
    if "：" in value:
        left, right = [part.strip() for part in value.split("：", 1)]
        return left, right
    if ":" in value:
        left, right = [part.strip() for part in value.split(":", 1)]
        return left, right
    return value.strip()[:8], value.strip()


def is_enabled(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes", "on"}


def cleanup_title_source(text: str) -> str:
    cleaned = re.sub(r"[：:，,。.!！？?（）()\-\s]+", "", text)
    for token in ["为什么", "现在", "如何", "怎么", "指南", "教程", "企业", "团队", "要在", "彻底", "实现", "进行", "一个", "关于"]:
        cleaned = cleaned.replace(token, "")
    cleaned = re.sub(r"^要", "", cleaned)
    cleaned = cleaned.replace("大模型", "模型")
    return cleaned or text.strip()


def stylize_title(slide_name: str) -> list[dict]:
    source = cleanup_title_source(slide_name)
    primary = None
    secondary = None

    if "本地" in source and "部署" in source:
        primary = "本地部署"
        if "模型" in source:
            secondary = "模型上桌"
    elif "离线" in source and ("AI" in slide_name or "智能" in slide_name or "编程" in slide_name):
        primary = "彻底离线"
        secondary = "AI 进场"
    elif "编程助手" in slide_name or "代码助手" in slide_name:
        primary = "编程 AI"
        secondary = "助手上桌"
    elif "优势" in source:
        primary = "部署优势"
        secondary = "本地接管" if "本地" in slide_name else None
    elif "接入" in source:
        primary = "工作流接管"
        secondary = "接口开口"

    if primary is None:
        primary = source[: min(6, len(source))]
        secondary = source[min(6, len(source)) : min(12, len(source))] or None

    title = [{"text": primary[:8], "color": "CYAN", "size": 120}]
    if secondary:
        title.append({"text": secondary[:8], "color": "WHITE", "size": 110})
    return title


def infer_layout_from_body(slide: dict) -> None:
    if slide.get("_layout_explicit"):
        return
    body_items = slide.get("body", [])
    if not body_items:
        return
    if any(key in slide for key in ("cards", "nodes", "left", "right", "steps", "rows", "code", "lines")):
        return

    parsed = [parse_body_item(item) for item in body_items]
    if len(parsed) <= 4:
        slide["layout"] = "poster_cards" if len(parsed) <= 3 else "grid_four"
        accents = ["ORANGE", "CYAN", "PINK", "YELLOW"]
        slide["cards"] = [
            {"title": title[:10], "accent": accents[idx % len(accents)], "lines": [body]}
            for idx, (title, body) in enumerate(parsed)
        ]
    else:
        slide["layout"] = "wide_stack"
        accents = ["ORANGE", "CYAN", "PINK", "YELLOW"]
        slide["rows"] = [
            {"title": title[:12], "body": body, "accent": accents[idx % len(accents)]}
            for idx, (title, body) in enumerate(parsed)
        ]


def parse_slide_blocks(
    lines: list[str],
    slide_index: int,
    slide_name: str,
    tag_prefix: str,
    default_layout: str,
    auto_style_titles: bool,
) -> dict:
    slide: dict = {
        "layout": default_layout,
        "_layout_explicit": False,
        "tag": f"{tag_prefix} {slide_index:02d}".strip(),
        "title": [],
        "subtitle": [],
    }
    current_block = None

    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue

        if current_block and line.lstrip().startswith("- "):
            item = line.lstrip()[2:].strip()
            if current_block == "Title":
                text, color, size = split_parts(item, 3)[:3]
                slide.setdefault("title", []).append({"text": text, "color": color, "size": int(size)})
            elif current_block == "Subtitle":
                slide.setdefault("subtitle", []).append(item)
            elif current_block == "Body":
                slide.setdefault("body", []).append(item)
            elif current_block == "Chips":
                text, color = split_parts(item, 2)[:2]
                slide.setdefault("chips", []).append({"text": text, "color": color})
            elif current_block == "Cards":
                slide.setdefault("cards", []).append(parse_card_like(item))
            elif current_block == "Nodes":
                title, body, accent = split_parts(item, 3)[:3]
                slide.setdefault("nodes", []).append({"title": title, "body": body, "accent": accent})
            elif current_block in {"Left", "Right"}:
                card = parse_card_like(item)
                slide[current_block.lower()] = card
            elif current_block == "Code":
                slide.setdefault("code", []).append(item)
            elif current_block == "Steps":
                num, label, accent = split_parts(item, 3)[:3]
                slide.setdefault("steps", []).append({"num": num, "label": label, "accent": accent})
            elif current_block == "Rows":
                title, body, accent = split_parts(item, 3)[:3]
                slide.setdefault("rows", []).append({"title": title, "body": body, "accent": accent})
            elif current_block == "Lines":
                text, color = split_parts(item, 2)[:2]
                slide.setdefault("lines", []).append({"text": text, "color": color})
            continue

        block_match = re.match(r"^([A-Za-z][A-Za-z ]*):\s*(.*)$", line.strip())
        if block_match:
            key, value = block_match.groups()
            key = key.strip()
            if key in SECTION_KEYS:
                current_block = key
                if value:
                    if key == "Subtitle":
                        slide.setdefault("subtitle", []).append(value)
                    elif key == "Footer":
                        slide["footer"] = value
                continue

            current_block = None
            if key == "Layout":
                slide["layout"] = value
                slide["_layout_explicit"] = True
            elif key == "Ghost":
                slide["ghost"] = value
            elif key == "Tag":
                slide["tag"] = value
            elif key == "Footer":
                slide["footer"] = value
            elif key == "Title":
                pass
            else:
                slide[key.lower().replace(" ", "_")] = value
            continue

        current_block = None

    if not slide["title"]:
        if auto_style_titles:
            slide["title"] = stylize_title(slide_name)
        else:
            raise ValueError(f"Slide {slide_index} is missing Title block")

    infer_layout_from_body(slide)
    slide.pop("_layout_explicit", None)
    return slide


def build_cover_slide(deck_title: str, tag_prefix: str, canvas: str) -> dict:
    chips = [
        {"text": "Markdown 直出", "color": "ORANGE"},
        {"text": "Cyberpunk", "color": "CYAN"},
        {"text": "Editable", "color": "PINK"},
    ]
    if canvas == "xhs-vertical":
        chips = chips[:2]
    return {
        "tag": f"{tag_prefix} 01".strip(),
        "layout": "cover",
        "ghost": cleanup_title_source(deck_title)[:8].upper(),
        "title": stylize_title(deck_title),
        "subtitle": ["从 Markdown 一键点亮。", "封面 结构 输出一次完成。"],
        "chips": chips,
        "cards": [{"title": "入口", "accent": "PINK", "lines": ["Markdown 大纲", "自动排版", "同风格整套"]}],
    }


def build_ending_slide(deck_title: str, tag_prefix: str) -> dict:
    return {
        "tag": f"{tag_prefix} END".strip(),
        "layout": "ending",
        "ghost": "GLOW",
        "title": [
            {"text": "让内容发光", "color": "WHITE", "size": 132},
            {"text": "让工作流接满", "color": "CYAN", "size": 106},
        ],
        "subtitle": [f"{deck_title} 已转成赛博整套。", "下一步只改内容，不再重做风格。"],
        "footer": "CYBERPUNK PPT / MARKDOWN AUTO DECK",
    }


def normalize_tags(slides: list[dict], tag_prefix: str) -> None:
    for idx, slide in enumerate(slides, 1):
        if not slide.get("tag"):
            slide["tag"] = f"{tag_prefix} {idx:02d}"


def parse_markdown_outline(text: str) -> dict:
    lines = text.splitlines()
    global_lines: list[str] = []
    slide_sections: list[tuple[str, list[str]]] = []

    deck_title = "Cyberpunk Deck"
    current_name = None
    current_lines: list[str] = []
    for line in lines:
        if line.startswith("# ") and current_name is None:
            deck_title = line[2:].strip() or deck_title
        if line.startswith("## "):
            if current_name is not None:
                slide_sections.append((current_name, current_lines))
            current_name = line[3:].strip()
            current_lines = []
        elif current_name is None:
            global_lines.append(line)
        else:
            current_lines.append(line)
    if current_name is not None:
        slide_sections.append((current_name, current_lines))

    tag_prefix = "CYBER / CUT"
    default_layout = "poster_cards"
    auto_style_titles = True
    canvas = "widescreen"
    batch_deck = False
    for line in global_lines:
        stripped = line.strip()
        if stripped.startswith("Tag Prefix:"):
            tag_prefix = stripped.split(":", 1)[1].strip()
        elif stripped.startswith("Default Layout:"):
            default_layout = stripped.split(":", 1)[1].strip()
        elif stripped.startswith("Auto Style Titles:"):
            auto_style_titles = is_enabled(stripped.split(":", 1)[1].strip())
        elif stripped.startswith("Canvas:"):
            canvas = stripped.split(":", 1)[1].strip()
        elif stripped.startswith("Batch Deck:"):
            batch_deck = is_enabled(stripped.split(":", 1)[1].strip())

    slides = []
    for idx, (name, section_lines) in enumerate(slide_sections, 1):
        slide = parse_slide_blocks(section_lines, idx, name, tag_prefix, default_layout, auto_style_titles)
        slide.setdefault("ghost", name[:12].upper())
        slides.append(slide)

    if batch_deck and slides:
        slides = [build_cover_slide(deck_title, tag_prefix, canvas)] + slides + [build_ending_slide(deck_title, tag_prefix)]
    normalize_tags(slides, tag_prefix)
    return {"canvas": canvas, "deck_title": deck_title, "slides": slides}


def write_spec(spec: dict, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert markdown outline to cyberpunk PPT JSON spec.")
    parser.add_argument("--input", required=True, help="Markdown outline file.")
    parser.add_argument("--output", help="Output JSON spec file. Omit to auto-organize under ~/ai-gen-ppt/.")
    parser.add_argument("--pptx-output", help="Optional PPTX output path.")
    parser.add_argument("--pdf-output", help="Optional PDF output path.")
    parser.add_argument("--png-dir", help="Optional directory for PNG slide exports.")
    parser.add_argument("--assets-dir", help="Optional directory for generated background assets.")
    args = parser.parse_args()

    input_path = Path(args.input)
    spec = parse_markdown_outline(input_path.read_text(encoding="utf-8"))

    if args.output:
        output_path = Path(args.output)
    else:
        deck_title = extract_deck_title(spec)
        out_dir = resolve_output_dir(deck_title)
        output_path = out_dir / "spec.json"

    write_spec(spec, output_path)

    if args.assets_dir:
        assets_dir = Path(args.assets_dir)
    elif args.output:
        assets_dir = output_path.parent / "generated_cyberpunk_assets"
    else:
        assets_dir = output_path.parent / "assets"

    if args.output:
        pptx_output = Path(args.pptx_output) if args.pptx_output else None
        pdf_output = Path(args.pdf_output) if args.pdf_output else None
        png_dir = Path(args.png_dir) if args.png_dir else None
    else:
        deck_title = extract_deck_title(spec)
        safe_title = sanitize_dirname(deck_title)
        out_dir = output_path.parent
        pptx_output = out_dir / f"{safe_title}.pptx" if args.pptx_output is None or args.pptx_output == "" else Path(args.pptx_output)
        if pptx_output is None:
            pptx_output = out_dir / f"{safe_title}.pptx"
        pdf_output = out_dir / f"{safe_title}.pdf" if args.pdf_output else None
        png_dir = out_dir / "png" if args.png_dir is None or args.png_dir == "" else None
        if png_dir is None and args.png_dir:
            png_dir = Path(args.png_dir)

    if pptx_output:
        make_presentation(spec, pptx_output, assets_dir)
        if pdf_output:
            export_pdf(pptx_output, pdf_output)
    elif pdf_output:
        with tempfile.TemporaryDirectory(prefix="cyberpunk-markdown-") as tmpdir:
            temp_ppt = Path(tmpdir) / "outline.pptx"
            make_presentation(spec, temp_ppt, assets_dir)
            export_pdf(temp_ppt, pdf_output)

    if png_dir:
        keep_pptx = str(pptx_output) if pptx_output else None
        export_images(output_path, png_dir, assets_dir=assets_dir, pptx_path=Path(keep_pptx) if keep_pptx else None)

    print(f"Wrote {output_path} with {len(spec['slides'])} slides")


if __name__ == "__main__":
    main()
