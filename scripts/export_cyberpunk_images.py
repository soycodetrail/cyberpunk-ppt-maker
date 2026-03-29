#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import subprocess
import tempfile

from generate_cyberpunk_ppt import export_pdf, load_spec, make_presentation


def export_images(spec_path: Path, output_dir: Path, assets_dir: Path | None = None, pptx_path: Path | None = None) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    if assets_dir is None:
        assets_dir = output_dir / "assets"

    with tempfile.TemporaryDirectory(prefix="cyberpunk-images-") as tmpdir:
        tmpdir_path = Path(tmpdir)
        working_ppt = pptx_path or (tmpdir_path / "slides.pptx")
        working_pdf = tmpdir_path / "slides.pdf"
        spec = load_spec(spec_path)
        make_presentation(spec, working_ppt, assets_dir)
        export_pdf(working_ppt, working_pdf)

        prefix = tmpdir_path / "page"
        subprocess.run(
            ["pdftoppm", "-png", str(working_pdf), str(prefix)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        for idx, png in enumerate(sorted(tmpdir_path.glob("page-*.png")), 1):
            shutil.copy2(png, output_dir / f"slide_{idx:02d}.png")


def main() -> None:
    parser = argparse.ArgumentParser(description="Export cyberpunk slide images from a JSON spec.")
    parser.add_argument("--spec", required=True, help="Path to JSON spec file.")
    parser.add_argument("--output-dir", required=True, help="Directory for exported PNG slide images.")
    parser.add_argument("--assets-dir", help="Optional directory for generated background assets.")
    parser.add_argument("--keep-pptx", help="Optional path to keep the intermediate PPTX.")
    args = parser.parse_args()

    export_images(
        spec_path=Path(args.spec),
        output_dir=Path(args.output_dir),
        assets_dir=Path(args.assets_dir) if args.assets_dir else None,
        pptx_path=Path(args.keep_pptx) if args.keep_pptx else None,
    )


if __name__ == "__main__":
    main()
