#!/usr/bin/env python3
"""
watermark_pdf.py

Add an image (logo) watermark to the center of every page of a PDF.

Usage:
    python watermark_pdf.py INPUT.pdf LOGO.png OUTPUT.pdf [options]

Options:
    --scale FLOAT       Logo width as a fraction of page width (0.0-1.0).
                         Default: 0.5 (logo takes up 50% of page width)
    --width PT          Exact logo width in points, overrides --scale.
    --opacity FLOAT     Watermark opacity, 0.0 (invisible) - 1.0 (fully opaque).
                         Default: 0.15
    --rotate DEGREES    Rotation angle in degrees (counter-clockwise). Default: 0

Examples:
    # Default: logo at 50% of page width, 15% opacity, centered
    python watermark_pdf.py report.pdf logo.png report_watermarked.pdf

    # Big, faint watermark at 70% of page width
    python watermark_pdf.py report.pdf logo.png out.pdf --scale 0.7 --opacity 0.1

    # Exact width in points, rotated 45 degrees
    python watermark_pdf.py report.pdf logo.png out.pdf --width 200 --rotate 45
"""

import argparse
import io
import sys

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image


def make_watermark_page(logo_path, page_width, page_height, scale, width_pt, opacity, rotate):
    """Create a single-page PDF (in-memory) containing the logo centered on
    a page of the given dimensions, at the requested size/opacity/rotation."""

    # Load logo to get its native aspect ratio
    pil_img = Image.open(logo_path)
    img_w, img_h = pil_img.size
    aspect = img_h / img_w

    # Determine target width in points
    if width_pt is not None:
        target_w = width_pt
    else:
        target_w = page_width * scale

    target_h = target_w * aspect

    # Apply opacity by adjusting the image's alpha channel
    if pil_img.mode != "RGBA":
        pil_img = pil_img.convert("RGBA")

    if opacity < 1.0:
        r, g, b, a = pil_img.split()
        a = a.point(lambda px: int(px * opacity))
        pil_img = Image.merge("RGBA", (r, g, b, a))

    img_buffer = io.BytesIO()
    pil_img.save(img_buffer, format="PNG")
    img_buffer.seek(0)
    img_reader = ImageReader(img_buffer)

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_width, page_height))

    c.saveState()
    # Move origin to page center, rotate, then draw image centered on that origin
    c.translate(page_width / 2, page_height / 2)
    if rotate:
        c.rotate(rotate)
    c.drawImage(
        img_reader,
        -target_w / 2,
        -target_h / 2,
        width=target_w,
        height=target_h,
        mask="auto",
    )
    c.restoreState()
    c.save()

    buf.seek(0)
    return PdfReader(buf).pages[0]


def add_watermark(input_pdf, logo_path, output_pdf, scale=0.5, width_pt=None,
                   opacity=0.15, rotate=0):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # Cache watermark pages by (width, height) since most PDFs have uniform
    # page sizes, but handle mixed sizes correctly if they occur.
    watermark_cache = {}

    for page in reader.pages:
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        key = (round(pw, 2), round(ph, 2))

        if key not in watermark_cache:
            watermark_cache[key] = make_watermark_page(
                logo_path, pw, ph, scale, width_pt, opacity, rotate
            )

        page.merge_page(watermark_cache[key])
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    print(f"Watermarked PDF saved to: {output_pdf}")
    print(f"  Pages processed: {len(reader.pages)}")
    print(f"  Logo scale: {'{}pt width'.format(width_pt) if width_pt else f'{scale*100:.0f}% of page width'}")
    print(f"  Opacity: {opacity}")
    if rotate:
        print(f"  Rotation: {rotate}°")


def main():
    parser = argparse.ArgumentParser(
        description="Add a centered logo watermark to every page of a PDF.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input_pdf", help="Path to the source PDF")
    parser.add_argument("logo", help="Path to the logo image (PNG/JPG)")
    parser.add_argument("output_pdf", help="Path to save the watermarked PDF")
    parser.add_argument(
        "--scale", type=float, default=0.5,
        help="Logo width as fraction of page width (0.0-1.0). Default: 0.5"
    )
    parser.add_argument(
        "--width", type=float, default=None, dest="width_pt",
        help="Exact logo width in points (overrides --scale)"
    )
    parser.add_argument(
        "--opacity", type=float, default=0.15,
        help="Watermark opacity, 0.0-1.0. Default: 0.15"
    )
    parser.add_argument(
        "--rotate", type=float, default=0,
        help="Rotation angle in degrees. Default: 0"
    )

    args = parser.parse_args()

    if not (0.0 <= args.opacity <= 1.0):
        parser.error("--opacity must be between 0.0 and 1.0")
    if args.width_pt is None and not (0.0 < args.scale <= 1.0):
        parser.error("--scale must be between 0.0 (exclusive) and 1.0")

    add_watermark(
        args.input_pdf,
        args.logo,
        args.output_pdf,
        scale=args.scale,
        width_pt=args.width_pt,
        opacity=args.opacity,
        rotate=args.rotate,
    )


if __name__ == "__main__":
    main()
