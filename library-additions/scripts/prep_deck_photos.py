#!/usr/bin/env python
"""
prep_deck_photos.py -- prepare owner-supplied property photos for a bov-deck
(or OM) build: strip the MLS/board copyright strip, downsize, re-encode.

WHY
    Owner-supplied interiors usually come straight out of the MLS and carry a
    "Copyright <year> <Board> of Realtors(R)" watermark baked into the bottom
    edge. Cropping the bottom 4-5% removes it cleanly on standard 3:2 MLS
    exports without visibly changing the composition. Drone/aerial shots are
    normally clean and need no crop, only downsizing -- a 4K JPEG used as a
    full-bleed background is the main reason a rendered deck balloons past
    10 MB.

INPUTS
    One or more IMAGE specs of the form

        <src>[:<dest>][:<crop_bottom_fraction>][:<max_width_px>]

    e.g.  inbox/DJI_0069.jpg:photo_aerial.jpg:0:2400
          inbox/Kitchen-3.jpg:photo_kitchen.jpg:0.045:1600

    Defaults when a field is omitted: dest = source filename,
    crop_bottom = 0.045 (MLS strip), max_width = 1600.

    --outdir  destination directory (default: ./assets, i.e. the bov-deck
              working assets/ folder the HTML references)

OUTPUT
    Re-encoded JPEGs at quality 88, optimized, written to --outdir. Prints the
    before/after dimensions of each so the crop can be sanity-checked.

USAGE
    python prep_deck_photos.py --outdir work/assets \
        "inbox/DJI_0069.jpg:photo_aerial.jpg:0:2400" \
        "inbox/Kitchen-3.jpg:photo_kitchen.jpg" \
        "inbox/br-1.jpg:photo_bedroom.jpg"

NOTES
    - Always LOOK at the cropped output before shipping. If the watermark sits
      higher than the default 4.5%, raise the fraction; if a photo has no
      watermark, pass 0.
    - Real photos replace the template's `.ph` placeholder frames entirely --
      delete those divs from the HTML rather than layering over them, or the
      "Property Photo Placeholder" label can peek out from behind a card.
    - Requires Pillow (`pip install pillow`).
"""
import argparse
import os

from PIL import Image

DEFAULT_CROP = 0.045
DEFAULT_MAXW = 1600


def parse_spec(spec):
    parts = spec.split(":")
    # Guard against Windows drive letters ("C:\path\x.jpg") being split.
    if len(parts) > 1 and len(parts[0]) == 1 and parts[1][:1] in ("\\", "/"):
        parts = [parts[0] + ":" + parts[1]] + parts[2:]
    src = parts[0]
    dest = parts[1] if len(parts) > 1 and parts[1] else os.path.basename(src)
    crop = float(parts[2]) if len(parts) > 2 and parts[2] else DEFAULT_CROP
    maxw = int(parts[3]) if len(parts) > 3 and parts[3] else DEFAULT_MAXW
    return src, dest, crop, maxw


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("images", nargs="+",
                    help="src[:dest][:crop_bottom_fraction][:max_width_px]")
    ap.add_argument("--outdir", default="assets")
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    for spec in args.images:
        src, dest, crop, maxw = parse_spec(spec)
        im = Image.open(src).convert("RGB")
        w0, h0 = im.size
        if crop:
            im = im.crop((0, 0, w0, int(h0 * (1 - crop))))
        w, h = im.size
        if w > maxw:
            im = im.resize((maxw, int(h * maxw / w)), Image.LANCZOS)
        out = os.path.join(args.outdir, dest)
        im.save(out, quality=88, optimize=True)
        print(f"{os.path.basename(src)} {w0}x{h0} -> {dest} {im.size[0]}x{im.size[1]}"
              f"  (crop_bottom={crop}, max_width={maxw})")


if __name__ == "__main__":
    main()
