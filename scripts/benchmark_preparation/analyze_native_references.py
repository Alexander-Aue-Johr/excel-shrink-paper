"""Summarize completed native references and separate white extents from changed pixels."""
import argparse
import hashlib
import json
from pathlib import Path

from PIL import Image, ImageChops


def white_exterior(image, width, height):
    strips = []
    if image.width > width:
        strips.append((width, 0, image.width, image.height))
    if image.height > height:
        strips.append((0, height, width, image.height))
    return all(image.crop(box).getextrema() == ((255, 255),) * 4 for box in strips)


def compare_images(directory, region):
    paths = [directory / entry["file"] for entry in region["images"]]
    for path, identity in zip(paths, region["images"]):
        if hashlib.sha256(path.read_bytes()).hexdigest() != identity["sha256"]:
            raise ValueError(f"Reference image identity mismatch: {path}")
    with Image.open(paths[0]) as original, Image.open(paths[1]) as candidate:
        left = original.convert("RGBA")
        right = candidate.convert("RGBA")
    width = min(left.width, right.width)
    height = min(left.height, right.height)
    difference = ImageChops.difference(left.crop((0, 0, width, height)), right.crop((0, 0, width, height)))
    channels = difference.split()
    maximum = channels[0]
    for channel in channels[1:]:
        maximum = ImageChops.lighter(maximum, channel)
    changed = width * height - maximum.histogram()[0]
    same_origin = all(region["original_geometry"][axis] == region["default_geometry"][axis] for axis in ("left", "top"))
    white_margin_only = same_origin and changed == 0 and white_exterior(left, width, height) and white_exterior(right, width, height)
    return {"common_size": [width, height], "changed_common_rgba_pixels": changed,
            "difference_bounds": maximum.getbbox(), "same_origin": same_origin,
            "equal_or_white_margin_only": white_margin_only}


def summarize_reference(path):
    reference = json.loads(path.read_text())
    launcher_path = path.parent.with_name(path.parent.name + ".runtime") / "reference-launcher.json"
    launcher = json.loads(launcher_path.read_text()) if launcher_path.exists() else None
    complete = reference["status"] == "completed" and launcher is not None and launcher["status"] == "completed" and launcher["exit_code"] == 0
    rows = []
    for region in reference["regions"]:
        row = {key: region.get(key) for key in ("sheet", "range", "pixels_compared", "pixels_differ", "geometry_differs", "rendering_differs")}
        if complete and region["pixels_compared"]:
            row.update(compare_images(path.parent, region))
        rows.append(row)
    return {"directory": path.parent.name, "status": reference["status"], "launcher_confirmed_complete": complete,
            "error": reference.get("error"), "plan_sha256": reference["plan_sha256"],
            "original_sha256": reference["original_sha256"], "default_sha256": reference["default_sha256"],
            "regions": rows, "limitations": reference["limitations"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--study-root", type=Path, required=True)
    root = parser.parse_args().study_root / "comparators"
    references = [summarize_reference(path) for path in sorted((root / "diagnostics").glob("*/*/reference.json"))]
    output = root / "native-reference-summary.json"
    output.write_text(json.dumps(references, indent=2) + "\n")
    for reference in references:
        rows = reference["regions"]
        print(json.dumps({"directory": reference["directory"], "complete": reference["launcher_confirmed_complete"],
                          "regions": len(rows), "common_pixel_differences": sum(row.get("changed_common_rgba_pixels", 0) > 0 for row in rows),
                          "white_extent_or_equal": sum(row.get("equal_or_white_margin_only", False) for row in rows)}))


if __name__ == "__main__":
    main()
