import os
import shutil
from PIL import Image, ImageDraw
from main import compress_img


ROOT = os.path.dirname(__file__)
SCENARIO_DIR = os.path.join(ROOT, "scenario_demo")
BEFORE_DIR = os.path.join(SCENARIO_DIR, "before")
AFTER_DIR = os.path.join(SCENARIO_DIR, "after")


def ensure_dirs():
    os.makedirs(BEFORE_DIR, exist_ok=True)
    os.makedirs(AFTER_DIR, exist_ok=True)


def print_sizes(folder, label=""):
    print(f"\n{label} sizes:")
    for fn in sorted(os.listdir(folder)):
        path = os.path.join(folder, fn)
        if os.path.isfile(path):
            size = os.path.getsize(path)
            print(f" - {fn}: {size/1024:.1f} KB")


def run_demo():
    ensure_dirs()

    print_sizes(BEFORE_DIR, "Before compression (originals)")

    # Compress each file in BEFORE_DIR, then move the compressed file to AFTER_DIR
    for fn in sorted(os.listdir(BEFORE_DIR)):
        src = os.path.join(BEFORE_DIR, fn)
        if not os.path.isfile(src):
            continue
        name, ext = os.path.splitext(fn)
        # choose whether to convert to jpg: keep png as png, keep jpg as jpg
        to_jpg = False if ext.lower() == ".png" else True
        # call compress_img; use a resize ratio to make a visible difference
        # compress_img returns the path to the generated compressed file
        compressed_path = compress_img(src, new_size_ratio=0.6, quality=85, to_jpg=to_jpg)

        # Move/rename compressed file into AFTER_DIR with the SAME filename as original
        dest = os.path.join(AFTER_DIR, fn)
        # If the compressed file has a different extension (e.g., converted to .jpg), adjust dest
        comp_name, comp_ext = os.path.splitext(os.path.basename(compressed_path))
        if comp_ext.lower() != ext.lower():
            # keep same basename but use compressed file's extension
            dest = os.path.join(AFTER_DIR, f"{name}{comp_ext}")

        # If destination exists, overwrite
        if os.path.exists(dest):
            os.remove(dest)
        shutil.move(compressed_path, dest)
        print(f"Compressed {fn} -> {os.path.basename(dest)}")

    print_sizes(AFTER_DIR, "After compression (same filenames in different folder)")


if __name__ == "__main__":
    run_demo()
