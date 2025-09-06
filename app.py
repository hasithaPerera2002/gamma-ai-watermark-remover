import os
from watermark_remover import PPTXWatermarkRemover

def remove_pptx_watermark(input_path: str, output_path: str):
    remover = PPTXWatermarkRemover()
    print(f"Processing {input_path}...")

    shapes_removed, links_removed, corner_removed = remover.clean_pptx_from_target_domain(
        input_path, output_path
    )

    total_removed = shapes_removed + links_removed + corner_removed
    print(f"✅ Done! Saved to {output_path}")
    print(f"Details: shapes={shapes_removed}, links={links_removed}, corner_pictures={corner_removed}, total={total_removed}")

if __name__ == "__main__":
    input_file = "COCO-SCAN.pptx"           # your file
    output_file = "COCO-SCAN-cleaned.pptx"  # output filename
    remove_pptx_watermark(input_file, output_file)