#!/usr/bin/env python3

import subprocess
import sys

import yaml
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor


def ensure_built_in_styles_exist(doc, style_names):
    """
    Add a temporary paragraph for each built-in style so that
    those styles are not 'latent' (some built-in styles can be 
    hidden until they are first used).
    """
    for style_name in style_names:
        doc.add_paragraph("X", style=style_name)
    return doc


def remove_temp_paragraphs(doc):
    """Remove the dummy paragraphs that contain only 'X'."""
    to_remove = []
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "X":
            to_remove.append(i)
    for idx in reversed(to_remove):
        p = doc.paragraphs[idx]
        p._element.getparent().remove(p._element)


def override_built_in_style(doc, base_name, custom_name,
                            font_size, bold, italic,
                            font_color, space_before, space_after):
    """
    Override a built-in style in `doc.styles[base_name]`.
    We rename it to `custom_name` visually, but keep the style ID
    so Pandoc will still recognize it for e.g. "Heading 1".
    """
    try:
        style = doc.styles[base_name]
    except KeyError:
        print(f"[WARNING] style '{base_name}' not found. Skipping.")
        return

    # Make sure it's a paragraph style
    if style.type != WD_STYLE_TYPE.PARAGRAPH:
        print(f"[WARNING] style '{base_name}' is not a paragraph style. Skipping.")
        return

    # Save the style_id so we don't break Pandoc
    original_id = style.style_id

    # Rename for visual clarity in Word (e.g. "IBM Heading 1"),
    # but keep the same underlying style_id
    style.name = custom_name
    style.style_id = original_id

    font = style.font
    font.name = "IBM Plex Sans"  # or "IBM Plex Serif", etc.
    font.size = Pt(font_size)
    font.bold = bold
    font.italic = italic

    r, g, b = font_color
    font.color.rgb = RGBColor(r, g, b)

    pf = style.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)


def create_reference_doc(config_file, reference_docx):
    """
    1) Create a fresh doc
    2) Force built-in styles to appear
    3) Override each built-in style from the YAML data
    4) Remove dummy paragraphs
    5) Save as reference.docx
    """
    doc = Document()

    # We might glean the needed built-in styles from the YAML
    # but here we just ensure that the typical set (Normal, Title, etc.)
    # definitely appear in the doc so we can override them.
    needed_styles = [
        "Normal", "Title", "Subtitle",
        "Heading 1", "Heading 2", "Heading 3",
        "Heading 4", "Heading 5", "Heading 6",
    ]
    ensure_built_in_styles_exist(doc, needed_styles)

    # 2) Load config from YAML
    with open(config_file, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)

    # 3) For each style definition in YAML, override the built-in style
    for style_def in data.get("styles", []):
        base_name = style_def["base_name"]      # e.g. "Heading 1"
        custom_name = style_def["custom_name"]  # e.g. "IBM Heading 1"

        font_size = style_def["font_size"]
        bold = style_def.get("bold", False)
        italic = style_def.get("italic", False)
        font_color = tuple(style_def.get("font_color", [0, 0, 0]))
        space_before = style_def.get("space_before", 0)
        space_after = style_def.get("space_after", 0)

        override_built_in_style(
            doc,
            base_name=base_name,
            custom_name=custom_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            font_color=font_color,
            space_before=space_before,
            space_after=space_after
        )

    # 4) Remove dummy paragraphs
    remove_temp_paragraphs(doc)

    # 5) Save
    doc.save(reference_docx)
    print(f"[INFO] Created reference doc: {reference_docx}")


def convert_md_to_word(input_md, output_docx, reference_docx):
    """
    Runs Pandoc with --reference-doc=reference_docx to generate final DOCX.
    """
    cmd = [
        "pandoc",
        input_md,
        "--reference-doc=" + reference_docx,
        "-o",
        output_docx
    ]
    print(f"[INFO] Running: {' '.join(cmd)}")
    subprocess.run(cmd, check=True)
    print(f"[INFO] Converted {input_md} -> {output_docx} using {reference_docx}")


def main():
    """
    Usage:
       python format_doc.py <yaml_config> <input_md> <output_docx>
    """
    if len(sys.argv) < 4:
        print("Usage: format_doc.py <yaml_config> <input_md> <output_docx>")
        sys.exit(1)

    config_file = sys.argv[1]
    input_md = sys.argv[2]
    output_docx = sys.argv[3]

    reference_doc = "reference.docx"

    # Create the reference doc with your custom IBM styles
    create_reference_doc(config_file, reference_doc)

    # Convert the Markdown to DOCX with Pandoc using that reference
    convert_md_to_word(input_md, output_docx, reference_doc)


if __name__ == "__main__":
    main()
