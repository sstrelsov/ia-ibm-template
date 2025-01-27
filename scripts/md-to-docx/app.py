#!/usr/bin/env python3

import os
import sys

import yaml
from doc_utils.overrides import convert_md_to_word, create_reference_doc
from doc_utils.table_style import apply_table_style


def main():
    """
    Usage:
       python app.py <input_md> <output_docx>
    """
    if len(sys.argv) < 3:
        print("Usage: python app.py <input_md> <output_docx>")
        sys.exit(1)

    config_file = './config.yaml'
    input_md = sys.argv[1]
    output_docx = sys.argv[2]

    # Load the config YAML
    with open(config_file, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    reference_docx = "reference.docx"

    try:
        # 1) Create reference doc with styles
        create_reference_doc(config, reference_docx)

        # 2) Convert from Markdown to Word using Pandoc
        convert_md_to_word(
            input_md,
            output_docx,
            reference_docx,
            from_format=config.get("pandoc_options", {}).get("from", "markdown+footnotes+mark")
        )

        # 3) Apply table style (if you want to style tables in the final doc)
        #    If 'pandoc_options.table_style' not set, default to "Light Shading".
        table_style = config.get("pandoc_options", {}).get("table_style", "Light Shading")
        apply_table_style(output_docx, table_style, output_docx)

        print(f"[INFO] Finished creating {output_docx} with reference doc {reference_docx}")
    finally:
        # Remove the reference doc to avoid clutter
        if os.path.exists(reference_docx):
            os.remove(reference_docx)
            print(f"[INFO] Deleted the reference DOCX: {reference_docx}")


if __name__ == "__main__":
    main()
