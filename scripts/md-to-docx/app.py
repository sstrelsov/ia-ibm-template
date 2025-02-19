#!/usr/bin/env python3

import os
import subprocess
import sys
import yaml
from dotenv import load_dotenv
import questionary

from doc_utils.overrides import convert_md_to_word, create_reference_doc
from doc_utils.table_style import apply_table_style

# Load environment variables from .env
load_dotenv()

def open_document(filepath):
    """
    Open the document using the default application based on the operating system.
    """
    if sys.platform.startswith('win'):
        os.startfile(filepath)
    elif sys.platform.startswith('darwin'):
        subprocess.call(['open', filepath])
    else:
        subprocess.call(['xdg-open', filepath])

def select_markdown_file(directory):
    """
    Uses Questionary to select an item (a markdown file or a folder) from the provided directory.
    If a folder is selected, lists markdown files within that folder for a secondary choice.
    Returns the full path to the chosen markdown file.
    """
    try:
        items = sorted(os.listdir(directory))
    except FileNotFoundError:
        questionary.print(f"Directory not found: {directory}", style="bold fg:red")
        sys.exit(1)

    if not items:
        questionary.print(f"No items found in directory: {directory}", style="bold fg:red")
        sys.exit(1)

    choice = questionary.select(
        "Select a file or folder:",
        choices=items
    ).ask()

    if not choice:
        sys.exit(1)

    selected_item = os.path.join(directory, choice)

    if os.path.isfile(selected_item):
        if selected_item.lower().endswith(".md"):
            return selected_item
        else:
            questionary.print("Selected file is not a markdown (.md) file.", style="bold fg:red")
            sys.exit(1)
    elif os.path.isdir(selected_item):
        # List markdown files within the selected folder
        md_files = [f for f in sorted(os.listdir(selected_item)) if f.lower().endswith(".md")]
        if not md_files:
            questionary.print("No markdown files found in the selected folder.", style="bold fg:red")
            sys.exit(1)
        md_choice = questionary.select(
            "Select a markdown file from the folder:",
            choices=md_files
        ).ask()
        if not md_choice:
            sys.exit(1)
        return os.path.join(selected_item, md_choice)
    else:
        questionary.print("The selected item is neither a file nor a directory.", style="bold fg:red")
        sys.exit(1)

def main():
    # Get markdown directory from .env (or use default)
    markdown_dir = os.environ.get("MARKDOWN_DIR", "/Users/a6127238/markdown")
    questionary.print(f"Markdown directory: {markdown_dir}", style="bold")

    # Interactive selection of input markdown file
    input_md = select_markdown_file(markdown_dir)
    questionary.print(f"Selected markdown file: {input_md}", style="fg:green")

    # Prompt for output document name (without extension)
    output_name = questionary.text(
        "Enter the output document name (without .docx extension):"
    ).ask()

    if not output_name:
        questionary.print("Output document name cannot be empty.", style="bold fg:red")
        sys.exit(1)
    output_docx = f"{output_name}.docx"

    # Load the configuration YAML
    config_file = './config.yaml'
    try:
        with open(config_file, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
    except FileNotFoundError:
        questionary.print(f"Configuration file not found: {config_file}", style="bold fg:red")
        sys.exit(1)

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

        # 3) Apply table style (if applicable)
        table_style = config.get("pandoc_options", {}).get("table_style", "Light Shading")
        apply_table_style(output_docx, table_style, output_docx)

        questionary.print(f"\n[INFO] Finished creating {output_docx} with reference doc {reference_docx}", style="bold fg:green")
    finally:
        # Clean up: remove the reference document to avoid clutter
        if os.path.exists(reference_docx):
            os.remove(reference_docx)
            questionary.print(f"[INFO] Deleted the reference DOCX: {reference_docx}", style="fg:yellow")

    # Automatically open the generated Word document
    open_document(output_docx)

if __name__ == "__main__":
    main()
