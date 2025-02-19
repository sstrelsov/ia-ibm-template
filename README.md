# iA Writer Markdown Template

Inspired by iA Writer's own example GitHub template. Uses IBM Plex Sans as the base font.

## Installation

To install onto iA writer on desktop:

1. Open iA Writer
2. `cmd+,` to open preferences
3. Nav to "Templates" and click the (+) in "Custom Templates"
4. Drop in the file with extension `.iatemplate`

After making changes locally, you must re-add the template. Be sure to cmd+shift+r to reload any lingering styles!

To enable devtools on iA Writer previews, enter the below command in a terminal:

```zsh
defaults write pro.writer.mac WebKitDeveloperExtras -bool true
```

## Converting Markdown -> DOCX

Python script for converting markdown to docx, with specified formatting. If you export directly from iA Writer, the styles in Word are wack. Running the md file through this script, however, leads to a beautiful docx file:

### Instructions

1. Specify a dir containing markdown files in a .env file as `MARKDOWN_DIR`
2. Run `python app.py` and follow the multiple choice prompts
3. The docx should automatically open

In the `/scripts` dir, `styles.yaml` specify the styles for the word doc. Included is `test.md` to demonstrate what the output looks like
