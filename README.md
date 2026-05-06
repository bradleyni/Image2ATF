# Image2ATF

Image2ATF is a desktop tool for extracting text from images and generating dual-column ATF-style Word files for translation preparation. It is designed for DTP, localization, OCR review, image text extraction, and Adobe-related DTP workflows.

The main script is:

```text
Image2ATF_V2.8.py
```

## Features

- OCR text extraction from images
- Generate dual-column ATF Word files
- Batch image processing
- Drag-and-drop image loading
- Image preview and OCR result editing
- Manual re-OCR for individual images
- Auto-rotation support
- Mask Mode for selecting specific text areas
- AI-assisted text box detection
- Rich text marker handling:
  - Bold
  - Italic
  - Underline
  - Superscript
  - Subscript
  - Colored text
- Optional Image-to-PSD workflow
- Optional Adobe Photoshop / Illustrator export and import workflows
- PyQt6 desktop interface

## Typical Use Cases

This tool is intended for:

- Image text extraction
- Translation preparation
- ATF file creation
- DTP handoff preparation
- OCR cleanup
- Diagram text extraction
- Localized image preparation
- PSD / AI text export and import workflows

## Supported Input Formats

Common image formats are supported, including:

```text
.jpg
.jpeg
.png
.webp
.gif
.bmp
.tiff
.tif
```

## Output

The tool can generate Word files in ATF-style dual-column format.

Typical output:

```text
.docx
```

The generated Word file contains:

- Source image reference
- Extracted OCR text
- Hidden source text column
- Visible editable text column for translation / review

## Requirements

Recommended environment:

- Python 3.10 or newer
- Windows 10 / 11 is recommended
- Microsoft Word is recommended for reviewing generated `.docx` files
- Internet access for AI API calls
- Valid API key for at least one supported OCR provider

Python packages are listed in:

```text
requirements.txt
```

Main dependencies include:

- PyQt6
- Pillow
- requests
- python-docx
- google-genai
- pywin32

## Optional Adobe Requirements

Some DTP-related features require additional software.

For Photoshop / Illustrator automation:

- Windows is required
- Adobe Photoshop and / or Adobe Illustrator must be installed
- `pywin32` must be installed
- Adobe COM automation must be available on the machine

Without these, the core OCR and ATF generation features can still be used, but Adobe-related export / import functions may not work.

## Installation

Clone or download this repository, then install dependencies.

```bash
pip install -r requirements.txt
```

On Windows, you can also run:

```text
install_windows.bat
```

## Running the Tool

Run the main Python script:

```bash
python Image2ATF_V2.8.py
```

On Windows, you can also use:

```text
run_windows.bat
```

## Basic Workflow

1. Open the application.
2. Select the OCR provider tab.
3. Enter the required API key.
4. Add images or an image folder.
5. Review image orientation.
6. Run OCR.
7. Edit OCR results if needed.
8. Generate the ATF Word file.
9. Review the generated `.docx` before sending it to translation or DTP.

## Supported OCR Providers

The tool supports OCR engines through API calls:

- Gemini

You only need to configure the provider you plan to use.

## API Keys

Do not commit API keys to GitHub.

Recommended:

- Store API keys locally only
- Do not hard-code private keys into the script
- Do not share API keys in screenshots, logs, or sample files
- Do not upload confidential client images to public repositories

## Mask Mode

Mask Mode allows users to manually select specific text areas from an image.

This is useful when:

- The image contains mixed text and graphics
- Only part of the image needs translation
- OCR should ignore irrelevant areas
- The automatic OCR result is too noisy
- Text blocks need to be extracted in a controlled order

## Rich Text Markers

The OCR workflow may preserve formatting using markers such as:

```text
**bold**
*italic*
__underline__
^superscript^
~subscript~
[color:#RRGGBB]colored text[/color]
```

These markers are processed when generating the Word output.

## Repository Structure

```text
.
├── Image2ATF_V2.8.py
├── requirements.txt
├── README.md
├── LICENSE
├── SECURITY.md
├── .gitignore
├── install_windows.bat
├── run_windows.bat
└── docs/
    └── GITHUB_UPLOAD_GUIDE_ZH.md
```

## Security and Privacy

This tool may process client images, screenshots, technical diagrams, and localized assets.

Before uploading anything to GitHub, make sure you do not include:

- API keys
- Internal proxy addresses
- Client images
- Confidential screenshots
- Translation files
- Generated ATF files from client projects
- Logs containing private information
- Adobe project files containing client content

For safety, keep the repository private unless you are certain the code and sample files are safe to publish.

## Limitations

OCR results depend on image quality and model performance. Manual review is required.

Common issues may include:

- Missing small text
- Incorrect reading order
- Incorrect formatting markers
- Mistaken superscript or subscript
- Incorrect color detection
- Text missed in dense diagrams
- Incorrect auto-rotation
- Partial extraction from low-resolution images

## Recommended QA

After OCR and ATF generation, check:

- Text completeness
- Reading order
- Line breaks
- Numbers and symbols
- Units and technical labels
- Superscript and subscript
- Bold / italic / underline formatting
- Colored text
- Image-to-text matching
- Hidden and visible Word columns

## License

All rights reserved unless otherwise stated in the `LICENSE` file.

## Disclaimer

This software is provided as an internal productivity tool. Users are responsible for reviewing all generated output before use, delivery, or publication.