# Web Page to PDF Converter

A Python script that reads URLs from `pages.txt`, downloads web pages with images and resources, and creates a single PDF file.

## Requirements

- Python 3.6+
- `requests` - for downloading web pages
- `beautifulsoup4` - for HTML parsing
- `pdfkit` - for PDF generation
- `pillow` - for image handling
- `wkhtmltopdf` - command line tool for PDF conversion (system dependency)

## Installation

```bash
pip install requests beautifulsoup4 pdfkit pillow

# Install wkhtmltopdf:
# Ubuntu/Debian: sudo apt-get install wkhtmltopdf
# macOS: brew install wkhtmltopdf
# Windows: Download from https://wkhtmltopdf.org/downloads.html
```

## Usage

1. Create a `pages.txt` file with one URL per line:
```
https://example.com
https://example.org
https://example.net
```

2. Run the script:
```bash
python page_to_docx.py
```

3. The output will be saved as `output.pdf`

## Features

- Downloads complete web pages with images and CSS
- Handles relative URLs
- Inlines CSS files for better PDF rendering
- Creates a single PDF with all pages
- Includes proper page breaks between pages
- Cleans up temporary files after completion

## Notes

- Be respectful of websites' terms of service and robots.txt
- Add delays between requests to avoid overloading servers
- Some complex JavaScript-heavy sites may not render perfectly
- Large pages may take time to process