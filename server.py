#!/usr/bin/env python3

"""
Simple HTTP server for uploading a PDF invoice, extracting product information,
and providing the extracted data as an Excel file download.

This server uses only Python's standard library along with pandas. PDF text
extraction is performed via the external `pdftotext` command from Poppler,
which is available in the environment. No external Python packages are required.

Run this script, then navigate to http://localhost:8000/ in your browser.
"""

import io
import os
import re
import tempfile
from pdfminer.high_level import extract_text
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import parse_qs
import cgi
import pandas as pd


def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Convert a PDF to text using the ``pdfminer.six`` library.

    This implementation avoids reliance on external command-line utilities like
    ``pdftotext``. If extraction fails for any reason, an empty string is
    returned.
    """
    try:
        return extract_text(pdf_path)
    except Exception:
        return ''


def parse_products_from_text(text: str) -> list[dict]:
    """
    Parse product information from extracted text using a two-step approach.

    1. Identify description lines that describe windows or products. These lines
       contain a four-digit series ID, a window style description, and raw
       dimensions separated by an ``x`` (e.g., ``7300 Series Direct Set Half Round 73 x 36 1/2 - FLANGE``).
    2. Pair each description line with the next ``Frame:`` line, which contains
       numeric width and height values and the frame colour. The frame colour is
       assumed to be the final comma-separated token on the line.

    This logic is based on the sample invoice provided by the user and should
    generalise to similar invoices that follow this structure.

    Returns a list of dictionaries with columns matching the sample output:
    Manufacturer, Frame Color, Series ID, Window Style, Width, Height.
    """
    lines = [line.strip() for line in text.split('\n')]
    products: list[dict] = []
    pending: dict | None = None
    desc_pattern = re.compile(
        r'^(\d{4})\s+Series\s+(.+?)\s+([\d\s/]+)\s*x\s*([\d\s/]+)\s*-',
        re.IGNORECASE
    )
    for line in lines:
        if not line:
            continue
        m = desc_pattern.match(line)
        if m:
            pending = {
                'Series ID': m.group(1).strip() + ' Series',
                'Window Style': m.group(2).strip(),
            }
            continue
        if pending and line.startswith('Frame:'):
            w_match = re.search(r'Width =\s*([\d\.]+)', line)
            h_match = re.search(r'Height =\s*([\d\.]+)', line)
            width = float(w_match.group(1)) if w_match else None
            height = float(h_match.group(1)) if h_match else None
            colour = line.split(',')[-1].strip()
            products.append({
                'Manufacturer': 'CWS',
                'Frame Color': colour,
                'Series ID': pending['Series ID'],
                'Window Style': pending['Window Style'],
                'Width': width,
                'Height': height
            })
            pending = None
    return products


class InvoiceHandler(BaseHTTPRequestHandler):
    def _send_html(self, content, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        if isinstance(content, str):
            content = content.encode('utf-8')
        self.wfile.write(content)

    def do_GET(self):
        # Serve the upload form for the root path
        if self.path == '/' or self.path.startswith('/?'):
            html = """
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Invoice Product Extractor</title>
                <style>
                    body { font-family: Arial, sans-serif; background: #f5f5f5; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
                    .container { background: #fff; padding: 20px 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
                    h1 { margin-top: 0; }
                    input[type=file] { margin: 10px 0; }
                    button { padding: 10px 20px; background-color: #007BFF; color: #fff; border: none; border-radius: 4px; cursor: pointer; }
                    button:hover { background-color: #0056b3; }
                    .error { color: red; margin-top: 10px; }
                </style>
            </head>
            <body>
                <div class="container">
                    <h1>Invoice Product Extractor</h1>
                    <form action="/" method="post" enctype="multipart/form-data">
                        <input type="file" name="file" accept="application/pdf" required><br>
                        <button type="submit">Upload and Extract</button>
                    </form>
                </div>
            </body>
            </html>
            """
            self._send_html(html)
        else:
            self.send_error(404, "File not found")

    def do_POST(self):
        # Handle file upload and extraction
        content_type = self.headers.get('Content-Type')
        if not content_type:
            self.send_error(400, "Content-Type header missing")
            return
        # Use cgi.FieldStorage to parse form data
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={'REQUEST_METHOD': 'POST', 'CONTENT_TYPE': content_type},
        )
        if 'file' not in form:
            self._send_html("<h3>No file part in the request.</h3>", status=400)
            return
        file_field = form['file']
        if not file_field.filename:
            self._send_html("<h3>No file selected.</h3>", status=400)
            return
        filename = os.path.basename(file_field.filename)
        # Only allow pdf
        if not filename.lower().endswith('.pdf'):
            self._send_html("<h3>Only PDF files are allowed.</h3>", status=400)
            return
        # Save uploaded file to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
            tmp_pdf.write(file_field.file.read())
            tmp_pdf_path = tmp_pdf.name

        # Extract text from the PDF
        text = extract_text_from_pdf(tmp_pdf_path)
        os.unlink(tmp_pdf_path)  # remove the temporary pdf
        if not text.strip():
            self._send_html("<h3>Failed to extract text from PDF.</h3>", status=500)
            return
        # Parse products
        products = parse_products_from_text(text)
        if not products:
            self._send_html("<h3>No product information found in the invoice.</h3>", status=200)
            return
        # Create Excel file in memory
        df = pd.DataFrame(products)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Products')
            # No explicit save() call is required; the context manager will save when closed.
        output.seek(0)
        # Send response headers for file download
        self.send_response(200)
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition', 'attachment; filename="extracted_products.xlsx"')
        self.end_headers()
        self.wfile.write(output.getvalue())


def run(server_class=HTTPServer, handler_class=InvoiceHandler, port=8000):
    server_address = ('', port)
    httpd = server_class(server_address, handler_class)
    print(f"Starting server on port {port}...")
    httpd.serve_forever()


if __name__ == '__main__':
    run()