from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
import pandas as pd
import io
import re
from pdfminer.high_level import extract_text

app = Flask(__name__)

# Configure upload folder and allowed extensions
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload size

ALLOWED_EXTENSIONS = {'pdf'}


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_invoice_data(pdf_path: str):
    """
    Extract product information from the invoice PDF.

    This implementation uses the external ``pdftotext`` command to convert
    the uploaded PDF into plain text. It then scans the text for product
    description lines and pairs each with the next encountered frame line to
    retrieve numeric dimensions and frame colour information.

    Lines containing product descriptions are expected to follow a pattern like:

        ``7300 Series Direct Set Half Round 73 x 36 1/2 - FLANGE``

    And a corresponding frame line follows later in the document, e.g.:

        ``Frame: Width = 73, Height = 36.5, Radius = 36.5, Bronze``

    For each product description, the series ID and window style are extracted
    directly from the description line. The width and height are extracted
    from the numerical values on the frame line. The frame colour is assumed
    to be the final comma-separated token on the frame line. All products
    are assumed to have a manufacturer of ``CWS``, based on the sample invoice.

    Args:
        pdf_path (str): Path to the uploaded PDF invoice.

    Returns:
        list[dict]: A list of dictionaries with keys: ``Manufacturer``,
        ``Frame Color``, ``Series ID``, ``Window Style``, ``Width`` and ``Height``.
    """
    # Convert the PDF to plain text using pdfminer.six. This avoids the need for
    # external utilities like ``pdftotext``, making the application easier to
    # deploy on systems without Poppler. If extraction fails, an empty string
    # will be returned and no products will be found.
    try:
        text = extract_text(pdf_path)
    except Exception:
        text = ''
    # Split into lines and strip whitespace
    lines = [line.strip() for line in text.split('\n')]

    products: list[dict] = []
    pending: dict | None = None

    # Pattern to capture series ID, window style, and raw dimensions from the description line
    desc_pattern = re.compile(
        r'^(\d{4})\s+Series\s+(.+?)\s+([\d\s/]+)\s*x\s*([\d\s/]+)\s*-',
        re.IGNORECASE
    )

    for line in lines:
        if not line:
            continue
        match = desc_pattern.match(line)
        if match:
            # Start of a new product description
            pending = {
                'Series ID': match.group(1).strip() + ' Series',
                'Window Style': match.group(2).strip(),
            }
            continue
        # Once a description has been found, look for the next frame line
        if pending and line.startswith('Frame:'):
            # Extract numeric width and height from the frame line
            w_match = re.search(r'Width =\s*([\d\.]+)', line)
            h_match = re.search(r'Height =\s*([\d\.]+)', line)
            width = float(w_match.group(1)) if w_match else None
            height = float(h_match.group(1)) if h_match else None
            # Frame colour is assumed to be the last comma-separated token
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


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if a file part is present
        if 'file' not in request.files:
            return render_template('index.html', error='No file part')
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', error='No selected file')
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            # Extract data from the uploaded PDF
            products = extract_invoice_data(filepath)
            if not products:
                return render_template('index.html', error='No product information found in the PDF.', products=None)

            # Create a DataFrame and Excel file in memory
            df = pd.DataFrame(products)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Products')
                # No explicit save() call is needed; the context manager will save when closed.
            output.seek(0)

            # Provide download link after uploading and processing
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=f"extracted_products.xlsx"
            )
        else:
            return render_template('index.html', error='Allowed file type is PDF')
    return render_template('index.html')


if __name__ == '__main__':
    # Run the Flask application
    app.run(host='0.0.0.0', port=8000, debug=True)