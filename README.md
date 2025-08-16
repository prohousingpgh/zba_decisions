# zba_decisions

Place the ZBA decision PDFs in the input/ directory, and then run the extract_data.py script. Data will be extracted and placed in output.csv.

Python dependencies:<br>
pip install pdfplumber<br>
pip install pdf2image<br>
pip install pytesseract

You also need to install poppler to get pdf2image working properly. See https://github.com/Belval/pdf2image?tab=readme-ov-file for instructions on how to do this.

You will need to install Tesseract OCR to get pytesseract working - instructions for this are on https://pypi.org/project/pytesseract/