# 📄➡️📝 PDF to Word Converter

This project is a simple yet powerful web application for converting PDF files to Word (DOCX) documents while preserving the original layout and formatting as much as possible.

**✨ [Try the Live Demo Here](https://pdf-to-word-converter-nayef.onrender.com/) ✨**

*(Note: The free server may need about 30 seconds to "wake up" on the first visit)*

## Key Features

- **Layout Preservation:** Replicates the precise position and indentation of each paragraph.
- **Formatting Retention:** Retains the original font type, size, bold, and italic styling.
- **Spacing Management:** Detects and preserves vertical spacing between paragraphs.
- **Simple Interface:** A clean and easy-to-use interface for uploading and converting files.

## Technologies Used

- **Backend:** Flask (Python)
- **PDF Processing:** PyMuPDF (fitz)
- **Word Creation:** python-docx
- **Production Server:** Gunicorn
- **Hosting:** Render

## Running Locally (for Developers)

If you wish to run the project on your local machine:

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/nayefzedane/pdf-to-word-converter.git](https://github.com/nayefzedane/pdf-to-word-converter.git)
    cd pdf-to-word-converter
    ```

2.  **Install the requirements:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Run the application:**
    ```bash
    python app.py
    ```

4.  Open your browser and navigate to `http://127.0.0.1:5000`.
