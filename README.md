# Doc-Tor 🩺

**Doc-Tor** is a powerful, all-in-one PDF manipulation tool built with Python. It provides a modern Graphical User Interface (GUI) and a robust Command Line Interface (CLI) for handling your PDF needs efficiently.

##  Features

*   **PDF to Word & Back:** Convert PDFs to editable Word documents (`.docx`) and convert them back to professional PDFs after editing.
*   **Edit & Redact:** Search and replace text directly in the PDF, or securely redact (black out) sensitive information.
*   **Page Manipulation:**
    *   **Rotate:** Fix orientation (90°, 180°, 270°).
    *   **Delete:** Remove unwanted pages easily.
    *   **Split/Extract:** Save specific page ranges as new files.
*   **Security:** Encrypt your PDFs with AES-256 passwords or decrypt protected files.
*   **Merge:** Combine multiple PDF files into a single document.
*   **Extract Text:** Dump text content from PDFs for analysis.
*   **Metadata:** View detailed file information (Author, Page Count, Encryption status).

##  Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/yourusername/Doc-Tor.git
    cd Doc-Tor
    ```

2.  **Install dependencies:**
    Doc-Tor relies on `Typer`, `PyMuPDF`, `CustomTkinter`, `pdf2docx`, and `docx2pdf`.
    ```bash
    pip install -r requirements.txt
    ```
    *Note: `docx2pdf` requires Microsoft Word to be installed on Windows.*

##  Usage

### Graphical User Interface (GUI)
Launch the modern dark-mode GUI to access all features in a user-friendly way.
```bash
python gui.py
```

### Command Line Interface (CLI)
For automation and quick tasks, use the CLI.

**Examples:**

*   **Convert PDF to Word:**
    ```bash
    python cli.py pdf-to-word document.pdf
    ```

*   **Redact Text:**
    ```bash
    python cli.py redact confidential.pdf "SECRET"
    ```

*   **Merge Files:**
    ```bash
    python cli.py merge part1.pdf part2.pdf --output full_report.pdf
    ```

*   **Encrypt a File:**
    ```bash
    python cli.py encrypt sensitive.pdf "mySecurePassword"
    ```

*   **View Help:**
    ```bash
    python cli.py --help
    ```

##  Dependencies

*   [Typer](https://typer.tiangolo.com/) - CLI building
*   [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/) - Core PDF processing
*   [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - Modern GUI
*   [pdf2docx](https://dothinking.github.io/pdf2docx/) - PDF to Word conversion
*   [docx2pdf](https://github.com/AlJohri/docx2pdf) - Word to PDF conversion
*   [Rich](https://github.com/Textualize/rich) - Beautiful terminal output

##  Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

##  License

This project is licensed under the MIT License - see the LICENSE file for details.
