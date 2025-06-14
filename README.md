# Intelligent PDF Converter

This is an advanced desktop application built with Python that creates a powerful document reconstruction pipeline. It takes scanned, non-selectable PDF files, performs high-accuracy Optical Character Recognition (OCR), intelligently formats the extracted text to identify headings, and re-compiles it into a clean, selectable, and beautifully typeset new PDF document.

This project was built to demonstrate end-to-end automation capabilities, integrating multiple external tools and leveraging parallel processing to create a sophisticated tool for content digitization, a key skill for the **AI Automation Developer** role at **TheSoul Publishing**.

---

### Key Result: Before & After

This tool transforms unusable, image-based PDFs into professional, fully searchable documents.


---

### System Architecture

The application orchestrates a multi-stage, parallel processing pipeline to ensure both accuracy and speed, even with large documents.



---

### The Solution: A Multi-Stage Pipeline

This application provides a complete, user-friendly solution by orchestrating a complex pipeline:

1.  **PDF Parsing:** Uses `PyMuPDF` to efficiently extract pages from the source PDF as high-resolution images.
2.  **Parallel OCR:** Leverages Python's `multiprocessing` module to perform OCR on multiple pages simultaneously, dramatically speeding up the process on multi-core CPUs.
3.  **High-Accuracy OCR:** Utilizes the powerful `Tesseract` OCR engine to convert the images into raw text.
4.  **Intelligent Formatting:** Applies a custom set of heuristic rules and regular expressions (`regex`) to analyze the raw text, identify structural elements like headings and chapters, and format them with Markdown syntax.
5.  **High-Quality Recompilation:** Uses the `Pandoc` universal document converter to transform the cleaned and formatted Markdown text into a new, professional-grade PDF with selectable text, proper fonts, and clean margins.

### Application Interface

A user-friendly and responsive interface built with `customtkinter` allows for easy file management and output customization.



### Technology Stack
*   **Python 3:**
    *   `customtkinter` for the GUI.
    *   `multiprocessing` for parallel execution.
    *   `pytesseract` & `Pillow (PIL)` for OCR interfacing.
    *   `PyMuPDF (fitz)` for PDF manipulation.
*   **Tesseract OCR Engine** (External Dependency)
*   **Pandoc** (External Dependency)

### Installation & Usage

**1. Install External Dependencies:**
This application requires Tesseract and Pandoc to be installed on your system and accessible via the command line (added to your system's PATH).
*   **Install Tesseract OCR:** [Tesseract Installation Guide](https://tesseract-ocr.github.io/tessdoc/Installation.html)
*   **Install Pandoc:** [Pandoc Installation Guide](https://pandoc.org/installing.html)

**2. Clone the Repository & Install Python Libraries:**
```bash
# Clone this repository
git clone https://github.com/Unreliable-Support/ocr-pdf-converter.git


# Navigate to the project directory
cd intelligent-pdf-converter

# Install required Python libraries
pip install -r requirements.txt
