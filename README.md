# Udemy Notes Extractor

This script extracts notes from a Udemy course's HTML file and saves them as **Excel**, **PDF**, and **Word** files. It preserves bold text formatting and organizes the notes by timestamp, titles, and content.

---

## Features
- Extracts notes with timestamps, titles, and content.
- Saves notes in **Excel**, **PDF**, and **Word** formats.
- Keeps bold text formatting in all output files.
- Works with HTML files saved from Udemy.

---

## Requirements
1. **Python 3.8 or newer** installed. Download it from [python.org](https://www.python.org/).
2. Install the required Python libraries:
   ```bash
   pip install bs4 fpdf xlsxwriter python-docx html2text
   ```

---

## How to Use

### 1. Save Your Udemy Notes as an HTML File
1. Open your Udemy course in your browser.
2. Go to the notes section where your saved bookmarks are visible.
3. Right-click anywhere on the page and select **Save As** (or **Save Page As**).
4. Choose **Webpage, Complete** as the format.
5. Save the file to your computer.
6. Rename the file (if needed) to something like `course.html` or keep the original name. The script works with any `.html` file saved from Udemy.

---

### 2. Run the Script
1. Place the HTML file in the same folder as the script.
2. Run the script:
   ```bash
   python extract_udemy_notes.py
   ```

---

### 3. View Your Notes
After running the script, you'll find the following files in the same folder:
- `Udemy_Notes.xlsx` (Excel format)
- `Udemy_Notes.pdf` (PDF format)
- `Udemy_Notes.docx` (Word format)

---

## Important Notes
- The script assumes your HTML file contains Udemy's default structure for notes. If the structure is different, the script may need adjustments.
- You can use any HTML file saved from the Udemy course notes section.

---

## Troubleshooting
1. **File Not Found**: Ensure the HTML file is in the same folder as the script and named correctly (`course.html` or any name you set in the script).
2. **Permission Error**: Close the output files (Excel, PDF, or Word) before running the script again.
3. **Missing Libraries**: Reinstall required libraries if you encounter errors:
   ```bash
   pip install bs4 fpdf xlsxwriter python-docx html2text
   ```

---

## Contributing
Feel free to suggest improvements or fork the repository to add new features!

---

## License
This project is licensed under the MIT License. See the `LICENSE` file for details.

---

