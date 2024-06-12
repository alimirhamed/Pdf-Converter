PDF to Word Converter GUI Tool
This project provides a Python-based graphical user interface (GUI) tool for converting PDF files into Word documents (.docx). It utilizes the pdf2image and docx libraries for extracting images from PDFs and creating Word documents, respectively. The tool is designed to be straightforward yet effective, featuring a progress bar to display conversion status and an option to cancel ongoing conversions.

Features:
PDF File Selection: Users can browse and select a PDF file for conversion.
Save Directory Selection: Option to choose the directory where the converted Word file will be saved.
Progress Bar: Displays conversion progress with a visual progress bar.
Cancellation: Ability to cancel ongoing conversions if needed.
Error Handling: Proper error messages are displayed in case of conversion failures.
Cross-platform: Works on Windows, macOS, and Linux.
Installation and Usage:
Clone the Repository:

bash
Copy code
git clone https://github.com/your-username/PdfConvert.git
cd PdfConvert
Install Dependencies:

Copy code
pip install -r requirements.txt
Run the Application:

css
Copy code
python main.py
Usage Instructions:

Click on the "Browse" button next to "PDF File Path" to select the PDF file.
Click on the "Browse" button next to "Save Directory" to choose where to save the converted Word file.
Click on the "Convert" button to start the conversion process. The progress bar will indicate the status.
Click on the "Cancel" button to abort ongoing conversions if needed.
Upon successful conversion, a message box will inform you of the completion and the location of the converted Word file.
Contributing:
Contributions are welcome! If you'd like to add features, improve existing functionality, or fix issues, please follow these steps:

Fork the repository and clone it locally.
Create a new branch for your feature: git checkout -b feature-name.
Make your changes and test thoroughly.
Commit your changes: git commit -am 'Add new feature'.
Push to the branch: git push origin feature-name.
Submit a pull request detailing your changes.
License:
This project is licensed under the MIT License - see the LICENSE file for details.

Acknowledgments:
This project utilizes the pdf2image and docx Python libraries.