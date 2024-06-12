from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, ttk
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Inches
import os
import threading

class PdfConvert:
    def __init__(self, root=None):
        if root is None:
            root = Tk()
            root.title("PDF to Word Converter")
            root.configure(background='blue')

        self.root = root

        # Labels
        Label(self.root, text="PDF File Path:", bg='blue', fg='white').grid(row=0, column=0, padx=10, pady=10)
        Label(self.root, text="Save Directory:", bg='blue', fg='white').grid(row=1, column=0, padx=10, pady=10)

        # Entry fields
        self.pdf_entry = Entry(self.root, width=50)
        self.pdf_entry.grid(row=0, column=1, padx=10, pady=10)
        self.save_entry = Entry(self.root, width=50)
        self.save_entry.grid(row=1, column=1, padx=10, pady=10)

        # Browse buttons
        Button(self.root, text="Browse", command=self.browse_pdf).grid(row=0, column=2, padx=10, pady=10)
        Button(self.root, text="Browse", command=self.browse_save).grid(row=1, column=2, padx=10, pady=10)

        # Convert and Cancel buttons
        self.convert_button = Button(self.root, text="Convert", command=self.convert_pdf_to_word)
        self.convert_button.grid(row=2, column=1, pady=20)
        self.cancel_button = Button(self.root, text="Cancel", command=self.cancel_conversion, state='disabled')
        self.cancel_button.grid(row=2, column=2, pady=20)

        # Progress bar
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=3, columnspan=3, padx=10, pady=10)

        self.output_folder = 'images'
        os.makedirs(self.output_folder, exist_ok=True)

    def browse_pdf(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        self.pdf_entry.delete(0, 'end')
        self.pdf_entry.insert(0, filename)

    def browse_save(self):
        save_dir = filedialog.askdirectory()
        self.save_entry.delete(0, 'end')
        self.save_entry.insert(0, save_dir)

    def convert_pdf_to_word(self):
        pdf_path = self.pdf_entry.get()
        save_dir = self.save_entry.get()

        if not pdf_path or not save_dir:
            messagebox.showerror("Error", "Please select PDF file and save directory.")
            return

        self.convert_button.config(state='disabled')
        self.cancel_button.config(state='normal')

        try:
            self.thread = threading.Thread(target=self.convert_thread, args=(pdf_path, save_dir))
            self.thread.start()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.reset_ui()

    def convert_thread(self, pdf_path, save_dir):
        try:
            images = convert_from_path(pdf_path)
            total_pages = len(images)
            self.progress["maximum"] = total_pages

            for i, image in enumerate(images):
                image_path = os.path.join(self.output_folder, f'page_{i+1}.png')
                image.save(image_path, 'PNG')
                self.progress["value"] = i + 1
                self.progress.update()

            docx_path = os.path.join(save_dir, "output.docx")
            self.png_to_word([os.path.join(self.output_folder, f'page_{i+1}.png') for i in range(total_pages)], docx_path)

            messagebox.showinfo("Conversion Successful", f"The PDF file was converted to PNG images and then saved as a Word document at {docx_path}.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during conversion: {str(e)}")
        finally:
            self.reset_ui()

    def cancel_conversion(self):
        if hasattr(self, 'thread') and self.thread.is_alive():
            self.thread.terminate()
            self.thread.join()

        self.reset_ui()

    def reset_ui(self):
        self.convert_button.config(state='normal')
        self.cancel_button.config(state='disabled')
        self.progress["value"] = 0

    def png_to_word(self, image_paths, docx_path):
        doc = Document()

        for image_path in image_paths:
            doc.add_picture(image_path, width=Inches(6))

        doc.save(docx_path)

