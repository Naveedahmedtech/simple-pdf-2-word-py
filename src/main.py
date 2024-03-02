import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import os


class PDFtoWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title('PDF to Word Converter')
        self.root.geometry('400x150')

        # Frame for Buttons
        frame = tk.Frame(self.root)
        frame.pack(pady=20)

        # Button to choose PDF
        self.btn_choose_pdf = tk.Button(frame, text='Choose PDF', command=self.choose_pdf)
        self.btn_choose_pdf.pack(side=tk.LEFT, padx=10)

        # Button to start conversion
        self.btn_convert = tk.Button(frame, text='Convert to Word', command=self.convert_to_word, state=tk.DISABLED)
        self.btn_convert.pack(side=tk.LEFT, padx=10)

        self.pdf_file_path = ''

    def choose_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[('PDF files', '*.pdf')])
        if file_path:
            self.pdf_file_path = file_path
            self.btn_convert['state'] = tk.NORMAL
            messagebox.showinfo("File Selected", f"File: {os.path.basename(file_path)}")

    def convert_to_word(self):
        if self.pdf_file_path:
            try:
                # Create the 'outputs' directory if it doesn't exist
                output_dir = os.path.join(os.getcwd(), 'outputs')
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)

                # Set the output path for the converted file
                output_path = os.path.join(output_dir, os.path.basename(self.pdf_file_path).replace('.pdf', '.docx'))

                # Perform the conversion
                converter = Converter(self.pdf_file_path)
                converter.convert(output_path, start=0, end=None)
                converter.close()

                messagebox.showinfo("Success", "PDF has been converted to Word successfully!")
                self.btn_convert['state'] = tk.DISABLED
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showerror("Error", "No PDF file selected!")


if __name__ == '__main__':
    root = tk.Tk()
    app = PDFtoWordApp(root)
    root.mainloop()
