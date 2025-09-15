"""
PDF to DOCX Converter Script
Converts a PDF file to DOCX format using the pdf2docx library.
"""


import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def pdf_to_docx(pdf_path, docx_path=None):
    """
    Convert PDF file to DOCX format
    """
    import os
    from pdf2docx import Converter
    # Check if PDF file exists
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    # Generate output path if not provided
    if docx_path is None:
        base_name = os.path.splitext(pdf_path)[0]
        docx_path = f"{base_name}.docx"
    # Convert
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()
    return docx_path


# --- GUI Application Skeleton ---
class ConverterApp:

    def __init__(self, root):
        self.root = root
        self.root.title("PDF/DOCX Converter")
        self.root.geometry("400x300")  # Increased height for logs
        self.root.configure(bg='#F5F5DC')  # Sepia background
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)  # Handle window close

        # Input file selection
        self.input_label = tk.Label(self.root, text="Input File:", bg='#F5F5DC')
        self.input_label.pack(pady=(20, 0))
        self.input_entry = tk.Entry(self.root, width=40)
        self.input_entry.pack(padx=10)
        self.input_button = tk.Button(self.root, text="Browse...", command=self.browse_input)
        self.input_button.pack(pady=5)

        # Conversion direction radio buttons
        self.direction_label = tk.Label(self.root, text="Conversion Type:", bg='#F5F5DC')
        self.direction_label.pack(pady=(10, 0))
        self.conversion_var = tk.StringVar(value="pdf2docx")
        self.radio_pdf2docx = tk.Radiobutton(self.root, text="PDF → DOCX", variable=self.conversion_var, value="pdf2docx", command=self.on_conversion_type_change, bg='#F5F5DC')
        self.radio_docx2pdf = tk.Radiobutton(self.root, text="DOCX → PDF", variable=self.conversion_var, value="docx2pdf", command=self.on_conversion_type_change, bg='#F5F5DC')
        self.radio_pdf2docx.pack()
        self.radio_docx2pdf.pack()

        # Convert button
        self.convert_button = tk.Button(self.root, text="Convert", state=tk.DISABLED, command=self.run_conversion)
        self.convert_button.pack(pady=5)

        # Log toggle button
        self.log_visible = False
        self.log_button = tk.Button(self.root, text="Show Logs", command=self.toggle_logs, bg='#F5F5DC')
        self.log_button.pack(pady=5)

        # Log frame (hidden by default)
        self.log_frame = tk.Frame(self.root, bg='#F5F5DC')
        self.log_text = scrolledtext.ScrolledText(self.log_frame, width=50, height=10, bg='white', fg='black')
        self.log_text.pack(padx=10, pady=10)
        # Don't pack the frame yet

    def on_closing(self):
        self.log("Application closing...")
        self.root.destroy()

    def toggle_logs(self):
        if self.log_visible:
            self.log_frame.pack_forget()
            self.log_button.config(text="Show Logs")
            self.log_visible = False
            self.root.geometry("400x250")  # Reset size
        else:
            self.log_frame.pack(pady=10)
            self.log_button.config(text="Hide Logs")
            self.log_visible = True
            self.root.geometry("400x350")  # Increase size

    def on_conversion_type_change(self):
        # Optionally, update file dialog filters or UI based on conversion type
        pass

    def log(self, message):
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)  # Scroll to end

    def run_conversion(self):
        input_path = self.input_entry.get()
        conv_type = self.conversion_var.get()

        if not input_path:
            error_msg = "Please select an input file."
            self.log(f"Error: {error_msg}")
            messagebox.showerror("Missing Information", error_msg)
            return

        # Auto-generate output path
        import os
        base = os.path.splitext(input_path)[0]
        if conv_type == "pdf2docx":
            output_path = base + ".docx"
        else:
            output_path = base + ".pdf"

        self.log(f"Output path: {output_path}")
        self.log(f"Input exists: {os.path.exists(input_path)}")
        self.log(f"Output dir writable: {os.access(os.path.dirname(output_path), os.W_OK)}")

        # Redirect stdout and stderr to prevent tqdm crash in bundled exe
        import sys
        from io import StringIO
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = StringIO()
        sys.stderr = StringIO()

        self.log(f"Starting conversion: {input_path} -> {output_path}")

        if conv_type == "pdf2docx":
            try:
                pdf_to_docx(input_path, output_path)
                success_msg = f"Conversion successful!\nSaved as: {output_path}"
                self.log(success_msg)
                messagebox.showinfo("Success", success_msg)
            except Exception as e:
                error_msg = f"PDF to DOCX conversion failed: {e}"
                self.log(f"Error: {error_msg}")
                messagebox.showerror("Error", error_msg)
        elif conv_type == "docx2pdf":
            try:
                from docx2pdf import convert as docx2pdf_convert
                self.log("Calling docx2pdf_convert...")
                docx2pdf_convert(input_path, output_path)
                success_msg = f"Conversion successful!\nSaved as: {output_path}"
                self.log(success_msg)
                messagebox.showinfo("Success", success_msg)
            except ImportError:
                error_msg = "docx2pdf is not installed. Please install it to use DOCX→PDF conversion."
                self.log(f"Error: {error_msg}")
                messagebox.showerror("Missing Dependency", error_msg)
            except Exception as e:
                import traceback
                error_details = traceback.format_exc()
                error_msg = f"DOCX to PDF conversion failed: {e}\n\nFull traceback:\n{error_details}"
                self.log(f"Error: {error_msg}")
                messagebox.showerror("Error", error_msg)

        # Restore stdout and stderr
        sys.stdout = old_stdout
        sys.stderr = old_stderr

    def browse_input(self):
        conv_type = self.conversion_var.get()
        if conv_type == "pdf2docx":
            filetypes = [("PDF files", "*.pdf")]
        else:
            filetypes = [("Word Documents", "*.docx")]

        filename = filedialog.askopenfilename(
            title="Select input file",
            filetypes=filetypes
        )
        if filename:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)
            self.convert_button.config(state=tk.NORMAL)
            self.log(f"Input file selected: {filename}")

def main():
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

