import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter
import os

class PDFSlicerApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="superhero")  # Try "cosmo", "darkly", etc.
        self.title("📄 PDF Slicer Dashboard")
        self.geometry("700x400")
        self.resizable(False, False)

        self.pdf_path = None
        self.total_pages = 0

        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self, text="PDF Slicer Dashboard", font=("Helvetica", 20, "bold")).pack(pady=20)

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=X, padx=40)

        # File section
        ttk.Label(frame, text="Select PDF File:", font=("Helvetica", 12)).grid(row=0, column=0, sticky=W, pady=10)
        self.file_label = ttk.Label(frame, text="No file selected", font=("Helvetica", 10), width=40)
        self.file_label.grid(row=0, column=1, sticky=W)
        ttk.Button(frame, text="Browse", bootstyle=INFO, command=self.browse_pdf).grid(row=0, column=2, padx=10)

        # Page range section
        ttk.Label(frame, text="Enter Page Range (e.g., 2-5, 7-9):", font=("Helvetica", 12)).grid(row=1, column=0, sticky=W, pady=10)
        self.range_entry = ttk.Entry(frame, width=30)
        self.range_entry.grid(row=1, column=1, sticky=W)

        # Output file
        ttk.Label(frame, text="Output File Name:", font=("Helvetica", 12)).grid(row=2, column=0, sticky=W, pady=10)
        self.output_entry = ttk.Entry(frame, width=30)
        self.output_entry.grid(row=2, column=1, sticky=W)
        self.output_entry.insert(0, "sliced_output.pdf")

        # Slice button
        ttk.Button(self, text="✂ Slice PDF", bootstyle=SUCCESS, command=self.slice_pdf).pack(pady=25)

        # Info area
        self.info_label = ttk.Label(self, text="", font=("Helvetica", 10))
        self.info_label.pack()

    def browse_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            try:
                reader = PdfReader(file_path)
                self.total_pages = len(reader.pages)
                self.info_label.config(text=f"✅ Loaded successfully. Total Pages: {self.total_pages}")
            except Exception as e:
                messagebox.showerror("Error", f"Cannot read PDF file.\n{e}")

    def parse_ranges(self, text):
        """Parse input like 1-3, 5-6 into [1,2,3,5,6]"""
        pages = set()
        for part in text.split(","):
            part = part.strip()
            if "-" in part:
                start, end = map(int, part.split("-"))
                pages.update(range(start, end + 1))
            else:
                pages.add(int(part))
        return sorted(pages)

    def slice_pdf(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return

        range_text = self.range_entry.get()
        if not range_text:
            messagebox.showerror("Error", "Please enter a valid page range.")
            return

        output_name = self.output_entry.get()
        if not output_name.endswith(".pdf"):
            output_name += ".pdf"

        try:
            reader = PdfReader(self.pdf_path)
            writer = PdfWriter()
            total_pages = len(reader.pages)
            pages_to_add = self.parse_ranges(range_text)

            for p in pages_to_add:
                if 1 <= p <= total_pages:
                    writer.add_page(reader.pages[p - 1])

            save_path = filedialog.asksaveasfilename(
                initialfile=output_name,
                defaultextension=".pdf",
                filetypes=[("PDF Files", "*.pdf")]
            )

            if not save_path:
                return

            with open(save_path, "wb") as f:
                writer.write(f)

            messagebox.showinfo("Success", f"✅ Sliced PDF saved as:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong.\n{e}")

# Run the app
if __name__ == "__main__":
    app = PDFSlicerApp()
    app.mainloop()
