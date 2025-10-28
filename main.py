import os
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, simpledialog

# Alternative to PyMuPDF: pikepdf (QPDF-based, works on Python 3.13)
import pikepdf


def parse_ranges(text):
    """
    Turn '1-3, 8, 10-12' into sorted, deduplicated 1-based page numbers.
    """
    pages = set()
    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start, end = int(a), int(b)
            if start > end:
                start, end = end, start
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    return sorted(p for p in pages if p > 0)


class PDFSlicerApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="darkly")  # try "superhero", "cosmo", "flatly", etc.
        self.title("📄 PDF Slicer Dashboard (pikepdf)")
        self.geometry("820x460")
        self.resizable(False, False)

        self.pdf_path = None
        self.total_pages = 0
        self._password_cache = None  # remember password for encrypted docs in this session

        self.build_ui()

    def build_ui(self):
        # Header
        header = ttk.Frame(self, padding=(20, 20, 20, 10))
        header.pack(fill=X)
        ttk.Label(header, text="PDF Slicer Dashboard", font=("Helvetica", 22, "bold")).pack(side=LEFT)
        ttk.Label(header, text="Powered by pikepdf (QPDF)", bootstyle=INFO).pack(side=RIGHT, padx=10)

        # Main card
        card = ttk.Frame(self, padding=20, bootstyle="secondary")
        card.pack(fill=X, padx=24, pady=10)

        # Row: file
        ttk.Label(card, text="Select PDF:", font=("-size", 11)).grid(row=0, column=0, sticky=W, pady=6)
        self.file_label = ttk.Label(card, text="No file selected", width=50, anchor=W)
        self.file_label.grid(row=0, column=1, sticky=W, padx=(10, 10))
        ttk.Button(card, text="Browse", bootstyle=INFO, command=self.browse_pdf).grid(row=0, column=2, padx=(6, 0))

        # Row: page ranges
        ttk.Label(card, text="Page Ranges:", font=("-size", 11)).grid(row=1, column=0, sticky=W, pady=6)
        self.range_entry = ttk.Entry(card, width=40)
        self.range_entry.grid(row=1, column=1, sticky=W, padx=(10, 10))
        ttk.Label(card, text="e.g., 1-3, 8, 10-12", bootstyle=SECONDARY).grid(row=1, column=2, sticky=W)

        # Row: output name
        ttk.Label(card, text="Output File Name:", font=("-size", 11)).grid(row=2, column=0, sticky=W, pady=6)
        self.output_entry = ttk.Entry(card, width=40)
        self.output_entry.grid(row=2, column=1, sticky=W, padx=(10, 10))
        self.output_entry.insert(0, "sliced_output.pdf")

        # Info area
        self.info = ttk.Label(self, text="Load a PDF to see details.", bootstyle=SECONDARY)
        self.info.pack(fill=X, padx=30, pady=(4, 12))

        # Progress
        pwrap = ttk.Frame(self, padding=(30, 0, 30, 0))
        pwrap.pack(fill=X)
        self.progress = ttk.Progressbar(pwrap, bootstyle="info-striped", mode="determinate")
        self.progress.pack(fill=X)

        # Action buttons
        btns = ttk.Frame(self, padding=20)
        btns.pack()
        ttk.Button(btns, text="✂ Slice PDF", bootstyle=SUCCESS, command=self.slice_pdf, width=20).pack(side=LEFT, padx=8)
        ttk.Button(btns, text="Reset", bootstyle=SECONDARY, command=self.reset, width=12).pack(side=LEFT, padx=8)
        ttk.Button(btns, text="Quit", bootstyle=DANGER, command=self.destroy, width=10).pack(side=LEFT, padx=8)

    def reset(self):
        self.pdf_path = None
        self.total_pages = 0
        self._password_cache = None
        self.file_label.config(text="No file selected")
        self.range_entry.delete(0, "end")
        self.output_entry.delete(0, "end")
        self.output_entry.insert(0, "sliced_output.pdf")
        self.info.config(text="Load a PDF to see details.")
        self.progress["value"] = 0
        self.progress["maximum"] = 100

    def _open_pdf_for_info(self, path):
        """Open the PDF just to read basic info (page count), handling encryption."""
        try:
            # Try without password first
            with pikepdf.open(path) as pdf:
                return len(pdf.pages), None
        except pikepdf.PasswordError:
            # Ask password
            pwd = simpledialog.askstring("Encrypted PDF", "Enter password:", show="*")
            if not pwd:
                raise
            # Try with provided password
            with pikepdf.open(path, password=pwd) as pdf:
                return len(pdf.pages), pwd

    def browse_pdf(self):
        path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF Files", "*.pdf")])
        if not path:
            return
        self.pdf_path = path
        self.file_label.config(text=os.path.basename(path))

        try:
            count, pwd = self._open_pdf_for_info(path)
            self.total_pages = count
            self._password_cache = pwd  # remember if supplied
            self.info.config(text=f"✅ Loaded: {self.total_pages} pages | {self.pdf_path}")
        except pikepdf.PasswordError:
            self.pdf_path = None
            self.total_pages = 0
            messagebox.showerror("Error", "PDF is encrypted and no/invalid password was provided.")
        except Exception as e:
            self.pdf_path = None
            self.total_pages = 0
            messagebox.showerror("Error", f"Failed to open PDF.\n{e}")

    def slice_pdf(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return

        ranges_text = self.range_entry.get().strip()
        if not ranges_text:
            messagebox.showerror("Error", "Please enter page ranges (e.g., 1-3, 8, 10-12).")
            return

        pages_1based = parse_ranges(ranges_text)
        if not pages_1based:
            messagebox.showerror("Error", "No valid pages parsed from your input.")
            return

        # Choose save path early
        default_name = self.output_entry.get().strip() or "sliced_output.pdf"
        if not default_name.lower().endswith(".pdf"):
            default_name += ".pdf"

        save_path = filedialog.asksaveasfilename(
            title="Save sliced PDF as...",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not save_path:
            return

        try:
            # Open source (handle encryption with cached or prompted password)
            try:
                src = pikepdf.open(self.pdf_path)  # try without password
            except pikepdf.PasswordError:
                pwd = self._password_cache or simpledialog.askstring("Encrypted PDF", "Enter password:", show="*")
                if not pwd:
                    messagebox.showerror("Error", "Password required to open this PDF.")
                    return
                src = pikepdf.open(self.pdf_path, password=pwd)
                self._password_cache = pwd

            total = len(src.pages)
            # Validate bounds
            bad = [p for p in pages_1based if p < 1 or p > total]
            if bad:
                src.close()
                messagebox.showerror("Error", f"Pages out of range (PDF has {total} pages): {', '.join(map(str, bad))}")
                return

            # Create output
            out = pikepdf.Pdf.new()

            # Progress bar
            self.progress["value"] = 0
            self.progress["maximum"] = len(pages_1based) or 1
            self.update_idletasks()

            # Append pages (convert to 0-based)
            for i, p in enumerate(pages_1based, start=1):
                out.pages.append(src.pages[p - 1])
                self.progress["value"] = i
                self.update_idletasks()

            # Save (let pikepdf/qpdf handle compression)
            out.save(save_path)
            out.close()
            src.close()

            messagebox.showinfo("Success", f"✅ Sliced PDF saved:\n{save_path}")

        except pikepdf.PasswordError:
            messagebox.showerror("Error", "Incorrect password for encrypted PDF.")
        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong.\n{e}")


if __name__ == "__main__":
    app = PDFSlicerApp()
    app.mainloop()
