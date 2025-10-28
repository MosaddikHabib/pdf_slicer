import os
import sys
import webbrowser
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, simpledialog

# Alternative to PyMuPDF: pikepdf (QPDF-based, works on Python 3.13)
import pikepdf


def parse_ranges(text, *, keep_input_order=False):
    """
    Turn '1-3, 8, 10-12' into 1-based page numbers.
    - Deduplicated
    - Positive only
    - keep_input_order: preserve order the user typed (vs sorted numeric)
    """
    # Use an ordered structure to dedupe while preserving insertion order.
    seen = {}
    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start, end = int(a), int(b)
            except ValueError:
                continue
            if start > end:
                start, end = end, start
            for p in range(start, end + 1):
                if p > 0:
                    seen[p] = True
        else:
            try:
                p = int(part)
                if p > 0:
                    seen[p] = True
            except ValueError:
                continue

    pages = list(seen.keys())
    if not keep_input_order:
        pages.sort()
    return pages


class PDFSlicerApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="darkly")  # try "superhero", "cosmo", "flatly", etc.
        self.title("📄 PDF Slicer Dashboard (pikepdf)")
        self.geometry("880x520")
        self.resizable(False, False)

        self.pdf_path = None
        self.total_pages = 0
        self._password_cache = None  # remember password for encrypted docs in this session

        # UI state vars
        self.mode_var = ttk.StringVar(value="extract")  # extract | delete
        self.keep_order_var = ttk.BooleanVar(value=False)
        self.open_when_done_var = ttk.BooleanVar(value=True)

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
        self.file_label = ttk.Label(card, text="No file selected", width=52, anchor=W)
        self.file_label.grid(row=0, column=1, sticky=W, padx=(10, 10))
        ttk.Button(card, text="Browse", bootstyle=INFO, command=self.browse_pdf).grid(row=0, column=2, padx=(6, 0))

        # Row: page ranges + helpers
        ttk.Label(card, text="Page Ranges:", font=("-size", 11)).grid(row=1, column=0, sticky=W, pady=6)
        self.range_entry = ttk.Entry(card, width=40)
        self.range_entry.grid(row=1, column=1, sticky=W, padx=(10, 10))

        helpers = ttk.Frame(card)
        helpers.grid(row=1, column=2, sticky=W)
        ttk.Button(helpers, text="All", bootstyle=SECONDARY, command=self.fill_all).pack(side=LEFT, padx=2)
        ttk.Button(helpers, text="Odd", bootstyle=SECONDARY, command=self.fill_odd).pack(side=LEFT, padx=2)
        ttk.Button(helpers, text="Even", bootstyle=SECONDARY, command=self.fill_even).pack(side=LEFT, padx=2)

        ttk.Label(card, text="e.g., 1-3, 8, 10-12", bootstyle=SECONDARY).grid(row=2, column=1, sticky=W, pady=(2, 0))

        # Row: mode (extract/delete) + keep order
        mode_row = ttk.Frame(card)
        mode_row.grid(row=3, column=1, sticky=W, pady=(10, 0))
        ttk.Radiobutton(mode_row, text="Extract selected pages", variable=self.mode_var,
                        value="extract", bootstyle=SUCCESS).pack(side=LEFT, padx=(0, 10))
        ttk.Radiobutton(mode_row, text="Delete selected pages", variable=self.mode_var,
                        value="delete", bootstyle=WARNING).pack(side=LEFT)

        opts_row = ttk.Frame(card)
        opts_row.grid(row=4, column=1, sticky=W, pady=(6, 0))
        ttk.Checkbutton(opts_row, text="Keep input order", variable=self.keep_order_var).pack(side=LEFT, padx=(0, 10))
        ttk.Checkbutton(opts_row, text="Open file when done", variable=self.open_when_done_var).pack(side=LEFT)

        # Row: output name
        ttk.Label(card, text="Output File Name:", font=("-size", 11)).grid(row=5, column=0, sticky=W, pady=10)
        self.output_entry = ttk.Entry(card, width=40)
        self.output_entry.grid(row=5, column=1, sticky=W, padx=(10, 10))
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

    # -------- Helpers --------
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
        self.mode_var.set("extract")
        self.keep_order_var.set(False)
        self.open_when_done_var.set(True)

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

    def fill_all(self):
        if self.total_pages <= 0:
            messagebox.showinfo("Info", "Load a PDF first.")
            return
        self.range_entry.delete(0, "end")
        self.range_entry.insert(0, f"1-{self.total_pages}")

    def fill_odd(self):
        if self.total_pages <= 0:
            messagebox.showinfo("Info", "Load a PDF first.")
            return
        odds = ",".join(str(p) for p in range(1, self.total_pages + 1, 2))
        self.range_entry.delete(0, "end")
        self.range_entry.insert(0, odds)

    def fill_even(self):
        if self.total_pages <= 0:
            messagebox.showinfo("Info", "Load a PDF first.")
            return
        evens = ",".join(str(p) for p in range(2, self.total_pages + 1, 2))
        self.range_entry.delete(0, "end")
        self.range_entry.insert(0, evens)

    # -------- Core slicing --------
    def slice_pdf(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return

        ranges_text = self.range_entry.get().strip()
        if not ranges_text:
            messagebox.showerror("Error", "Please enter page ranges (e.g., 1-3, 8, 10-12) or use helpers.")
            return

        pages_1based = parse_ranges(ranges_text, keep_input_order=self.keep_order_var.get())
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

        # Prevent overwriting the source PDF
        if os.path.abspath(save_path) == os.path.abspath(self.pdf_path):
            messagebox.showerror("Error", "Output path cannot be the same as the source PDF.")
            return

        # Check directory is writable
        outdir = os.path.dirname(save_path) or "."
        if not os.path.isdir(outdir):
            messagebox.showerror("Error", f"Directory does not exist:\n{outdir}")
            return
        if not os.access(outdir, os.W_OK):
            messagebox.showerror("Error", f"Directory is not writable:\n{outdir}")
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

            # Determine which pages to write, based on mode
            if self.mode_var.get() == "extract":
                target_pages = pages_1based
            else:  # delete mode
                # keep all pages except listed ones
                remove_set = set(pages_1based)
                target_pages = [p for p in range(1, total + 1) if p not in remove_set]

            if not target_pages:
                src.close()
                messagebox.showerror("Error", "No pages to write after applying your selection.")
                return

            # Create output; copy document metadata when possible
            out = pikepdf.Pdf.new()
            try:
                out.root.Info = src.root.Info
            except Exception:
                pass  # metadata might not exist; ignore

            # Progress bar setup
            self.progress["value"] = 0
            self.progress["maximum"] = len(target_pages) or 1
            self.update_idletasks()

            # Append pages (convert to 0-based)
            for i, p in enumerate(target_pages, start=1):
                out.pages.append(src.pages[p - 1])
                self.progress["value"] = i
                self.update_idletasks()

            # Save (let pikepdf/qpdf handle compression)
            out.save(save_path, linearize=False)  # set linearize=True for web-optimized
            out.close()
            src.close()

            messagebox.showinfo("Success", f"✅ Sliced PDF saved:\n{save_path}")

            if self.open_when_done_var.get():
                self._open_file_crossplatform(save_path)

        except pikepdf.PasswordError:
            messagebox.showerror("Error", "Incorrect password for encrypted PDF.")
        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong.\n{e}")

    @staticmethod
    def _open_file_crossplatform(path):
        try:
            if sys.platform.startswith("darwin"):
                os.system(f'open "{path}"')
            elif os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                # Linux
                try:
                    os.system(f'xdg-open "{path}"')
                except Exception:
                    webbrowser.open_new_tab(f"file://{os.path.abspath(path)}")
        except Exception:
            pass


if __name__ == "__main__":
    app = PDFSlicerApp()
    app.mainloop()
