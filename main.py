import os
import sys
import webbrowser
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, simpledialog
import tkinter as tk  # for Canvas

# Core PDF I/O (slice)
import pikepdf

# Preview stack (no PyMuPDF)
try:
    import pypdfium2 as pdfium
except Exception:
    pdfium = None

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None


def parse_ranges(text, *, keep_input_order=False):
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
        super().__init__(themename="superhero")
        self.title("📄 PDF Slicer Dashboard (pikepdf)")
        self.geometry("1120x610")
        self.resizable(False, False)

        self.pdf_path = None
        self.total_pages = 0
        self._password_cache = None

        # UI state
        self.mode_var = ttk.StringVar(value="extract")  # extract | delete
        self.keep_order_var = ttk.BooleanVar(value=False)
        self.open_when_done_var = ttk.BooleanVar(value=True)

        # Preview state (pypdfium2)
        self.preview_doc = None          # pdfium.PdfDocument
        self.preview_page_var = ttk.IntVar(value=1)
        self._preview_photo = None       # keep reference for Tk
        self._preview_enabled = bool(pdfium and Image and ImageTk)

        self.build_ui()

    # ---------- UI ----------
    def build_ui(self):
        header = ttk.Frame(self, padding=(18, 18, 18, 10))
        header.pack(fill=X)
        ttk.Label(header, text="PDF Slicer Dashboard", font=("Helvetica", 22, "bold")).pack(side=LEFT)
        info_txt = "Powered by pikepdf • Preview via PDFium"
        ttk.Label(header, text=info_txt, bootstyle=INFO).pack(side=RIGHT, padx=10)

        body = ttk.Frame(self, padding=(16, 0, 16, 16))
        body.pack(fill=BOTH, expand=YES)

        # LEFT controls
        left = ttk.Frame(body)
        left.pack(side=LEFT, fill=Y)

        card = ttk.Frame(left, padding=20, bootstyle="secondary")
        card.pack(fill=X, padx=(0, 16), pady=10)

        ttk.Label(card, text="Select PDF:", font=("-size", 11)).grid(row=0, column=0, sticky=W, pady=(2, 6))
        self.file_label = ttk.Label(card, text="No file selected", width=52, anchor=W)
        self.file_label.grid(row=0, column=1, sticky=W, padx=(10, 10))
        ttk.Button(card, text="Browse", bootstyle=INFO, command=self.browse_pdf).grid(row=0, column=2, padx=(6, 0))

        self.page_range_badge = ttk.Label(card, text="Pages: —", bootstyle="INFO-INVERSE", padding=(8, 4))
        self.page_range_badge.grid(row=1, column=1, sticky=W, pady=(0, 10), padx=(10, 10))

        ttk.Label(card, text="Page Ranges:", font=("-size", 11)).grid(row=2, column=0, sticky=W, pady=6)
        self.range_entry = ttk.Entry(card, width=40)
        self.range_entry.grid(row=2, column=1, sticky=W, padx=(10, 10))
        helpers = ttk.Frame(card); helpers.grid(row=2, column=2, sticky=W)
        ttk.Button(helpers, text="All", bootstyle=SECONDARY, command=self.fill_all).pack(side=LEFT, padx=2)

        ttk.Label(card, text="e.g., 1-3, 8, 10-12", bootstyle=SECONDARY).grid(row=3, column=1, sticky=W, pady=(2, 8))

        mode_row = ttk.Frame(card); mode_row.grid(row=4, column=1, sticky=W, pady=(2, 0))
        ttk.Radiobutton(mode_row, text="Extract selected pages", variable=self.mode_var,
                        value="extract", bootstyle=SUCCESS).pack(side=LEFT, padx=(0, 10))
        ttk.Radiobutton(mode_row, text="Delete selected pages", variable=self.mode_var,
                        value="delete", bootstyle=WARNING).pack(side=LEFT)

        opts_row = ttk.Frame(card); opts_row.grid(row=5, column=1, sticky=W, pady=(6, 0))
        ttk.Checkbutton(opts_row, text="Keep input order", variable=self.keep_order_var).pack(side=LEFT, padx=(0, 10))
        ttk.Checkbutton(opts_row, text="Open file when done", variable=self.open_when_done_var).pack(side=LEFT)

        ttk.Label(card, text="Output File Name:", font=("-size", 11)).grid(row=6, column=0, sticky=W, pady=10)
        self.output_entry = ttk.Entry(card, width=40)
        self.output_entry.grid(row=6, column=1, sticky=W, padx=(10, 10))
        self.output_entry.insert(0, "sliced_output.pdf")

        self.info = ttk.Label(left, text="Load a PDF to see details.", bootstyle=SECONDARY, padding=(6, 4))
        self.info.pack(fill=X, pady=(8, 8))

        pwrap = ttk.Frame(left, padding=(0, 0, 0, 6)); pwrap.pack(fill=X)
        self.progress = ttk.Progressbar(pwrap, bootstyle="info-striped", mode="determinate")
        self.progress.pack(fill=X)

        btns = ttk.Frame(left, padding=(0, 6, 0, 0)); btns.pack()
        ttk.Button(btns, text="✂ Slice PDF", bootstyle=SUCCESS, command=self.slice_pdf, width=20).pack(side=LEFT, padx=6)
        ttk.Button(btns, text="Reset", bootstyle=INFO, command=self.reset, width=12).pack(side=LEFT, padx=6)
        ttk.Button(btns, text="Quit", bootstyle=DANGER, command=self.destroy, width=10).pack(side=LEFT, padx=6)

        # RIGHT preview
        preview_card = ttk.Frame(body, padding=16, bootstyle="PRIMARY")
        preview_card.pack(side=LEFT, fill=BOTH, expand=YES)

        ttk.Label(preview_card, text="Quick Preview", font=("Helvetica", 14, "bold")).pack(anchor=W)
        ttk.Separator(preview_card).pack(fill=X, pady=6)

        self.preview_wrap = ttk.Frame(preview_card, padding=(6, 6, 6, 6), bootstyle="secondary")
        self.preview_wrap.pack(fill=BOTH, expand=YES)

        self.canvas_w, self.canvas_h = 420, 520
        self.preview_canvas = tk.Canvas(self.preview_wrap, width=self.canvas_w, height=self.canvas_h,
                                        highlightthickness=0, bg="#111")
        self.preview_canvas.pack(padx=4, pady=4)

        control_row = ttk.Frame(preview_card); control_row.pack(fill=X, pady=(8, 0))
        self.page_spin = ttk.Spinbox(control_row, from_=1, to=1, textvariable=self.preview_page_var,
                                     width=6, command=self._preview_from_controls, bootstyle="light")
        self.page_spin.pack(side=RIGHT, padx=(6, 0))
        self.page_slider = ttk.Scale(control_row, from_=1, to=1, orient=HORIZONTAL,
                                     bootstyle="info", command=self._slider_changed)
        self.page_slider.pack(side=RIGHT, fill=X, expand=YES, padx=(6, 6))

        self.preview_hint = ttk.Label(preview_card, bootstyle=SECONDARY, padding=(6, 4))
        self.preview_hint.pack(anchor=W, pady=(6, 0))
        self._update_preview_hint()

    def _update_preview_hint(self, msg=None):
        if msg:
            self.preview_hint.config(text=msg)
            return
        if self._preview_enabled:
            self.preview_hint.config(text="Tip: Use the slider or box to change the preview page.")
        else:
            self.preview_hint.config(
                text="Preview unavailable. Install:  pip install pypdfium2 pillow"
            )

    # ---------- Helpers ----------
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
        self.page_range_badge.config(text="Pages: —")
        self._close_preview_doc()
        self._clear_preview()

    def _open_pdf_for_info(self, path):
        try:
            with pikepdf.open(path) as pdf:
                return len(pdf.pages), None
        except pikepdf.PasswordError:
            pwd = simpledialog.askstring("Encrypted PDF", "Enter password:", show="*")
            if not pwd:
                raise
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
            self._password_cache = pwd
            self.info.config(text=f"✅ Loaded: {self.total_pages} pages | {self.pdf_path}")
            self.page_range_badge.config(text=f"Pages: 1–{self.total_pages}")
            self._setup_preview_controls()
            self._open_preview_doc()
            self.after(50, lambda: self._render_preview(1))
        except pikepdf.PasswordError:
            self.pdf_path = None
            self.total_pages = 0
            messagebox.showerror("Error", "PDF is encrypted and no/invalid password was provided.")
            self.page_range_badge.config(text="Pages: —")
            self._close_preview_doc()
            self._clear_preview()
        except Exception as e:
            self.pdf_path = None
            self.total_pages = 0
            messagebox.showerror("Error", f"Failed to open PDF.\n{e}")
            self.page_range_badge.config(text="Pages: —")
            self._close_preview_doc()
            self._clear_preview()

    def fill_all(self):
        if self.total_pages <= 0:
            messagebox.showinfo("Info", "Load a PDF first.")
            return
        self.range_entry.delete(0, "end")
        self.range_entry.insert(0, f"1-{self.total_pages}")

    # ---------- Preview handling (pypdfium2) ----------
    def _setup_preview_controls(self):
        maxp = max(1, self.total_pages or 1)
        self.page_slider.configure(from_=1, to=maxp)
        self.page_spin.configure(from_=1, to=maxp)
        self.preview_page_var.set(1)

    def _open_preview_doc(self):
        self._close_preview_doc()
        if not self._preview_enabled or not self.pdf_path:
            return
        try:
            # pypdfium2 handles passwords in the constructor
            self.preview_doc = pdfium.PdfDocument(self.pdf_path, password=self._password_cache)
            # If password is wrong, rendering will fail—show a hint then.
        except Exception as e:
            self.preview_doc = None
            self._update_preview_hint(f"Preview error: {e}")

    def _close_preview_doc(self):
        # PdfDocument closes when dereferenced; explicit close not required.
        self.preview_doc = None

    def _clear_preview(self):
        self.preview_canvas.delete("all")
        self._preview_photo = None
        self.preview_canvas.create_text(
            self.canvas_w // 2, self.canvas_h // 2,
            text="No Preview",
            fill="#888",
            font=("Helvetica", 16, "bold")
        )

    def _render_preview(self, page_num: int):
        if not self._preview_enabled or not self.preview_doc:
            self._clear_preview()
            return
        try:
            # bounds
            page_count = len(self.preview_doc)
            pn = max(1, min(page_num, page_count))
            page = self.preview_doc[pn - 1]

            # Render using PDFium -> PIL image
            # scale ~ 1.5x for crisp preview; you can tune as needed
            bitmap = page.render(scale=1.5)
            pil_img = bitmap.to_pil()

            # fit to canvas
            pil_img.thumbnail((self.canvas_w - 12, self.canvas_h - 12))
            self._preview_photo = ImageTk.PhotoImage(pil_img)

            self.preview_canvas.delete("all")
            x = (self.canvas_w - self._preview_photo.width()) // 2
            y = (self.canvas_h - self._preview_photo.height()) // 2
            self.preview_canvas.create_image(x, y, anchor="nw", image=self._preview_photo)

            # page badge
            self.preview_canvas.create_rectangle(8, 8, 98, 34, fill="#0d6efd", width=0)
            self.preview_canvas.create_text(54, 21, text=f"Page {pn}", fill="white", font=("Helvetica", 10, "bold"))

            self._update_preview_hint()
        except Exception as e:
            self._update_preview_hint(f"Preview render error: {e}")
            self._clear_preview()

    def _slider_changed(self, value):
        try:
            pn = int(float(value))
        except Exception:
            pn = self.preview_page_var.get()
        self.preview_page_var.set(pn)
        self._render_preview(pn)

    def _preview_from_controls(self):
        pn = self.preview_page_var.get()
        self.page_slider.set(pn)
        self._render_preview(pn)

    # ---------- Core slicing ----------
    def slice_pdf(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please select a PDF file first.")
            return

        ranges_text = self.range_entry.get().strip()
        if not ranges_text:
            messagebox.showerror("Error", "Please enter page ranges (e.g., 1-3, 8, 10-12) or use All.")
            return

        pages_1based = parse_ranges(ranges_text, keep_input_order=self.keep_order_var.get())
        if not pages_1based:
            messagebox.showerror("Error", "No valid pages parsed from your input.")
            return

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

        if os.path.abspath(save_path) == os.path.abspath(self.pdf_path):
            messagebox.showerror("Error", "Output path cannot be the same as the source PDF.")
            return

        outdir = os.path.dirname(save_path) or "."
        if not os.path.isdir(outdir):
            messagebox.showerror("Error", f"Directory does not exist:\n{outdir}")
            return
        if not os.access(outdir, os.W_OK):
            messagebox.showerror("Error", f"Directory is not writable:\n{outdir}")
            return

        try:
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
            bad = [p for p in pages_1based if p < 1 or p > total]
            if bad:
                src.close()
                messagebox.showerror("Error", f"Pages out of range (PDF has {total} pages): {', '.join(map(str, bad))}")
                return

            if self.mode_var.get() == "extract":
                target_pages = pages_1based
            else:
                remove_set = set(pages_1based)
                target_pages = [p for p in range(1, total + 1) if p not in remove_set]

            if not target_pages:
                src.close()
                messagebox.showerror("Error", "No pages to write after applying your selection.")
                return

            out = pikepdf.Pdf.new()
            try:
                out.root.Info = src.root.Info
            except Exception:
                pass

            self.progress["value"] = 0
            self.progress["maximum"] = len(target_pages) or 1
            self.update_idletasks()

            for i, p in enumerate(target_pages, start=1):
                out.pages.append(src.pages[p - 1])
                self.progress["value"] = i
                self.update_idletasks()

            out.save(save_path, linearize=False)
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
                try:
                    os.system(f'xdg-open "{path}"')
                except Exception:
                    webbrowser.open_new_tab(f"file://{os.path.abspath(path)}")
        except Exception:
            pass


if __name__ == "__main__":
    app = PDFSlicerApp()
    app.mainloop()
