import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

def slice_pdf():
    # Open file dialog to choose PDF
    file_path = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not file_path:
        messagebox.showwarning("No file selected", "Please select a PDF file.")
        return

    # Ask user for page range (e.g. 2-5)
    page_range = simpledialog.askstring(
        "Page Range",
        "Enter page range (e.g., 2-5):"
    )

    if not page_range or "-" not in page_range:
        messagebox.showerror("Invalid input", "Please enter a valid page range like 2-5.")
        return

    try:
        start, end = map(int, page_range.split("-"))
    except ValueError:
        messagebox.showerror("Invalid input", "Page range must be two numbers separated by '-'.")
        return

    # Read original PDF
    reader = PdfReader(file_path)
    writer = PdfWriter()

    total_pages = len(reader.pages)

    if start < 1 or end > total_pages or start > end:
        messagebox.showerror("Error", f"Invalid range! PDF has {total_pages} pages.")
        return

    # Add pages to new PDF
    for i in range(start - 1, end):
        writer.add_page(reader.pages[i])

    # Ask where to save
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")],
        title="Save new PDF as..."
    )

    if not save_path:
        return

    # Write new PDF
    with open(save_path, "wb") as output_file:
        writer.write(output_file)

    messagebox.showinfo("Success", f"New PDF created successfully!\nSaved to:\n{save_path}")

# GUI setup
root = tk.Tk()
root.withdraw()  # Hide main window
slice_pdf()
