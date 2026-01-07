import os
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2docx import Converter
import fitz  # PyMuPDF
from PIL import Image, ImageTk
from docx import Document

class PDFtoDOCXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to DOCX Converter with Preview")
        self.root.geometry("1000x700")
        
        self.pdf_path = None
        self.preview_images = []
        self.current_page = 0
        
        self.setup_ui()
        
    def setup_ui(self):
        # Top frame for file selection
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10)
        
        tk.Button(top_frame, text="Select PDF", command=self.select_pdf, 
                 bg="#4CAF50", fg="white", padx=20).pack(side=tk.LEFT, padx=5)
        
        self.file_label = tk.Label(top_frame, text="No file selected", width=50)
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # Main content frame
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Preview frame
        preview_frame = tk.LabelFrame(main_frame, text="PDF Preview")
        preview_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5))
        
        # Canvas for PDF preview
        self.canvas = tk.Canvas(preview_frame, bg='gray90')
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Preview controls
        control_frame = tk.Frame(preview_frame)
        control_frame.pack(pady=5)
        
        tk.Button(control_frame, text="◀ Previous", command=self.prev_page).pack(side=tk.LEFT, padx=2)
        self.page_label = tk.Label(control_frame, text="Page 1 of 1")
        self.page_label.pack(side=tk.LEFT, padx=10)
        tk.Button(control_frame, text="Next ▶", command=self.next_page).pack(side=tk.LEFT, padx=2)
        
        # Conversion frame
        convert_frame = tk.LabelFrame(main_frame, text="Conversion Settings", width=300)
        convert_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5,0))
        convert_frame.pack_propagate(False)
        
        # Settings
        tk.Label(convert_frame, text="Page Range:").pack(anchor=tk.W, pady=(15,0), padx=10)
        
        range_frame = tk.Frame(convert_frame)
        range_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.start_page = tk.Entry(range_frame, width=8)
        self.start_page.pack(side=tk.LEFT)
        tk.Label(range_frame, text=" to ").pack(side=tk.LEFT, padx=2)
        self.end_page = tk.Entry(range_frame, width=8)
        self.end_page.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(convert_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=10, pady=20)
        
        # Convert button
        tk.Button(convert_frame, text="Convert to DOCX", command=self.convert_pdf,
                 bg="#2196F3", fg="white", padx=30, pady=10).pack(pady=20)
        
        # Status label
        self.status_label = tk.Label(convert_frame, text="Ready", fg="gray")
        self.status_label.pack(pady=10)
        
    def select_pdf(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            self.pdf_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.load_pdf_preview(file_path)
            
    def load_pdf_preview(self, pdf_path):
        try:
            # Clear previous preview
            self.preview_images = []
            self.current_page = 0
            
            # Open PDF
            pdf_document = fitz.open(pdf_path)
            self.total_pages = len(pdf_document)
            
            # Convert each page to image
            for page_num in range(min(10, self.total_pages)):  # Limit to first 10 pages for performance
                page = pdf_document.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(150/72, 150/72))  # 150 DPI
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                self.preview_images.append(img)
            
            pdf_document.close()
            
            # Update page label
            self.update_page_display()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
            
    def update_page_display(self):
        if self.preview_images:
            # Display current page
            img = self.preview_images[self.current_page]
            img.thumbnail((550, 700), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            self.canvas.delete("all")
            self.canvas.create_image(10, 10, anchor=tk.NW, image=photo)
            self.canvas.image = photo  # Keep reference
            
            # Update page label
            self.page_label.config(text=f"Page {self.current_page + 1} of {len(self.preview_images)}")
            
            # Update range fields
            if not self.start_page.get():
                self.start_page.delete(0, tk.END)
                self.start_page.insert(0, "1")
            if not self.end_page.get():
                self.end_page.delete(0, tk.END)
                self.end_page.insert(0, str(self.total_pages))
    
    def next_page(self):
        if self.preview_images and self.current_page < len(self.preview_images) - 1:
            self.current_page += 1
            self.update_page_display()
    
    def prev_page(self):
        if self.preview_images and self.current_page > 0:
            self.current_page -= 1
            self.update_page_display()
    
    def convert_pdf(self):
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first!")
            return
        
        # Get save location
        save_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        
        if not save_path:
            return
        
        try:
            # Get page range
            start_page = int(self.start_page.get() or 1)
            end_page = int(self.end_page.get() or self.total_pages)
            
            # Validate range
            if start_page < 1 or end_page > self.total_pages or start_page > end_page:
                messagebox.showerror("Error", "Invalid page range!")
                return
            
            # Start conversion
            self.status_label.config(text="Converting...", fg="blue")
            self.progress.start()
            self.root.update()
            
            # Convert PDF to DOCX
            cv = Converter(self.pdf_path)
            cv.convert(save_path, start=start_page-1, end=end_page-1)
            cv.close()
            
            # Stop progress bar
            self.progress.stop()
            self.status_label.config(text="Conversion Complete!", fg="green")
            
            # Ask to open the file
            if messagebox.askyesno("Success", "Conversion completed successfully!\nOpen the document?"):
                os.startfile(save_path)
                
        except ValueError:
            messagebox.showerror("Error", "Please enter valid page numbers!")
            self.progress.stop()
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
            self.progress.stop()
            self.status_label.config(text="Conversion Failed", fg="red")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    # Install required packages first:
    # pip install pdf2docx PyMuPDF Pillow python-docx
    
    root = tk.Tk()
    app = PDFtoDOCXConverter(root)
    app.run()