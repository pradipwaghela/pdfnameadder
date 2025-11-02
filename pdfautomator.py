import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
import csv
import os
from pathlib import Path
import io
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from weasyprint import HTML
from PyPDF2 import PdfReader, PdfWriter
from io import BytesIO
# Try to import arabic_reshaper and bidi for proper text shaping
try:
    from PIL import ImageFont
    import PIL.features
    # Check if Pillow has harfbuzz support
    HAS_HARFBUZZ = PIL.features.check_feature('raqm')
except:
    HAS_HARFBUZZ = False

class InvitationNameAdder:
    def __init__(self, root):
        self.root = root
        self.root.title("Wedding Invitation Name Adder - ગુજરાતી")
        self.root.geometry("1200x800")
        
        self.pdf_path = None
        self.pdf_doc = None
        self.current_page = 0
        self.positions = []  # [(page, x, y, font_size), ...]
        self.zoom = 1.0
        self.csv_path = None
        self.font_path = None
        self.font_size = 20
        
        # Check text shaping support
        if HAS_HARFBUZZ:
            self.text_shaping_available = True
            print("✓ Advanced text shaping available (Harfbuzz/Raqm)")
        else:
            self.text_shaping_available = False
            print("⚠ Basic text rendering (conjuncts may not work perfectly)")
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Control Panel (Left) with Scrollbar
        control_container = ttk.Frame(main_frame)
        control_container.grid(row=0, column=0, rowspan=3, sticky=(tk.N, tk.S, tk.W), padx=(0, 10))
        
        # Canvas for scrolling
        control_canvas = tk.Canvas(control_container, width=280, bg='white')
        control_scrollbar = ttk.Scrollbar(control_container, orient="vertical", command=control_canvas.yview)
        control_frame = ttk.Frame(control_canvas)
        
        control_canvas.configure(yscrollcommand=control_scrollbar.set)
        
        control_scrollbar.pack(side=tk.RIGHT, fill='y')
        control_canvas.pack(side=tk.LEFT, fill='both', expand=True)
        
        canvas_frame = control_canvas.create_window((0, 0), window=control_frame, anchor='nw')
        
        def configure_scroll_region(event):
            control_canvas.configure(scrollregion=control_canvas.bbox("all"))
        
        control_frame.bind('<Configure>', configure_scroll_region)
        
        # Mouse wheel scrolling
        def on_mousewheel(event):
            control_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        control_canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # STEP 0: Font Selection
        ttk.Label(control_frame, text="STEP 0: Select Gujarati Font", font=('', 11, 'bold'), 
                 background='pink', padding=5).pack(fill='x', pady=(0,5))
        
        ttk.Button(control_frame, text="🔤 Select Gujarati Font (.ttf)", 
                   command=self.load_font, width=30).pack(pady=5)
        
        self.font_label = ttk.Label(control_frame, text="No font loaded\n(Required for Gujarati!)", 
                                   wraplength=250, foreground="red", font=('', 9, 'bold'))
        self.font_label.pack(pady=5)
        
        ttk.Label(control_frame, text="Download font from:\nfonts.google.com/noto", 
                 foreground="blue", font=('', 8), cursor="hand2").pack()
        
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # PDF Selection
        ttk.Label(control_frame, text="STEP 1: Load PDF", font=('', 11, 'bold'), 
                 background='lightblue', padding=5).pack(fill='x', pady=(0,5))
        ttk.Button(control_frame, text="📄 Select PDF Template", 
                   command=self.load_pdf, width=30).pack(pady=5)
        
        self.pdf_label = ttk.Label(control_frame, text="No PDF loaded", 
                                   wraplength=200, foreground="gray")
        self.pdf_label.pack(pady=5)
        
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Page Navigation
        ttk.Label(control_frame, text="STEP 2: Navigate & Click", font=('', 11, 'bold'),
                 background='lightgreen', padding=5).pack(fill='x', pady=(0,5))
        nav_frame = ttk.Frame(control_frame)
        nav_frame.pack(pady=5)
        
        ttk.Button(nav_frame, text="◀ Prev", command=self.prev_page, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(nav_frame, text="Next ▶", command=self.next_page, width=10).pack(side=tk.LEFT, padx=2)
        
        self.page_label = ttk.Label(control_frame, text="Page: 0/0")
        self.page_label.pack(pady=5)
        
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Font Settings
        ttk.Label(control_frame, text="Font Settings:").pack(anchor=tk.W)
        
        font_frame = ttk.Frame(control_frame)
        font_frame.pack(fill='x', pady=5)
        
        ttk.Label(font_frame, text="Size:").pack(anchor=tk.W)
        self.size_var = tk.IntVar(value=self.font_size)
        ttk.Spinbox(font_frame, from_=8, to=120, textvariable=self.size_var, 
                   width=22).pack(fill='x', pady=2)
        
        # Text rendering options
        ttk.Label(font_frame, text="Rendering:").pack(anchor=tk.W, pady=(5,0))
        self.rendering_var = tk.StringVar(value="normal")
        ttk.Radiobutton(font_frame, text="Normal", variable=self.rendering_var, 
                       value="normal").pack(anchor=tk.W)
        ttk.Radiobutton(font_frame, text="Better Quality (Slower)", variable=self.rendering_var, 
                       value="quality").pack(anchor=tk.W)
        
        # Color settings
        ttk.Label(font_frame, text="Text Color:").pack(anchor=tk.W, pady=(5,0))
        color_frame = ttk.Frame(font_frame)
        color_frame.pack(fill='x', pady=2)
        
        self.color_var = tk.StringVar(value="black")
        colors = [("Black", "black"), ("Red", "red"), ("Blue", "blue"), 
                  ("Gold", "gold"), ("Green", "green"), ("Maroon", "maroon")]
        
        for i, (label, color) in enumerate(colors):
            ttk.Radiobutton(color_frame, text=label, variable=self.color_var, 
                           value=color).grid(row=i//2, column=i%2, sticky=tk.W, padx=2)
        
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Instructions
        ttk.Label(control_frame, text="Instructions:", font=('', 10, 'bold')).pack(anchor=tk.W)
        instructions = ttk.Label(control_frame, 
                                text="1. Load Gujarati font (IMPORTANT!)\n"
                                     "2. Load PDF template\n"
                                     "3. Click on PDF to mark positions\n"
                                     "4. Load guest CSV file\n"
                                     "5. Test, then generate all",
                                justify=tk.LEFT, wraplength=200)
        instructions.pack(anchor=tk.W, pady=5)
        
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Positions List
        ttk.Label(control_frame, text="STEP 3: Manage Positions", font=('', 11, 'bold'),
                 background='lightyellow', padding=5).pack(fill='x', pady=(0,5))
        ttk.Label(control_frame, text="Added Positions:").pack(anchor=tk.W)
        
        positions_frame = ttk.Frame(control_frame)
        positions_frame.pack(fill='both', expand=True, pady=5)
        
        self.positions_listbox = tk.Listbox(positions_frame, height=6, width=30)
        self.positions_listbox.pack(side=tk.LEFT, fill='both', expand=True)
        
        scrollbar = ttk.Scrollbar(positions_frame, orient="vertical", 
                                 command=self.positions_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill='y')
        self.positions_listbox.configure(yscrollcommand=scrollbar.set)
        
        btn_frame = ttk.Frame(control_frame)
        btn_frame.pack(fill='x', pady=5)
        ttk.Button(btn_frame, text="❌ Remove Selected Position", 
                  command=self.remove_position, width=30).pack()
        
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # CSV and Generate
        ttk.Label(control_frame, text="STEP 4: Load Guest List", font=('', 11, 'bold'),
                 background='lightcoral', padding=5).pack(fill='x', pady=(0,5))
        ttk.Button(control_frame, text="📋 Load Guest CSV File", 
                  command=self.load_csv, width=30).pack(pady=5)
        
        self.csv_label = ttk.Label(control_frame, text="No CSV loaded", 
                                   wraplength=200, foreground="gray")
        self.csv_label.pack(pady=5)
        
        ttk.Separator(control_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Generate Button - BIG and PROMINENT
        ttk.Label(control_frame, text="STEP 5: Generate!", font=('', 11, 'bold'),
                 background='lightgreen', padding=5).pack(fill='x', pady=(0,5))
        
        # Test with sample name first
        ttk.Button(control_frame, text="🧪 Test with Sample Name", 
                  command=self.test_sample, width=30).pack(pady=5)
        
        generate_btn = ttk.Button(control_frame, text="🎉 GENERATE ALL INVITATIONS 🎉", 
                                 command=self.generate_invitations)
        generate_btn.pack(pady=10, ipady=10, fill='x')
        
        # Add some padding at bottom for scrolling
        ttk.Label(control_frame, text=" ").pack(pady=20)
        
        # PDF Display Area (Right)
        display_frame = ttk.LabelFrame(main_frame, text="PDF Preview (Click to add name position)", 
                                      padding="10")
        display_frame.grid(row=0, column=1, rowspan=3, sticky=(tk.N, tk.S, tk.E, tk.W))
        display_frame.columnconfigure(0, weight=1)
        display_frame.rowconfigure(0, weight=1)
        
        # Canvas with scrollbars
        canvas_frame = ttk.Frame(display_frame)
        canvas_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)
        
        self.canvas = tk.Canvas(canvas_frame, bg='gray', cursor='crosshair')
        self.canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.E, tk.W))
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.canvas.bind('<Button-1>', self.on_canvas_click)
        
        # Zoom controls
        zoom_frame = ttk.Frame(display_frame)
        zoom_frame.grid(row=1, column=0, pady=5)
        
        ttk.Button(zoom_frame, text="🔍−", command=self.zoom_out, width=5).pack(side=tk.LEFT, padx=2)
        self.zoom_label = ttk.Label(zoom_frame, text="100%")
        self.zoom_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(zoom_frame, text="🔍+", command=self.zoom_in, width=5).pack(side=tk.LEFT, padx=2)
    
    def load_font(self):
        path = filedialog.askopenfilename(
            title="Select Gujarati Font File (.ttf)",
            filetypes=[("TrueType Font", "*.ttf"), ("All files", "*.*")]
        )
        
        if path:
            try:
                # Test if font can be loaded
                test_font = ImageFont.truetype(path, 20)
                self.font_path = path
                self.font_label.config(
                    text=f"✓ Loaded: {Path(path).name}", 
                    foreground="green"
                )
                messagebox.showinfo("Success", "Gujarati font loaded successfully!\n\nYou can now proceed with adding positions.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load font:\n{str(e)}")
    
    def load_pdf(self):
        path = filedialog.askopenfilename(
            title="Select Wedding Invitation PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if path:
            try:
                self.pdf_path = path
                self.pdf_doc = fitz.open(path)
                self.current_page = 0
                self.positions = []
                self.positions_listbox.delete(0, tk.END)
                
                self.pdf_label.config(text=f"Loaded: {Path(path).name}", foreground="green")
                self.display_page()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PDF:\n{str(e)}")
    
    def display_page(self):
        if not self.pdf_doc:
            return
        
        page = self.pdf_doc[self.current_page]
        
        # Render page to image
        mat = fitz.Matrix(self.zoom * 2, self.zoom * 2)  # 2x for better quality
        pix = page.get_pixmap(matrix=mat)
        
        # Convert to PIL Image then to PhotoImage
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.photo = ImageTk.PhotoImage(img)
        
        # Update canvas
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
        
        # Draw markers for positions on current page
        for pos in self.positions:
            if pos[0] == self.current_page:
                x, y = pos[1] * self.zoom * 2, pos[2] * self.zoom * 2
                self.canvas.create_oval(x-10, y-10, x+10, y+10, 
                                       fill='red', outline='white', width=2)
                self.canvas.create_text(x, y-20, text="📝", font=('', 16))
        
        # Update page label
        self.page_label.config(text=f"Page: {self.current_page + 1}/{len(self.pdf_doc)}")
    
    def on_canvas_click(self, event):
        if not self.pdf_doc:
            return
        
        if not self.font_path:
            messagebox.showwarning("Font Required", 
                "Please load a Gujarati font first (STEP 0)!\n\n"
                "Download Noto Sans Gujarati from:\nfonts.google.com/noto/specimen/Noto+Sans+Gujarati")
            return
        
        # Convert canvas coordinates to PDF coordinates
        canvas_x = self.canvas.canvasx(event.x)
        canvas_y = self.canvas.canvasy(event.y)
        
        pdf_x = canvas_x / (self.zoom * 2)
        pdf_y = canvas_y / (self.zoom * 2)
        
        # Add position
        font_size = self.size_var.get()
        self.positions.append((self.current_page, pdf_x, pdf_y, font_size))
        
        # Update listbox
        self.positions_listbox.insert(tk.END, 
            f"Page {self.current_page + 1}: ({int(pdf_x)}, {int(pdf_y)}) Size: {font_size}")
        
        # Redraw to show marker
        self.display_page()
        
        messagebox.showinfo("Position Added", 
                          f"Name position added on page {self.current_page + 1}")
    
    def remove_position(self):
        selection = self.positions_listbox.curselection()
        if selection:
            idx = selection[0]
            del self.positions[idx]
            self.positions_listbox.delete(idx)
            self.display_page()
    
    def prev_page(self):
        if self.pdf_doc and self.current_page > 0:
            self.current_page -= 1
            self.display_page()
    
    def next_page(self):
        if self.pdf_doc and self.current_page < len(self.pdf_doc) - 1:
            self.current_page += 1
            self.display_page()
    
    def zoom_in(self):
        self.zoom = min(self.zoom + 0.2, 3.0)
        self.zoom_label.config(text=f"{int(self.zoom * 100)}%")
        self.display_page()
    
    def zoom_out(self):
        self.zoom = max(self.zoom - 0.2, 0.5)
        self.zoom_label.config(text=f"{int(self.zoom * 100)}%")
        self.display_page()
    
    def load_csv(self):
        path = filedialog.askopenfilename(
            title="Select Guest Names CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if path:
            try:
                # Validate CSV
                with open(path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    headers = next(reader)
                    row_count = sum(1 for row in reader)
                
                self.csv_path = path
                self.csv_label.config(
                    text=f"Loaded: {Path(path).name}\n({row_count} guests)", 
                    foreground="green"
                )
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV:\n{str(e)}")
    
    # def add_text_to_pdf_page(self, pdf_page, text, x, y, font_size, color_rgb):
    #     """Add text to PDF using image overlay method with proper text shaping for Gujarati"""
    #     # Get page dimensions
    #     page_rect = pdf_page.rect
    #     page_width = page_rect.width
    #     page_height = page_rect.height
        
    #     # Increase resolution for better quality
    #     scale = 3 if self.rendering_var.get() == "quality" else 2
        
    #     # Create transparent image for text with higher resolution
    #     img_width = int(page_width * scale)
    #     img_height = int(page_height * scale)
    #     text_img = Image.new('RGBA', (img_width, img_height), (255, 255, 255, 0))
    #     draw = ImageDraw.Draw(text_img)
        
    #     # Load font with scaled size
    #     try:
    #         font = ImageFont.truetype(self.font_path, font_size * scale, layout_engine=ImageFont.Layout.RAQM)
    #     except:
    #         # Fallback to basic layout
    #         font = ImageFont.truetype(self.font_path, font_size * scale)
        
    #     # Draw text at scaled position with proper text shaping
    #     try:
    #         # Use proper text layout for complex scripts
    #         draw.text((x * scale, y * scale), text, font=font, fill=color_rgb, 
    #                  features=['-liga', '-clig'])  # Enable ligatures
    #     except:
    #         # Fallback to basic drawing
    #         draw.text((x * scale, y * scale), text, font=font, fill=color_rgb)
        
    #     # Resize back to original dimensions with high-quality resampling
    #     text_img = text_img.resize((int(page_width), int(page_height)), Image.Resampling.LANCZOS)
        
    #     # Convert PIL image to bytes
    #     img_buffer = io.BytesIO()
    #     text_img.save(img_buffer, format='PNG')
    #     img_buffer.seek(0)
        
    #     # Insert image into PDF
    #     pdf_page.insert_image(page_rect, stream=img_buffer.getvalue(), overlay=True)

    def add_text_to_pdf_page(self, pdf_page, text, x, y, font_size, color_rgb):
        """
        Add Gujarati text to PDF correctly using WeasyPrint (handles shaping like 'શ્રી').
        Works directly with PyMuPDF (fitz) page object.
        """
        # Handle RGB or RGBA
        if len(color_rgb) == 4:
            r, g, b, a = color_rgb
            a = a / 255.0
            color_css = f"rgba({r},{g},{b},{a})"
        else:
            r, g, b = color_rgb
            color_css = f"rgb({r},{g},{b})"

        # Get page size
        page_rect = pdf_page.rect
        page_width = int(page_rect.width)
        page_height = int(page_rect.height)

        # Build HTML for WeasyPrint
        html_content = f"""
        <html>
        <head>
            <meta charset="utf-8" />
            <style>
                @font-face {{
                    font-family: "GujaratiFont";
                    src: url("file://{self.font_path}");
                }}
                html, body {{
                    margin: 0; padding: 0;
                    width: {page_width}px; height: {page_height}px;
                    background: transparent;
                }}
                .text {{
                    position: absolute;
                    left: {x}px;
                    bottom: {y}px;
                    font-family: "GujaratiFont", sans-serif;
                    font-size: {font_size}px;
                    color: {color_css};
                    white-space: pre;
                }}
            </style>
        </head>
        <body>
            <div class="text">{text}</div>
        </body>
        </html>
        """

        # Render HTML -> PDF (WeasyPrint)
        pdf_buffer = io.BytesIO()
        HTML(string=html_content).write_pdf(pdf_buffer)
        pdf_buffer.seek(0)

        # Open the overlay as PyMuPDF document
        overlay_doc = fitz.open(stream=pdf_buffer.getvalue(), filetype="pdf")
        overlay_page = overlay_doc.load_page(0)

        # Render overlay page to PNG (to preserve complex text correctly)
        pix = overlay_page.get_pixmap(alpha=True)
        img_bytes = pix.tobytes("png")

        # Insert PNG image onto existing PDF page
        pdf_page.insert_image(
            pdf_page.rect,  # full-page overlay
            stream=img_bytes,  # PNG data
            keep_proportion=False,
            overlay=True
        )

        overlay_doc.close()


    
    def test_sample(self):
        """Test with a sample name to check font rendering"""
        if not self.pdf_path:
            messagebox.showerror("Error", "Please load a PDF template first!")
            return
        
        if not self.font_path:
            messagebox.showerror("Error", "Please load a Gujarati font first!")
            return
        
        if not self.positions:
            messagebox.showerror("Error", "Please add at least one name position!")
            return
        
        # Ask for sample name
        test_window = tk.Toplevel(self.root)
        test_window.title("Test Sample Name")
        test_window.geometry("400x150")
        test_window.transient(self.root)
        test_window.grab_set()
        
        ttk.Label(test_window, text="Enter sample name in Gujarati:", 
                 font=('', 11)).pack(pady=10)
        
        sample_entry = ttk.Entry(test_window, width=40, font=('', 12))
        sample_entry.pack(pady=10)
        sample_entry.insert(0, "શ્રી રાજેશભાઈ પટેલ")
        
        def generate_test():
            sample_name = sample_entry.get()
            if not sample_name:
                messagebox.showwarning("Warning", "Please enter a sample name!")
                return
            
            try:
                # Select output location
                output_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf")],
                    initialfile="test_invitation.pdf"
                )
                
                if not output_path:
                    return
                
                # Create test PDF
                doc = fitz.open(self.pdf_path)
                
                color_name = self.color_var.get()
                color_map = {
                    'black': (0, 0, 0, 255),
                    'red': (255, 0, 0, 255),
                    'blue': (0, 0, 255, 255),
                    'gold': (255, 215, 0, 255),
                    'green': (0, 128, 0, 255),
                    'maroon': (128, 0, 0, 255)
                }
                text_color = color_map.get(color_name, (0, 0, 0, 255))
                
                # Add name to each position
                for pos in self.positions:
                    page_num, x, y, size = pos
                    page = doc[page_num]
                    self.add_text_to_pdf_page(page, sample_name, x, y, size, text_color)
                
                doc.save(output_path)
                doc.close()
                
                test_window.destroy()
                
                result = messagebox.askyesno("Test Complete", 
                    f"Test invitation saved to:\n{output_path}\n\n"
                    "Please open and check if the Gujarati text displays correctly.\n\n"
                    "Does the text look good?")
                
                if result:
                    messagebox.showinfo("Great!", 
                        "Perfect! Now you can generate all invitations with confidence.")
                else:
                    messagebox.showinfo("Adjustment Needed", 
                        "Try adjusting:\n"
                        "1. Font size (in Font Settings)\n"
                        "2. Position (remove and re-add)\n"
                        "3. Different font file\n"
                        "Then test again!")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate test:\n{str(e)}")
        
        ttk.Button(test_window, text="Generate Test PDF", 
                  command=generate_test).pack(pady=10)
    
    def generate_invitations(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please load a PDF template first!")
            return
        
        if not self.font_path:
            messagebox.showerror("Error", "Please load a Gujarati font first!")
            return
        
        if not self.csv_path:
            messagebox.showerror("Error", "Please load a guest CSV file first!")
            return
        
        if not self.positions:
            messagebox.showerror("Error", "Please add at least one name position!")
            return
        
        # Confirm before generating
        confirm = messagebox.askyesno("Confirm Generation",
            "Have you tested with a sample name and verified it looks good?\n\n"
            "Click Yes to proceed with generating all invitations.")
        
        if not confirm:
            return
        
        # Select output directory
        output_dir = filedialog.askdirectory(title="Select Output Directory")
        if not output_dir:
            return
        
        try:
            # Read guest names
            with open(self.csv_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                guests = list(reader)
            
            if not guests:
                messagebox.showerror("Error", "CSV file is empty!")
                return
            
            # Check for 'name' column
            if 'name' not in guests[0]:
                messagebox.showerror("Error", 
                    "CSV must have a 'name' column!\n\nExample CSV format:\nname\nશ્રી રાજેશભાઈ\nશ્રીમતી સીતાબેન")
                return
            
            # Generate invitations
            progress = tk.Toplevel(self.root)
            progress.title("Generating Invitations")
            progress.geometry("400x150")
            progress.transient(self.root)
            progress.grab_set()
            
            ttk.Label(progress, text="Generating invitations...", 
                     font=('', 12)).pack(pady=20)
            
            progress_var = tk.DoubleVar()
            progress_bar = ttk.Progressbar(progress, variable=progress_var, 
                                          maximum=len(guests), length=300)
            progress_bar.pack(pady=10)
            
            status_label = ttk.Label(progress, text="")
            status_label.pack(pady=5)
            
            color_name = self.color_var.get()
            color_map = {
                'black': (0, 0, 0, 255),
                'red': (255, 0, 0, 255),
                'blue': (0, 0, 255, 255),
                'gold': (255, 215, 0, 255),
                'green': (0, 128, 0, 255),
                'maroon': (128, 0, 0, 255)
            }
            text_color = color_map.get(color_name, (0, 0, 0, 255))
            
            for i, guest in enumerate(guests):
                guest_name = guest['name']
                status_label.config(text=f"Processing: {guest_name}")
                progress.update()
                
                # Create new PDF
                doc = fitz.open(self.pdf_path)
                
                # Add name to each position
                for pos in self.positions:
                    page_num, x, y, size = pos
                    page = doc[page_num]
                    self.add_text_to_pdf_page(page, guest_name, x, y, size, text_color)
                
                # Save
                safe_name = "".join(c for c in guest_name if c.isalnum() or c in (' ', '_', '-'))
                output_path = os.path.join(output_dir, f"invitation_{safe_name}.pdf")
                doc.save(output_path)
                doc.close()
                
                progress_var.set(i + 1)
            
            progress.destroy()
            
            messagebox.showinfo("Success", 
                f"✅ Generated {len(guests)} invitations!\n\nSaved to: {output_dir}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate invitations:\n{str(e)}")

def main():
    root = tk.Tk()
    app = InvitationNameAdder(root)
    root.mainloop()

if __name__ == "__main__":
    main()