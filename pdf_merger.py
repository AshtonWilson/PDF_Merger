import os
import sys
import logging
from io import BytesIO
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from PIL import Image
import tempfile

# Get the application path for resources
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    application_path = os.path.dirname(sys.executable)
else:
    # Running as script
    application_path = os.path.dirname(os.path.abspath(__file__))

# Set up logging with a log file in the application directory
log_file = os.path.join(application_path, 'pdf_merger.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# GUI setup
try:
    import tkinter as tk
    from tkinter import filedialog, simpledialog, messagebox
    HAS_GUI = True
except ImportError:
    HAS_GUI = False
    logger.warning("Tkinter not available. Running in command-line mode.")

def pick_files_and_title():
    """Use GUI to select files and title if available"""
    if not HAS_GUI:
        logger.error("No GUI available. Exiting.")
        sys.exit(1)
        
    root = tk.Tk()
    root.withdraw()
    
    main_pdf = filedialog.askopenfilename(
        title="Select Main Report PDF", 
        filetypes=[("PDF files", "*.pdf")]
    )
    if not main_pdf:
        return None, None, None
        
    trial_pdfs = filedialog.askopenfilenames(
        title="Select one or more trial-report PDFs", 
        filetypes=[("PDF files", "*.pdf")]
    )
    if not trial_pdfs:
        return None, None, None
        
    report_title = simpledialog.askstring(
        "Report Title", 
        "Enter the report title to display on each page:"
    )
    
    return main_pdf, list(trial_pdfs), report_title

def get_page_count(pdf_path):
    """Get the number of pages in a PDF file"""
    try:
        with open(pdf_path, 'rb') as f:
            reader = PdfReader(f)
            return len(reader.pages)
    except Exception as e:
        logger.error(f"Error getting page count for {pdf_path}: {str(e)}")
        return 0

def calculate_total_pages(main_pdf, trial_pdfs):
    """Calculate total pages in the final document"""
    main_pages = get_page_count(main_pdf)
    total = main_pages
    
    # Add pages for each trial PDF plus one cover page each
    for pdf_path in trial_pdfs:
        trial_pages = get_page_count(pdf_path)
        total += trial_pages + 1  # +1 for cover page
    
    logger.info(f"Total calculated pages: {total}")
    return total

def create_page_number_overlay(page_num, total_pages, report_title=""):
    """Create a footer overlay with page numbers"""
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    w, h = letter
    
    # Create white box to cover existing page numbers at the bottom
    c.setFillColorRGB(1, 1, 1)  # White
    c.rect(0, 0, w, 0.7 * inch, fill=1, stroke=0)
    
    # Add footer text
    c.setFont("Helvetica", 10)  # Use consistent font
    c.setFillColorRGB(0, 0, 0)  # Black
    
    # Draw the report title on the left
    if report_title:
        c.drawString(0.5 * inch, 0.25 * inch, report_title)
    
    # Draw the page number centered at the bottom
    page_text = f"Page {page_num} of {total_pages}"
    text_width = c.stringWidth(page_text, "Helvetica", 10)
    c.drawString((w - text_width) / 2, 0.25 * inch, page_text)
    
    # Add a line above the footer
    c.setStrokeColorRGB(0.5, 0.5, 0.5)  # Gray color
    c.line(0.5 * inch, 0.5 * inch, w - 0.5 * inch, 0.5 * inch)
    
    c.save()
    packet.seek(0)
    return PdfReader(packet)

def create_cover_page(title, page_num, total_pages, report_title=""):
    """Create a cover page for each trial PDF"""
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=letter)
    w, h = letter
    
    # Add title to center of page
    c.setFont("Helvetica-Bold", 24)
    title_width = c.stringWidth(title, "Helvetica-Bold", 24)
    c.drawString((w - title_width) / 2, h / 2, title)
    
    # Create white box to cover existing page numbers at the bottom
    c.setFillColorRGB(1, 1, 1)  # White
    c.rect(0, 0, w, 0.7 * inch, fill=1, stroke=0)
    
    # Add footer text
    c.setFont("Helvetica", 10)  # Use consistent font
    c.setFillColorRGB(0, 0, 0)  # Black
    
    # Draw the report title on the left
    if report_title:
        c.drawString(0.5 * inch, 0.25 * inch, report_title)
    
    # Draw the page number centered at the bottom
    page_text = f"Page {page_num} of {total_pages}"
    text_width = c.stringWidth(page_text, "Helvetica", 10)
    c.drawString((w - text_width) / 2, 0.25 * inch, page_text)
    
    # Add a line above the footer
    c.setStrokeColorRGB(0.5, 0.5, 0.5)  # Gray color
    c.line(0.5 * inch, 0.5 * inch, w - 0.5 * inch, 0.5 * inch)
    
    c.save()
    packet.seek(0)
    return PdfReader(packet)

def create_final_report(main_pdf, trial_pdfs, output_path, report_title=""):
    """Create the final report by merging PDFs and adding page numbers"""
    logger.info("Starting PDF assembly process")
    
    # Calculate total pages
    total_pages = calculate_total_pages(main_pdf, trial_pdfs)
    
    # Create a PDF writer for the final document
    writer = PdfWriter()
    current_page = 1
    
    # Process main PDF
    logger.info(f"Processing main PDF: {os.path.basename(main_pdf)}")
    try:
        reader = PdfReader(main_pdf)
        for i, page in enumerate(reader.pages):
            # Create footer overlay with page number
            overlay = create_page_number_overlay(current_page, total_pages, report_title)
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
            logger.info(f"Added main PDF page {i+1} as page {current_page}")
            current_page += 1
    except Exception as e:
        logger.error(f"Error processing main PDF: {str(e)}")
        if HAS_GUI:
            messagebox.showerror("Error", f"Error processing main PDF: {str(e)}")
        return None
    
    # Process each trial PDF
    for i, pdf_path in enumerate(trial_pdfs):
        title = os.path.splitext(os.path.basename(pdf_path))[0]
        logger.info(f"Processing trial PDF {i+1}: {title}")
        
        try:
            # Add cover page
            cover_page = create_cover_page(title, current_page, total_pages, report_title)
            writer.add_page(cover_page.pages[0])
            logger.info(f"Added cover page as page {current_page}")
            current_page += 1
            
            # Process the trial PDF
            reader = PdfReader(pdf_path)
            for j, page in enumerate(reader.pages):
                # Create footer overlay with page number
                overlay = create_page_number_overlay(current_page, total_pages, report_title)
                page.merge_page(overlay.pages[0])
                writer.add_page(page)
                logger.info(f"Added trial PDF {i+1} page {j+1} as page {current_page}")
                current_page += 1
        except Exception as e:
            logger.error(f"Error processing trial PDF {title}: {str(e)}")
            if HAS_GUI:
                messagebox.showerror("Error", f"Error processing trial PDF {title}: {str(e)}")
    
    # Write the final PDF
    logger.info(f"Writing final PDF to {output_path}")
    try:
        with open(output_path, "wb") as f:
            writer.write(f)
        
        logger.info(f"PDF assembly complete. Output: {output_path}")
        return output_path
    except Exception as e:
        logger.error(f"Error writing final PDF: {str(e)}")
        if HAS_GUI:
            messagebox.showerror("Error", f"Error writing final PDF: {str(e)}")
        return None

def main():
    try:
        # Show a welcome message if GUI is available
        if HAS_GUI:
            root = tk.Tk()
            root.title("PDF Merger Tool")
            
            # Center the window
            window_width = 400
            window_height = 200
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            x = (screen_width // 2) - (window_width // 2)
            y = (screen_height // 2) - (window_height // 2)
            root.geometry(f"{window_width}x{window_height}+{x}+{y}")
            
            # Welcome message
            tk.Label(root, text="PDF Merger Tool", font=("Arial", 16)).pack(pady=20)
            tk.Label(root, text="This tool will help you merge PDFs with custom covers and page numbers").pack(pady=10)
            
            def start_process():
                root.destroy()
                process_files()
            
            tk.Button(root, text="Start", command=start_process, width=20).pack(pady=20)
            
            root.mainloop()
        else:
            process_files()
            
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        if HAS_GUI:
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
        sys.exit(1)

def process_files():
    """Main processing function separated for better GUI integration"""
    main_pdf, trial_pdfs, report_title = pick_files_and_title()
    if not main_pdf or not trial_pdfs:
        logger.error("No files selected")
        return
    
    # Create output file path
    base, ext = os.path.splitext(main_pdf)
    output_pdf = f"{base}_WithCovers.pdf"
    
    # Let user select output location if GUI available
    if HAS_GUI:
        output_pdf = filedialog.asksaveasfilename(
            title="Save Final PDF As",
            initialfile=os.path.basename(output_pdf),
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if not output_pdf:
            logger.info("Output file selection cancelled")
            return
    
    logger.info(f"Main PDF: {main_pdf}")
    logger.info(f"Trial PDFs: {[os.path.basename(p) for p in trial_pdfs]}")
    logger.info(f"Report title: {report_title}")
    
    # Create the final report
    result = create_final_report(main_pdf, trial_pdfs, output_pdf, report_title)
    
    if result and HAS_GUI:
        if messagebox.askyesno("Success", f"PDF successfully created at:\n{output_pdf}\n\nWould you like to open it now?"):
            try:
                os.startfile(output_pdf)  # Windows
            except AttributeError:
                try:
                    import subprocess
                    subprocess.call(('open', output_pdf))  # macOS
                except:
                    try:
                        subprocess.call(('xdg-open', output_pdf))  # Linux
                    except:
                        messagebox.showinfo("Info", "Could not open the PDF automatically. Please open it manually.")

if __name__ == "__main__":
    main()
