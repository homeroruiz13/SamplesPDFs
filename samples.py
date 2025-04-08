import os
import fitz  # PyMuPDF for PDF manipulation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageEnhance
import tempfile
import shutil

# Define the correct footer file path
FOOTER_DIR = "/Users/homeroruiz/Downloads/Compound/LemonPark/"

# Define horizontal extension in points (convert from pixels to points)
# Bleed values: 2mm = 5.6693 points, 3mm = 8.5039 points
BLEED_2MM_POINTS = 5.6693
BLEED_3MM_POINTS = 8.5039

def enhance_image(image_path, contrast=1.2, brightness=1.1, sharpness=1.3):
    """
    Enhance image quality with adjustable parameters
    """
    img = Image.open(image_path).convert("RGB")
    
    # Apply enhancements
    img = ImageEnhance.Contrast(img).enhance(contrast)
    img = ImageEnhance.Brightness(img).enhance(brightness)
    img = ImageEnhance.Sharpness(img).enhance(sharpness)
    
    # Save enhanced image to a temporary file
    temp_path = f"temp_enhanced_{os.path.basename(image_path)}"
    img.save(temp_path, format="PNG", quality=100, dpi=(600, 600))
    
    return temp_path

def create_pdf(image_path, height_ft, substrate, width_ft=2, dpi=1200, 
              spacing_points=20, design_name=None, bleed_mm=2):
    """
    Create a tiled large-format PDF from an image, then overlay the correct footer at the bottom.
    Adds the design_name (or image filename if not provided) to the footer.
    
    Parameters:
    - bleed_mm: Bleed value in mm (2mm or 3mm)
    - substrate: One of "TRAD", "P&S", or "PP"
    """
    try:
        # Enhance the image first
        enhanced_image_path = enhance_image(image_path)
        
        # Convert feet to points (1 inch = 72 points, 1 foot = 12 inches)
        tile_width_points = width_ft * 12 * 72
        total_height_points = height_ft * 12 * 72
        
        # Set bleed value based on parameter
        bleed_points = BLEED_2MM_POINTS if bleed_mm == 2 else BLEED_3MM_POINTS
        
        # Add horizontal extension to each side (increasing tile width)
        extended_tile_width = tile_width_points + (2 * bleed_points)
        
        # Create output filename with bleed information
        bleed_label = f"{bleed_mm}mm"
        
        total_width_points = extended_tile_width
        output_pdf = f"temp_{height_ft}ft_{substrate}_{bleed_label}.pdf"

        # Open the enhanced image
        img = Image.open(enhanced_image_path).convert("RGB")
        img_width, img_height = img.size
        
        # Calculate scaling factor based on the extended tile width to eliminate white space
        # This ensures the image is scaled to fill the entire extended width
        scale_factor = extended_tile_width / img_width
        new_width = int(extended_tile_width)  # Convert to integer for PIL
        new_height = int(img_height * scale_factor)
        
        # Resize image using high-quality resampling
        img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        # Save the resized image with high quality settings
        temp_resized = f"temp_resized_{os.path.basename(image_path)}"
        img.save(temp_resized, format="PNG", quality=100, dpi=(dpi, dpi))

        # Calculate the number of times the image should be repeated vertically
        tile_count = (total_height_points // new_height) + 1

        # Create PDF with high DPI, using the extended width
        c = canvas.Canvas(output_pdf, pagesize=(total_width_points, total_height_points))
        c.setAuthor("Automated PDF Generator")
        c.setTitle(f"{substrate} {height_ft}ft {bleed_label}")
        c.setSubject(f"High-Quality Print for {design_name or os.path.basename(image_path)}")
        c.setKeywords(["large format", "high quality", "print", f"{bleed_mm}mm bleed"])

        # Use ImageReader for better quality rendering
        img_reader = ImageReader(temp_resized)
        
        y_position = 0
        for _ in range(tile_count):
            # Draw with best quality settings available
            # Place image at the edge (no x_offset needed since image is already sized correctly)
            c.drawImage(img_reader, 0, y_position, width=new_width, height=new_height, 
                        preserveAspectRatio=True, mask='auto')
            
            y_position += new_height  # Move up for the next tile

        # Set PDF metadata for better quality printing
        c.showPage()
        c.save()
        
        print(f"Temporary PDF saved to {output_pdf}")

        # Get design name (if not provided, use the image filename without path and extension)
        if design_name is None:
            design_name = os.path.splitext(os.path.basename(image_path))[0]
            
        # Overlay generic footer at the bottom with detailed information
        overlay_footer(output_pdf, height_ft, substrate, False, spacing_points, design_name, bleed_mm)
        
        # Clean up temporary files
        os.remove(enhanced_image_path)
        os.remove(temp_resized)

    except Exception as e:
        print(f"Error: {e}")

def overlay_footer(base_pdf_path, height_ft, substrate, double_blade=False, spacing_points=20, design_name=None, bleed_mm=2):
    """
    Overlay the appropriate footer (based on height) onto the generated base PDF at the bottom,
    add a logo in the middle, and add custom text for pattern name, substrate, and bleed.
    """
    try:
        # Use specific footer files based on height
        footer_name = f"LemonPark{height_ft}_Footer.pdf"
        footer_pdf_path = os.path.join(FOOTER_DIR, footer_name)
        
        # Path to logo file
        logo_pdf_path = os.path.join(FOOTER_DIR, "logo.pdf")
        
        if not os.path.exists(footer_pdf_path):
            print(f"Error: Footer file not found at {footer_pdf_path}")
            return
            
        if not os.path.exists(logo_pdf_path):
            print(f"Error: Logo file not found at {logo_pdf_path}")
            return

        # Open base PDF and footer PDF
        base_pdf = fitz.open(base_pdf_path)
        footer_pdf = fitz.open(footer_pdf_path)

        # Extract footer as a high-resolution image to preserve quality
        footer_page = footer_pdf[0]
        # Use higher resolution for 27ft footers
        resolution_factor = 6 if height_ft == 27 else 3
        footer_pixmap = footer_page.get_pixmap(matrix=fitz.Matrix(resolution_factor, resolution_factor), alpha=False)
        footer_image_path = f"temp_footer_{height_ft}ft.png"
        footer_pixmap.save(footer_image_path, output="png", jpg_quality=100)
        
        # Extract logo as a high-resolution image
        logo_pdf = fitz.open(logo_pdf_path)
        logo_page = logo_pdf[0]
        # Use higher resolution for 27ft pages
        logo_resolution = 6 if height_ft == 27 else 3
        logo_pixmap = logo_page.get_pixmap(matrix=fitz.Matrix(logo_resolution, logo_resolution), alpha=True)
        logo_image_path = "temp_logo.png"
        logo_pixmap.save(logo_image_path, output="png", jpg_quality=100)
        logo_pdf.close()

        # Use provided design_name or extract from file path if not provided
        if design_name is None:
            design_name = os.path.splitext(os.path.basename(base_pdf_path))[0]
            # Remove the temp_ prefix and format info for cleaner display
            design_name = design_name.replace("temp_", "").split("_")[0]

        for page_num in range(len(base_pdf)):
            base_page = base_pdf[page_num]
            base_rect = base_page.rect

            # Calculate extended tile width with bleed value
            bleed_points = BLEED_2MM_POINTS if bleed_mm == 2 else BLEED_3MM_POINTS
            extended_tile_width = base_rect.width
            
            # Calculate the width for the footer (should match the image width without overlay issues)
            footer_width = extended_tile_width
            footer_height = footer_pixmap.height * (footer_width / footer_pixmap.width)  # Maintain aspect ratio

            # Place footer at the very bottom of the page
            y1 = base_rect.height  # Bottom of the page
            y0 = y1 - footer_height  # Top of the footer
            
            # Prepare the detailed text to display (without size since it's in the footer)
            detail_text = f"{design_name} | {substrate} | {bleed_mm}mm"

            # Place footer at the same position as the image (no offset)
            base_page.insert_image(fitz.Rect(0, y0, footer_width, y1), 
                                 filename=footer_image_path)
            
            # Add detailed text - position to right side of footer
            text_x = footer_width * 0.75
            text_y = y1 - footer_height * 0.5  # Center in footer
            # Use larger font size for 27ft versions
            font_size = 18 if height_ft == 27 else 12
            add_text_to_page(base_page, detail_text, text_x, text_y, fontsize=font_size)
            
            # Calculate logo size and position (centered in footer)
            logo_width = footer_height * 1.5  # Make logo width proportional to footer height
            logo_height = footer_height * 0.7  # Leave some margin top and bottom
            logo_x = (footer_width - logo_width) / 2  # Center horizontally
            logo_y = y0 + (footer_height - logo_height) / 2  # Center vertically in footer
            
            # Place logo in the center of the footer (on top of the footer)
            base_page.insert_image(fitz.Rect(logo_x, logo_y, logo_x + logo_width, logo_y + logo_height), 
                                 filename=logo_image_path)

        # Generate final filename based on pattern: [image_name]_[substrate]_[size]_[bleed]
        output_name = design_name
        final_pdf_path = f"{output_name}_{substrate}_{height_ft}ft_{bleed_mm}mm.pdf"
        
        # Save with higher quality settings - use specialized settings for 27ft versions
        if height_ft == 27:
            # Use higher compression quality for 27ft files
            base_pdf.save(final_pdf_path, 
                         garbage=4,       # Maximum garbage collection
                         deflate=False,   # No deflate compression
                         clean=True,      # Clean unused objects
                         pretty=True,     # Pretty print
                         ascii=False,     # Use binary format for better quality
                         linear=True)     # Optimize for web viewing
        else:
            base_pdf.save(final_pdf_path, 
                         garbage=4,       # Maximum garbage collection
                         deflate=False,   # No deflate compression
                         clean=True)      # Clean unused objects
        
        base_pdf.close()
        footer_pdf.close()

        # Cleanup temporary images
        os.remove(footer_image_path)
        os.remove(logo_image_path)

        print(f"Final PDF with footer saved as {final_pdf_path}")

        # Clean up temporary file
        os.remove(base_pdf_path)

    except Exception as e:
        print(f"Error overlaying footer: {e}")
        
def add_text_to_page(page, text, x, y, fontname="Arial", fontsize=12):
    """
    Add text to a PDF page at specific coordinates using Arial font.
    """
    try:
        # Create text writer with specified font
        tw = fitz.TextWriter(page.rect)
        
        # Load the Arial font (or fallback to Helvetica if Arial isn't available)
        font = None
        try:
            font = fitz.Font(fontname)
        except:
            font = fitz.Font("helv")  # Helvetica as fallback
            print(f"Warning: Arial font not available, using Helvetica instead")
        
        # Add text with specified properties
        tw.append((x, y), text, font=font, fontsize=fontsize)
        
        # Write the text to the page with better rendering quality
        tw.write_text(page, color=(0, 0, 0))  # Black text for better readability
        
    except Exception as e:
        print(f"Error adding text to page: {e}")

if __name__ == "__main__":
    image_path = input("Enter the full path to the image file: ").strip()

    if not os.path.exists(image_path):
        print(f"Error: The specified image file '{image_path}' does not exist.")
    else:
        # Use the image filename as the design name
        design_name = os.path.splitext(os.path.basename(image_path))[0]
        
        # Process all combinations of parameters to create 6 panels
        # (2 lengths x 3 substrates) x 2 bleeds = 12 total pdfs, but no double blade
        
        # All substrates
        substrates = ["TRAD", "P&S", "PP"]
        
        # Heights
        heights = [13, 27]
        
        # Bleed values
        bleed_mm_values = [2, 3]
        
        # Generate only the single blade versions (6 panels)
        for substrate in substrates:
            for height in heights:
                for bleed_mm in bleed_mm_values:
                    # Create only single blade version
                    create_pdf(image_path, height_ft=height, substrate=substrate, 
                              design_name=design_name, bleed_mm=bleed_mm)
