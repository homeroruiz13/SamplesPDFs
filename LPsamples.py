import os
import fitz  # PyMuPDF for PDF manipulation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib import colors
from PIL import Image, ImageEnhance
import tempfile
import shutil

# Define the updated footer file path based on requested change
FOOTER_DIR = "C:\\Users\\plome\\Downloads\\Compound\\Lemon-park\\"
FOOTER_FILE = "LPfooter.pdf"

# Define font information
FONT_NAME = "AcuminPro"
FONT_SIZE = 12

# Try to register Acumin Pro font if available
try:
    # Path to the font file - using specified path
    font_path = "C:\\Users\\plome\\Downloads\\Compound\\Painted-paper\\Acumin-RPro.otf"
    pdfmetrics.registerFont(TTFont(FONT_NAME, font_path))
except:
    print("Warning: Acumin Pro font not found. Using Helvetica as fallback.")
    FONT_NAME = "Helvetica"  # Fallback to a standard font

# Define horizontal extension in points (convert from pixels to points)
# Bleed values: 2mm = 5.6693 points, 3mm = 8.5039 points
BLEED_2MM_POINTS = 5.6693
BLEED_3MM_POINTS = 8.5039

def optimize_raster_footer(footer_pdf_path, output_path=None, upscale_factor=4, sharpness_factor=1.2):
    """
    Optimize a raster-based footer for higher quality output.
    Allows setting a custom upscale factor and sharpness.
    """
    if output_path is None:
        output_path = footer_pdf_path.replace(".pdf", "_hq.pdf")

    # Extract the raster image from the PDF at highest resolution
    doc = fitz.open(footer_pdf_path)
    page = doc[0]

    # Get images from the page with high resolution
    images = page.get_images(full=True)

    if not images:
        print("No images found in footer PDF")
        return footer_pdf_path

    # Get the first image (assuming there is only one main image)
    img_index = 0
    xref = images[img_index][0]

    # Extract the image
    base_img = doc.extract_image(xref)
    img_data = base_img["image"]
    img_ext = base_img["ext"]

    # Save the image temporarily with highest quality
    temp_img_path = f"temp_footer_img.{img_ext}"
    with open(temp_img_path, "wb") as img_file:
        img_file.write(img_data)

    # Now upscale and sharpen the image
    from PIL import Image
    img = Image.open(temp_img_path)
    width, height = img.size

    new_width = int(width * upscale_factor)
    new_height = int(height * upscale_factor)
    img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

    # Apply sharpening
    enhancer = ImageEnhance.Sharpness(img)
    img = enhancer.enhance(sharpness_factor)

    # Save with high quality
    hq_temp_path = "temp_footer_hq.png"
    img.save(hq_temp_path, format="PNG", quality=100, dpi=(1200, 1200))

    # Create a new PDF with this high-quality image
    new_pdf = fitz.open()
    new_page = new_pdf.new_page(width=page.rect.width, height=page.rect.height)

    # Insert the high-quality image
    new_page.insert_image(page.rect, filename=hq_temp_path)

    # Save the new PDF with the high-quality raster
    new_pdf.save(output_path, garbage=4, deflate=True, clean=True)
    new_pdf.close()
    doc.close()

    # Clean up temp files
    os.remove(temp_img_path)
    os.remove(hq_temp_path)

    print(f"[✅] Optimized raster footer (upscaled x{upscale_factor}, sharpness {sharpness_factor}) saved to {output_path}")
    return output_path

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
               spacing_points=20, design_name=None, bleed_mm=2, footer_upscale=4, footer_sharpness=1.2):
    """
    Create a tiled large-format PDF from an image, then overlay the correct footer at the bottom.
    Adds the design_name (or image filename if not provided) to the footer.

    Parameters:
    - bleed_mm: Bleed value in mm (2mm or 3mm)
    - substrate: One of "TRAD", "P&S", or "PP"
    - footer_upscale: Upscale factor for the footer optimization (default: 4)
    - footer_sharpness: Sharpness factor for the footer optimization (default: 1.2)
    """
    try:
        # Check image resolution and provide warnings
        img_check = Image.open(image_path)
        img_width, img_height = img_check.size
        print(f"🔍 Checking image resolution: {img_width}x{img_height}")

        # Calculate required resolution for quality printing (1200 DPI)
        required_width = int(width_ft * 12 * dpi)  # Convert feet to inches, then to pixels at dpi
        required_height = int(height_ft * 12 * dpi)
        print(f"🧮 Required for {width_ft}ft x {height_ft}ft at {dpi} DPI: {required_width}x{required_height}")

        if img_width < required_width or img_height < required_height:
            print(f"⚠️ WARNING: Input image is too small and will be upscaled, resulting in reduced quality.")

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

        print(f"[✅] Base panel saved: {output_pdf}")

        # Get design name (if not provided, use the image filename without path and extension)
        if design_name is None:
            design_name = os.path.splitext(os.path.basename(image_path))[0]

        # Overlay footer
        overlay_footer(output_pdf, height_ft, substrate, False, spacing_points, design_name, bleed_mm,
                       footer_upscale=footer_upscale, footer_sharpness=footer_sharpness)

        # Clean up temporary files
        os.remove(enhanced_image_path)
        os.remove(temp_resized)

    except Exception as e:
        print(f"Error: {e}")

def overlay_footer(base_pdf_path, height_ft, substrate, double_blade=False, spacing_points=20, design_name=None, bleed_mm=2, footer_upscale=4, footer_sharpness=1.2):
    """
    Overlay the appropriate footer onto the generated base PDF at the bottom.
    Uses specialized approach for 27ft panels to maintain highest possible quality.
    Allows setting upscale and sharpness for footer optimization.
    Also adds text information about design name, material, and panel height.
    Text is positioned on the opposite side without labels.
    """
    try:
        # Use the specified footer file from Lemon-park instead of Painted-paper
        footer_pdf_path = os.path.join(FOOTER_DIR, FOOTER_FILE)

        if not os.path.exists(footer_pdf_path):
            print(f"Error: Footer file not found at {footer_pdf_path}")
            return

        # Open base PDF
        base_pdf = fitz.open(base_pdf_path)
        base_page = base_pdf[0]
        base_rect = base_page.rect

        # Get dimensions
        pdf_width = base_rect.width
        pdf_height = base_rect.height

        # Use provided design_name or extract from file path if not provided
        if design_name is None:
            design_name = os.path.splitext(os.path.basename(base_pdf_path))[0]
            # Remove the temp_ prefix and format info for cleaner display
            design_name = design_name.replace("temp_", "").split("_")[0]

        # Generate final filename based on pattern: [image_name]_[substrate]_[size]_[bleed]
        output_name = design_name
        final_pdf_path = f"{output_name}_{substrate}_{height_ft}ft_{bleed_mm}mm.pdf"

        # Get the full material name based on substrate code
        material_name = {
            "TRAD": "Traditional",
            "P&S": "Peel & Stick",
            "PP": "Pre-Pasted"
        }.get(substrate, substrate)
        
        # Panel height text with ft" format
        panel_height_text = f"{height_ft}ft\""

        # Special handling for 27ft panels - use PDF merging approach for best quality
        if height_ft == 27:
            # Create a new PDF with the same dimensions as the base PDF
            new_pdf = fitz.open()
            new_page = new_pdf.new_page(width=pdf_width, height=pdf_height)

            # Copy the base content to the new page with maximum quality
            new_page.show_pdf_page(
                fitz.Rect(0, 0, pdf_width, pdf_height),
                base_pdf,
                0,
                keep_proportion=True
            )

            # Open the footer PDF
            footer_pdf = fitz.open(footer_pdf_path)
            footer_page = footer_pdf[0]
            footer_rect = footer_page.rect

            # Calculate appropriate footer scaling to maintain aspect ratio
            footer_width = pdf_width
            footer_height = footer_rect.height * (footer_width / footer_rect.width)

            # Insert the footer at the bottom with maximum quality settings
            y1 = pdf_height  # Bottom of the page
            y0 = y1 - footer_height  # Top of the footer

            footer_rect = fitz.Rect(0, y0, footer_width, y1)
            new_page.show_pdf_page(
                footer_rect,
                footer_pdf,
                0,
                keep_proportion=True
            )
            
            # Add text information using PyMuPDF
            text_color = (0, 0, 0)  # Black text
            
            # Position text with consistent positions (same as 13ft panels)
            design_material_x = pdf_width - 430  # For design and material
            height_x = pdf_width - 155  # For height
            
            # Y positions - adjusted to match 13ft panels
            design_y = y0 + 49
            material_y = y0 + 65
            height_y = y0 + 55
            
            # Use PyMuPDF to add text with specified font
            text_font = FONT_NAME  # Use Acumin Pro or fallback
            
            # Add text to the document without labels
            new_page.insert_text((design_material_x, design_y), f"{design_name}", 
                                fontname=text_font, fontsize=FONT_SIZE, color=text_color)
            new_page.insert_text((design_material_x, material_y), f"{material_name}", 
                                fontname=text_font, fontsize=FONT_SIZE, color=text_color)
            new_page.insert_text((height_x, height_y), f"{panel_height_text}", 
                                fontname=text_font, fontsize=FONT_SIZE, color=text_color)

            footer_pdf.close()

            # Save with maximum quality settings
            new_pdf.save(
                final_pdf_path,
                garbage=4,
                deflate=True,
                clean=True,
                linear=True
            )

            new_pdf.close()
            base_pdf.close()
        else:
            # Standard approach for smaller panels
            footer_pdf = fitz.open(footer_pdf_path)
            footer_page = footer_pdf[0]
            footer_rect = footer_page.rect

            # Calculate appropriate footer scaling to maintain aspect ratio
            footer_width = pdf_width
            footer_height = footer_rect.height * (footer_width / footer_rect.width)

            # Place footer at the very bottom of the page
            y1 = pdf_height  # Bottom of the page
            y0 = y1 - footer_height  # Top of the footer

            # Insert the footer as a vector object
            base_page.show_pdf_page(
                fitz.Rect(0, y0, footer_width, y1),
                footer_pdf,
                0,
                clip=None,
                keep_proportion=True
            )
            
            # Add text information using PyMuPDF
            text_color = (0, 0, 0)  # Black text
            
            # Position text - keeping the original positions for 13ft panels
            design_material_x = pdf_width - 430  # For design and material
            height_x = pdf_width - 155  # For height
            
            # Y positions - adjust based on footer layout
            design_y = y0 + 49
            material_y = y0 + 65
            height_y = y0 + 55
            
            # Use PyMuPDF to add text with specified font
            text_font = FONT_NAME  # Use Acumin Pro or fallback
            
            # Add text to the document without labels
            base_page.insert_text((design_material_x, design_y), f"{design_name}", 
                                fontname=text_font, fontsize=FONT_SIZE, color=text_color)
            base_page.insert_text((design_material_x, material_y), f"{material_name}", 
                                fontname=text_font, fontsize=FONT_SIZE, color=text_color)
            base_page.insert_text((height_x, height_y), f"{panel_height_text}", 
                                fontname=text_font, fontsize=FONT_SIZE, color=text_color)

            footer_pdf.close()

            # Save with maximum quality settings
            base_pdf.save(
                final_pdf_path,
                garbage=4,
                deflate=True,
                clean=True,
                linear=True
            )

            base_pdf.close()

        print(f"[✅] Final PDF with footer and text information saved: {final_pdf_path}")

        # Clean up temporary file
        os.remove(base_pdf_path)

    except Exception as e:
        print(f"Error overlaying footer: {e}")

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

        # Define default footer optimization parameters
        default_footer_upscale = 4
        default_footer_sharpness = 1.2

        # Generate only the single blade versions (6 panels)
        for substrate in substrates:
            for height in heights:
                for bleed_mm in bleed_mm_values:
                    # Create only single blade version
                    create_pdf(image_path, height_ft=height, substrate=substrate,
                                      design_name=design_name, bleed_mm=bleed_mm,
                                      footer_upscale=default_footer_upscale,
                                      footer_sharpness=default_footer_sharpness)