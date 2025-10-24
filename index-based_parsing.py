import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageOps, ImageFile, ImageEnhance
import json
import os

# Increase PIL limits for large images
ImageFile.LOAD_TRUNCATED_IMAGES = True
Image.MAX_IMAGE_PIXELS = 200000000

# Data structure template
SLICE_TEMPLATE = {
    "to_email_address": "",
    "from_email_address": "",
    "subject_line": "",
    "preheader": "",
    "Header_hero_section": "",
    "Headline": "",
    "body_copy": "",
    "sisi": "",
    "primary_cta": "",
    "secondary_cta": "",
    "safety_info": {},
    "pricing_disclosure": "",
    "meet_invite": "",
    "closing": "",
    "stemline": {
        "name": "",
        "phone": "",
        "email": ""
    },
    "abbreviation": "",
    "references": [],
    "company_info": "",
    "mlr_code": "",
    "footer": ""
}


# Index-based parsing utility functions
def find_marker_position(text, marker, start_pos=0, case_sensitive=False):
    """Find position of a marker in text, starting from a specific position."""
    if case_sensitive:
        pos = text.find(marker, start_pos)
    else:
        pos = text.lower().find(marker.lower(), start_pos)
    return pos


def extract_by_indices(text, start_idx=0, end_idx=None):
    """Extract text between specific indices, preserving all content"""
    if end_idx is None:
        end_idx = len(text)
    return text[start_idx:end_idx].strip()


def split_sections_on_unsubscribe_indexed(text):
    """Split sections using index positions instead of regex"""
    unsubscribe_positions = []
    start = 0

    # Find all unsubscribe positions
    while True:
        pos = text.lower().find('unsubscribe', start)
        if pos == -1:
            break
        unsubscribe_positions.append(pos + len('unsubscribe'))
        start = pos + 1

    # Split based on positions
    if not unsubscribe_positions:
        return f"Section 1\n{text}"
    elif len(unsubscribe_positions) == 1:
        section1 = text[:unsubscribe_positions[0]]
        section2 = text[unsubscribe_positions[0]:]
        return f"Section 1\n{section1}\n\nSection 2\n{section2}"
    else:
        section1 = text[:unsubscribe_positions[0]]
        section2 = text[unsubscribe_positions[0]:unsubscribe_positions[1]]
        section3 = text[unsubscribe_positions[1]:]
        return f"Section 1\n{section1}\n\nSection 2\n{section2}\n\nSection 3\n{section3}"


def split_into_sections_indexed(text: str):
    """Split text into sections using index-based approach"""
    section_markers = ['Section 1', 'Section 2', 'Section 3']
    boundaries = []

    for marker in section_markers:
        pos = find_marker_position(text, marker)
        if pos != -1:
            line_end = text.find('\n', pos)
            if line_end == -1:
                line_end = pos + len(marker)
            else:
                line_end += 1

            boundaries.append({
                'marker': marker,
                'start': pos,
                'end': line_end
            })

    boundaries.sort(key=lambda x: x['start'])

    if len(boundaries) == 0:
        return (text.strip(), "", "")
    elif len(boundaries) == 1:
        section1 = extract_by_indices(text, boundaries[0]['end'])
        return (section1, "", "")
    elif len(boundaries) == 2:
        section1 = extract_by_indices(text, boundaries[0]['end'], boundaries[1]['start'])
        section2 = extract_by_indices(text, boundaries[1]['end'])
        return (section1, section2, "")
    else:
        section1 = extract_by_indices(text, boundaries[0]['end'], boundaries[1]['start'])
        section2 = extract_by_indices(text, boundaries[1]['end'], boundaries[2]['start'])
        section3 = extract_by_indices(text, boundaries[2]['end'])
        return (section1, section2, section3)


# OCR text extraction function
def extract_text_from_pdf_page(pdf_path, page_number=1):
    """Extract text from PDF page using OCR"""
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    try:
        # Convert PDF to images
        pages = convert_from_path(
            pdf_path,
            first_page=page_number,
            last_page=page_number,
            dpi=400
        )

        if not pages:
            raise ValueError(f"No page found at number {page_number} in the PDF.")

        image = pages[0]
        gray_image = ImageOps.grayscale(image)

        # Enhance image quality
        enhancer = ImageEnhance.Contrast(gray_image)
        contrast_image = enhancer.enhance(2.0)
        enhancer = ImageEnhance.Sharpness(contrast_image)
        sharp_image = enhancer.enhance(2.0)

        # OCR configuration
        ocr_config = '--psm 3 --oem 3'
        text = pytesseract.image_to_string(sharp_image, lang='eng', config=ocr_config)
        return text

    except Exception as e:
        print(f"Error extracting text from PDF: {str(e)}")
        raise


# Field extraction using index-based parsing - FINAL CORRECTED VERSION
def fill_slice_fields_indexed(text):
    """Fill slice fields using index-based parsing - only extract content that exists in PDF"""
    s1 = dict(SLICE_TEMPLATE)

    def extract_field_indexed(text, field_marker):
        pos = find_marker_position(text, field_marker)
        if pos == -1:
            return ""

        start_idx = pos + len(field_marker)
        end_idx = text.find('\n', start_idx)
        if end_idx == -1:
            end_idx = len(text)

        return extract_by_indices(text, start_idx, end_idx)

    # Extract basic email fields
    s1["to_email_address"] = extract_field_indexed(text, "To:")
    s1["from_email_address"] = extract_field_indexed(text, "From:")
    s1["subject_line"] = extract_field_indexed(text, "Subject Line:")
    s1["preheader"] = extract_field_indexed(text, "Preheader:")

    # Extract header section
    preheader_pos = find_marker_position(text, "Preheader:")
    dear_pos = find_marker_position(text, "Dear Healthcare Professional")

    if preheader_pos != -1 and dear_pos != -1:
        preheader_end = text.find('\n', preheader_pos + len("Preheader:"))
        if preheader_end != -1:
            s1["Header_hero_section"] = extract_by_indices(text, preheader_end + 1, dear_pos)

    # Extract headline
    orserdu_pos = find_marker_position(text, "ORSERDU:")
    if orserdu_pos != -1:
        line_end = text.find('\n', orserdu_pos)
        if line_end != -1:
            s1["Headline"] = extract_by_indices(text, orserdu_pos, line_end)
    else:
        elacestrant_pos = find_marker_position(text, "elacestrant")
        if elacestrant_pos != -1:
            line_end = text.find('\n', elacestrant_pos)
            if line_end != -1:
                s1["Headline"] = extract_by_indices(text, elacestrant_pos, line_end)

    # Extract body copy
    dear_pos = find_marker_position(text, "Dear Healthcare Professional,")
    safety_pos = find_marker_position(text, "SELECT IMPORTANT SAFETY INFORMATION")

    if dear_pos != -1 and safety_pos != -1:
        s1["body_copy"] = extract_by_indices(text, dear_pos, safety_pos)

    # --- SISI and Safety Info Extraction (Corrected Logic) ---
    sisi_start_marker = "SELECT IMPORTANT SAFETY INFORMATION"
    sisi_end_marker = "Please see additional Important Safety Information below"

    # Corrected markers for the main ISI section
    isi_start_marker = "IMPORTANT SAFETY INFORMATION"
    isi_end_marker = "www.fda.gov/medwatch"

    # Initialize fields
    s1["sisi"] = ""
    s1["safety_info"] = {"section": ""}

    # 1. Extract SISI
    sisi_start_pos = find_marker_position(text, sisi_start_marker)
    sisi_end_pos_marker = find_marker_position(text, sisi_end_marker, start_pos=sisi_start_pos)

    sisi_end_pos = -1
    if sisi_start_pos != -1 and sisi_end_pos_marker != -1:
        sisi_end_pos = sisi_end_pos_marker + len(sisi_end_marker)
        s1["sisi"] = extract_by_indices(text, sisi_start_pos, sisi_end_pos)

    # 2. Extract ISI (Safety Info), starting the search after SISI to find the correct section
    search_start_for_isi = sisi_end_pos if sisi_end_pos != -1 else (sisi_start_pos + 1 if sisi_start_pos != -1 else 0)

    # Find the start of the dedicated ISI section
    isi_start_pos = find_marker_position(text, isi_start_marker, start_pos=search_start_for_isi)

    if isi_start_pos != -1:
        # Find the end of the ISI section
        isi_end_pos_marker = find_marker_position(text, isi_end_marker, start_pos=isi_start_pos)

        if isi_end_pos_marker != -1:
            isi_end_pos = isi_end_pos_marker + len(isi_end_marker)
            # Extract the final, clean ISI text
            isi_text = extract_by_indices(text, isi_start_pos, isi_end_pos)
            s1["safety_info"] = {"section": isi_text}

    # Extract CTAs (Call to Actions)
    learn_more_pos = find_marker_position(text, "Learn more about ORSERDU >")
    if learn_more_pos != -1:
        line_end = text.find('\n', learn_more_pos)
        if line_end == -1:
            line_end = len(text)
        s1["primary_cta"] = extract_by_indices(text, learn_more_pos, line_end)

    click_here_pos = find_marker_position(text, "Click here to request more information")
    if click_here_pos != -1:
        line_end = text.find('\n', click_here_pos)
        if line_end == -1:
            line_end = len(text)
        s1["secondary_cta"] = extract_by_indices(text, click_here_pos, line_end)

    # Extract pricing disclosure
    pricing_pos = find_marker_position(text, "For State pricing disclosures")
    if pricing_pos != -1:
        line_end = text.find('\n', pricing_pos)
        if line_end == -1:
            line_end = len(text)
        s1["pricing_disclosure"] = extract_by_indices(text, pricing_pos, line_end)

    # Extract abbreviations
    abbr_pos = find_marker_position(text, "Abbreviations:")
    ref_pos = find_marker_position(text, "References:")

    if abbr_pos != -1 and ref_pos != -1:
        abbr_content = extract_by_indices(text, abbr_pos + len("Abbreviations:"), ref_pos)
        s1["abbreviation"] = abbr_content.replace('\n', ' ').strip()

    # Extract references
    if ref_pos != -1:
        menarini_pos = find_marker_position(text, "MENARINI")
        if menarini_pos != -1:
            s1["references"] = extract_by_indices(text, ref_pos + len("References:"), menarini_pos)

    # Extract company info
    company_start_pos = find_marker_position(text, "A Menarini Group Company")
    if company_start_pos != -1:
        privacy_pos = find_marker_position(text, "Privacy and Terms of Use")
        if privacy_pos != -1:
            s1["company_info"] = extract_by_indices(text, company_start_pos, privacy_pos)

    # Extract MLR code
    mlr_start = find_marker_position(text, "MAT-US-ELA-")
    if mlr_start != -1:
        mlr_end = mlr_start + len("MAT-US-ELA-")
        while mlr_end < len(text) and text[mlr_end] not in [' ', '\n', '\t']:
            mlr_end += 1
        s1["mlr_code"] = extract_by_indices(text, mlr_start, mlr_end)

    # Extract footer
    privacy_pos = find_marker_position(text, "Privacy and Terms of Use")
    if privacy_pos != -1:
        line_end = text.find('\n', privacy_pos)
        if line_end == -1:
            line_end = len(text)
        s1["footer"] = extract_by_indices(text, privacy_pos, line_end)

    # CORRECTED: Only populate stemline if there's an actual dedicated stemline section
    stemline_section_markers = [
        "Stemline:",
        "About Stemline:",
        "Stemline Information:",
        "Stemline Therapeutics:",
        "Contact Stemline:",
        "Stemline Contact Information:"
    ]

    stemline_section_found = False
    for marker in stemline_section_markers:
        stemline_section_pos = find_marker_position(text, marker)
        if stemline_section_pos != -1:
            stemline_section_found = True
            section_end = text.find('\n\n', stemline_section_pos)
            if section_end == -1:
                next_section_pos = find_marker_position(text, ":", start_pos=stemline_section_pos + len(marker))
                if next_section_pos != -1:
                    section_end = stemline_section_pos + len(marker) + next_section_pos
                else:
                    section_end = len(text)

            stemline_content = extract_by_indices(text, stemline_section_pos, section_end)

            if "Stemline Therapeutics" in stemline_content:
                name_start = stemline_content.find("Stemline Therapeutics")
                name_end = stemline_content.find('\n', name_start)
                if name_end == -1: name_end = len(stemline_content)
                s1["stemline"]["name"] = extract_by_indices(stemline_content, name_start, name_end)

            if "1-877-332-7961" in stemline_content:
                s1["stemline"]["phone"] = "1-877-332-7961"

            email_patterns = ["@stemline.com", "@menarini.com"]
            for pattern in email_patterns:
                email_pos = stemline_content.find(pattern)
                if email_pos != -1:
                    start_pos = email_pos
                    while start_pos > 0 and stemline_content[start_pos - 1] not in [' ', '\n', '\t', ':']:
                        start_pos -= 1
                    s1["stemline"]["email"] = extract_by_indices(stemline_content, start_pos, email_pos + len(pattern))
                    break
            break

    if not stemline_section_found:
        s1["stemline"]["name"] = ""
        s1["stemline"]["phone"] = ""
        s1["stemline"]["email"] = ""

    return s1


# Main processing functions
def text_to_json_indexed(text: str):
    """Convert text to JSON using index-based parsing"""
    section1, section2, section3 = split_into_sections_indexed(text)

    slice1 = fill_slice_fields_indexed(section1)
    slice2 = fill_slice_fields_indexed(section2)
    slice3 = fill_slice_fields_indexed(section3)

    result = {
        "page2": {
            "slice1": slice1,
            "slice2": slice2,
            "slice3": slice3
        }
    }

    return result

def process_pdf_with_index_parsing(pdf_path, page_number=2, output_file=None):
    """
    Main function to process PDF with index-based parsing
    """
    try:
        print(f"Processing PDF: {pdf_path}")
        print(f"Page number: {page_number}")

        # Extract text from PDF
        extracted_text = extract_text_from_pdf_page(pdf_path, page_number)

        # Process with sections
        sectioned_text = split_sections_on_unsubscribe_indexed(extracted_text)
        print("Section splitting complete")

        # Convert to structured JSON
        result = text_to_json_indexed(sectioned_text)
        print("JSON conversion complete")

        # Save to file if specified
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=2, ensure_ascii=False)
            print(f"Results saved to: {output_file}")

        return result, sectioned_text, extracted_text

    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        return None, None, None


# Example usage and main execution
if __name__ == "__main__":
    # CONFIGURE YOUR PDF PATH HERE
    PDF_FILE_PATH = r"C:\Users\Lenovo\Downloads\MAT-US-ELA-00626-v2_SFMC_email_ORSERDU vs. Fulvestrant_ver_4.01 - Copy (2).pdf" # Change this to your PDF file path
    PAGE_NUMBER = 2  # Change page number if needed
    OUTPUT_JSON_FILE = "ocr_results.json"  # Output file name

    # Process the PDF
    result, sectioned_text, raw_text = process_pdf_with_index_parsing(
        pdf_path=PDF_FILE_PATH,
        page_number=PAGE_NUMBER,
        output_file=OUTPUT_JSON_FILE
    )

    if result:
        print("\nProcessing completed successfully!")

        # Save sectioned text for review
        with open("sectioned_text.txt", "w", encoding="utf-8") as f:
            f.write(sectioned_text)
        print("\nSectioned text saved to: sectioned_text.txt")

    else:
        print("\nProcessing failed. Please check the error messages above.")