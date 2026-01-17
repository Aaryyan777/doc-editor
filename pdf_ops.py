import fitz
import os
from datetime import datetime

from pdf2docx import Converter
from docx2pdf import convert as docx_to_pdf_conv

def format_date(date_str):
    if not date_str:
        return "N/A"
    try:
        clean_date = date_str.replace("D:", "").split('+')[0].split('-')[0].replace("'", "")
        return datetime.strptime(clean_date[:14], "%Y%m%d%H%M%S").strftime("%Y-%m-%d %H:%M:%S")
    except:
        return date_str

def get_pdf_info(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    info = {
        "pages": doc.page_count,
        "encrypted": doc.is_encrypted,
        "format": doc.metadata.get("format", "PDF"),
        "title": doc.metadata.get("title", "N/A"),
        "author": doc.metadata.get("author", "N/A"),
        "creation_date": format_date(doc.metadata.get("creationDate", "")),
        "mod_date": format_date(doc.metadata.get("modDate", ""))
    }
    doc.close()
    return info

def convert_pdf_to_word(file_path, output_path=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
        
    if output_path is None:
        output_path = os.path.splitext(file_path)[0] + ".docx"
        
    cv = Converter(file_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()
    
    return output_path

def convert_word_to_pdf(file_path, output_path=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    # docx2pdf.convert(input, output)
    docx_to_pdf_conv(file_path, output_path)
    
    if output_path is None:
        output_path = os.path.splitext(file_path)[0] + ".pdf"
    
    return output_path

def rotate_pages(file_path, rotation, output_path=None):
    """Rotation must be 0, 90, 180, 270."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    for page in doc:
        page.set_rotation(rotation)
    
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_rotated{ext}"
    
    doc.save(output_path)
    doc.close()
    return output_path

def delete_pages(file_path, pages_to_delete, output_path=None):
    """pages_to_delete: list of integers (0-based index)"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    # Delete pages in reverse order to avoid index shifting
    for p in sorted(pages_to_delete, reverse=True):
        if 0 <= p < doc.page_count:
            doc.delete_page(p)
            
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_deleted{ext}"
    
    doc.save(output_path)
    doc.close()
    return output_path

def extract_page_range(file_path, start_page, end_page, output_path=None):
    """1-based start and end page numbers."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    # Create new doc
    new_doc = fitz.open()
    
    # Python slices are 0-based and exclusive at the end
    # User input 1-5 means index 0 to 4
    new_doc.insert_pdf(doc, from_page=start_page-1, to_page=end_page-1)
    
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_pages_{start_page}-{end_page}{ext}"
        
    new_doc.save(output_path)
    new_doc.close()
    doc.close()
    return output_path

def encrypt_pdf(file_path, password, output_path=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_encrypted{ext}"
    
    # Save with encryption (AES 256)
    doc.save(output_path, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=password, user_pw=password)
    doc.close()
    return output_path

def decrypt_pdf(file_path, password, output_path=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    if doc.is_encrypted:
        if not doc.authenticate(password):
            raise ValueError("Incorrect password.")
            
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_decrypted{ext}"
        
    doc.save(output_path) # Saving creates an unencrypted copy
    doc.close()
    return output_path

def extract_text_from_pdf(file_path, output_path=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    text = ""
    for page in doc:
        text += page.get_text() + "\n\n"
    
    if output_path is None:
        output_path = os.path.splitext(file_path)[0] + ".txt"
        
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)
    
    doc.close()
    return output_path

def merge_pdfs(file_list, output_path="merged.pdf"):
    merged_doc = fitz.open()
    for file in file_list:
        if os.path.exists(file):
            doc = fitz.open(file)
            merged_doc.insert_pdf(doc)
            doc.close()
    
    merged_doc.save(output_path)
    merged_doc.close()
    return output_path

def redact_pdf(file_path, text_to_redact, output_path=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    count = 0
    for page in doc:
        areas = page.search_for(text_to_redact)
        for area in areas:
            page.add_redact_annot(area, fill=(0, 0, 0))
            count += 1
        page.apply_redactions()
    
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_redacted{ext}"
        
    doc.save(output_path)
    doc.close()
    return count, output_path

def edit_pdf_text(file_path, old_text, new_text, output_path=None):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' not found.")
    
    doc = fitz.open(file_path)
    count = 0
    for page in doc:
        areas = page.search_for(old_text)
        for rect in areas:
            # Redact (white out)
            page.add_redact_annot(rect, fill=(1, 1, 1)) 
            page.apply_redactions()
            # Insert new text
            page.insert_textbox(rect, new_text, color=(0, 0, 0), fontsize=11)
            count += 1
            
    if output_path is None:
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_edited{ext}"
        
    doc.save(output_path)
    doc.close()
    return count, output_path
