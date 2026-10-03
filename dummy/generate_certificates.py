import os
import sys
import uuid
import re
import shutil
import pandas as pd
import qrcode
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
import win32com.client
import pythoncom

from config import (
    DUMMY_DATA_CSV,
    PPTX_TEMPLATES,
    QR_CONFIGS,
    PPTX_OUTPUT_DIR,
    PDF_OUTPUT_DIR,
    QR_LEFT,
    QR_TOP,
    QR_SIZE
)

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        pass

def get_all_text_frames(shapes):
    """Recursively retrieves all text frames from slide shapes, groups, and tables."""
    text_frames = []
    for shape in shapes:
        if shape.has_text_frame:
            text_frames.append(shape.text_frame)
        if shape.shape_type == 6:  # Group shape
            text_frames.extend(get_all_text_frames(shape.shapes))
        elif shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text_frame:
                        text_frames.append(cell.text_frame)
    return text_frames

def replace_text_in_slide(slide, tag, value):
    """Replaces text tag in slide text while preserving style."""
    text_frames = get_all_text_frames(slide.shapes)
    replaced_count = 0
    for tf in text_frames:
        is_name_field = False
        for para in tf.paragraphs:
            if re.search(re.escape(tag), para.text, re.IGNORECASE):
                if tag.lower() == "<<name>>":
                    is_name_field = True
        
        if is_name_field:
            tf.word_wrap = False

        for para in tf.paragraphs:
            pattern = re.compile(re.escape(tag), re.IGNORECASE)
            if pattern.search(para.text):
                found_in_single_run = False
                for run in para.runs:
                    if pattern.search(run.text):
                        run.text = pattern.sub(str(value), run.text)
                        replaced_count += 1
                        found_in_single_run = True
                        if tag.lower() == "<<name>>":
                            name_len = len(str(value))
                            if name_len > 20:
                                orig_size_emu = run.font.size
                                orig_size_pt = orig_size_emu.pt if orig_size_emu else 50.58
                                new_size_pt = max(28.0, orig_size_pt * (20.0 / name_len))
                                run.font.size = Pt(new_size_pt)
                        break
                
                if not found_in_single_run and len(para.runs) > 0:
                    full_text = para.text
                    new_text = pattern.sub(str(value), full_text)
                    para.runs[0].text = new_text
                    for r in para.runs[1:]:
                        r.text = ""
                    replaced_count += 1
                    
                    if tag.lower() == "<<name>>":
                        run = para.runs[0]
                        name_len = len(str(value))
                        if name_len > 20:
                            orig_size_emu = run.font.size
                            orig_size_pt = orig_size_emu.pt if orig_size_emu else 50.58
                            new_size_pt = max(28.0, orig_size_pt * (20.0 / name_len))
                            run.font.size = Pt(new_size_pt)
    return replaced_count

def insert_qr_code(slide, qr_url, temp_dir, left_pt=QR_LEFT, top_pt=QR_TOP, size_pt=QR_SIZE):
    """Generates transparent QR code and places it on the slide."""
    os.makedirs(temp_dir, exist_ok=True)
    qr_filename = os.path.join(temp_dir, f"qr_{uuid.uuid4().hex}.png")
    
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=4,
    )
    qr.add_data(qr_url)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    pix = img.load()
    
    # Make white background transparent
    for y in range(img.size[1]):
        for x in range(img.size[0]):
            r, g, b, a = pix[x, y]
            if r > 200 and g > 200 and b > 200:
                pix[x, y] = (255, 255, 255, 0)
                
    img.save(qr_filename)
    
    slide.shapes.add_picture(
        qr_filename,
        Inches(left_pt / 72.0),
        Inches(top_pt / 72.0),
        width=Inches(size_pt / 72.0),
        height=Inches(size_pt / 72.0)
    )
    
    try:
        os.remove(qr_filename)
    except Exception:
        pass

def convert_pptx_to_pdf_win32(pptx_paths, output_folder):
    """Converts a batch of PPTX files to PDF using PowerPoint via win32com."""
    print(f"🔄 Starting PowerPoint application to convert {len(pptx_paths)} presentations...")
    pythoncom.CoInitialize()
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    
    converted_count = 0
    try:
        for idx, pptx_path in enumerate(pptx_paths, 1):
            abs_pptx = os.path.abspath(pptx_path)
            base_name = os.path.splitext(os.path.basename(pptx_path))[0]
            abs_pdf = os.path.abspath(os.path.join(output_folder, f"{base_name}.pdf"))
            
            print(f"   [{idx}/{len(pptx_paths)}] Converting: {base_name}.pptx -> {base_name}.pdf")
            try:
                # Open read-only (1), without window/minimized (0)
                deck = powerpoint.Presentations.Open(abs_pptx, WithWindow=False)
                deck.SaveAs(abs_pdf, 32)  # 32 = ppSaveAsPDF
                deck.Close()
                converted_count += 1
            except Exception as e:
                print(f"   ❌ Failed to convert {base_name}: {e}")
    finally:
        try:
            powerpoint.Quit()
            print("✅ PowerPoint process closed successfully.")
        except Exception as e:
            print(f"⚠️ Error closing PowerPoint: {e}")
        pythoncom.CoUninitialize()
        
    return converted_count

def generate_all_certificates():
    print("🚀 Starting Saavan '26 Certificate Generation...")
    
    if not os.path.exists(DUMMY_DATA_CSV):
        raise FileNotFoundError(f"Data file '{DUMMY_DATA_CSV}' not found.")

    os.makedirs(PPTX_OUTPUT_DIR, exist_ok=True)
    os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)
    temp_qr_dir = os.path.join(PPTX_OUTPUT_DIR, "temp_qr")

    df = pd.read_csv(DUMMY_DATA_CSV)
    print(f"📋 Loaded {len(df)} dummy records across 3 types from {DUMMY_DATA_CSV}")

    generated_pptx_paths = []

    for idx, row in df.iterrows():
        cert_id = row['cert_id']
        name = row['name']
        position = str(row['position']) if pd.notna(row['position']) else ""
        event_name = row['event_name']
        qr_url = row['qr_code']
        cert_type = str(row['type']).strip().lower()

        template_file = PPTX_TEMPLATES.get(cert_type, PPTX_TEMPLATES["winner"])
        if not os.path.exists(template_file):
            raise FileNotFoundError(f"Template file '{template_file}' for type '{cert_type}' not found.")

        qr_cfg = QR_CONFIGS.get(cert_type, {"left": QR_LEFT, "top": QR_TOP, "size": QR_SIZE})

        print(f"   [{idx+1}/{len(df)}] Generating PPTX for {name} ({cert_type.upper()} - {event_name} - {position})")
        print(f"       Using template: {os.path.basename(template_file)}, QR top: {qr_cfg['top']} pt")

        prs = Presentation(template_file)
        slide = prs.slides[0]

        replace_text_in_slide(slide, "<<name>>", name)
        if cert_type == "winner":
            replace_text_in_slide(slide, "<<position>>", position)
        replace_text_in_slide(slide, "<<event_name>>", event_name)

        insert_qr_code(
            slide,
            qr_url,
            temp_qr_dir,
            left_pt=qr_cfg["left"],
            top_pt=qr_cfg["top"],
            size_pt=qr_cfg["size"]
        )

        pptx_out = os.path.join(PPTX_OUTPUT_DIR, f"{cert_id}.pptx")
        prs.save(pptx_out)
        generated_pptx_paths.append(pptx_out)

    if os.path.exists(temp_qr_dir):
        shutil.rmtree(temp_qr_dir, ignore_errors=True)

    print(f"✨ Successfully generated {len(generated_pptx_paths)} PPTX files.")
    
    # Convert PPTX to PDF
    pdf_count = convert_pptx_to_pdf_win32(generated_pptx_paths, PDF_OUTPUT_DIR)
    print(f"🎉 Certificate generation complete! {pdf_count} PDFs generated in '{PDF_OUTPUT_DIR}'.")
    return generated_pptx_paths

if __name__ == "__main__":
    generate_all_certificates()
