import os
import glob
import re
import csv
import sys
import shutil
import pandas as pd
import pymupdf as fitz # PyMuPDF

def sanitize_filename(name):
    # Replace spaces with underscores
    name = str(name).replace(' ', '_')
    # Remove Windows-invalid characters
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    # Clean up multiple consecutive underscores
    name = re.sub(r'_+', '_', name)
    return name.strip('_')

def normalize_col(col):
    return re.sub(r'\s+', '', str(col).lower())

def find_excel_file(cwd):
    excel_files = glob.glob(os.path.join(cwd, '*.xlsx')) + glob.glob(os.path.join(cwd, '*.xls'))
    valid_files = []
    
    reg_variations = ['regno', 'registrationno', 'registrationnumber']
    name_variations = ['name', 'studentname']
    
    for f in excel_files:
        try:
            df = pd.read_excel(f, nrows=0)
            cols = [normalize_col(c) for c in df.columns]
            has_reg = any(c in reg_variations for c in cols)
            has_name = any(c in name_variations for c in cols)
            if has_reg and has_name:
                valid_files.append(f)
        except Exception:
            continue
            
    if len(valid_files) == 0:
        print("Error: No suitable Excel file found with required columns.")
        sys.exit(1)
    if len(valid_files) > 1:
        print(f"Error: Ambiguous Excel files found: {valid_files}")
        sys.exit(1)
        
    return valid_files[0]

def find_pdf_template(cwd):
    pdf_files = glob.glob(os.path.join(cwd, '*.pdf'))
    valid_files = []
    out_dir_path = os.path.abspath(os.path.join(cwd, 'certificates'))
    
    for f in pdf_files:
        abs_f = os.path.abspath(f)
        if abs_f.startswith(out_dir_path):
            continue
        try:
            doc = fitz.open(f)
            text = ""
            for page in doc:
                text += page.get_text()
            if "{{NAME}}" in text and "{{REG_NO}}" in text:
                valid_files.append(f)
        except Exception:
            continue
            
    if len(valid_files) == 0:
        print("Error: No suitable PDF template found containing {{NAME}} and {{REG_NO}}.")
        sys.exit(1)
    if len(valid_files) > 1:
        print(f"Error: Ambiguous PDF templates found: {valid_files}")
        sys.exit(1)
        
    return valid_files[0]

def get_col_name(df, variations):
    cols = list(df.columns)
    for c in cols:
        if normalize_col(c) in variations:
            return c
    return None

def process_certificates():
    cwd = os.getcwd()
    
    try:
        excel_path = find_excel_file(cwd)
        pdf_path = find_pdf_template(cwd)
    except Exception as e:
        print(e)
        sys.exit(1)
        
    print(f"Detected Excel: {os.path.basename(excel_path)}")
    print(f"Detected PDF Template: {os.path.basename(pdf_path)}")
    
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)
        
    reg_col = get_col_name(df, ['regno', 'registrationno', 'registrationnumber'])
    name_col = get_col_name(df, ['name', 'studentname'])
    
    if not reg_col or not name_col:
        print("Error: Required columns missing in Excel.")
        sys.exit(1)
        
    df = df.dropna(subset=[reg_col, name_col], how='all')
    
    out_dir = os.path.join(cwd, 'certificates')
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)
    else:
        for f in glob.glob(os.path.join(out_dir, '*.pdf')):
            os.remove(f)
        
    report_path = os.path.join(out_dir, 'generation_report.csv')
    
    students = []
    for idx, row in df.iterrows():
        reg = str(row[reg_col]).strip()
        name = str(row[name_col]).strip()
        if not reg or reg == 'nan' or not name or name == 'nan':
            continue
        students.append({'reg': reg, 'name': name})
        
    success_count = 0
    failed_students = []
    
    # Analyze original template to get base font sizes
    doc_template = fitz.open(pdf_path)
    template_page = doc_template[0]
    template_dims = (template_page.rect.width, template_page.rect.height)
    
    # We will approximate font sizes. The original {{NAME}} is likely large.
    # Let's inspect the block containing {{NAME}} to get its font size if possible.
    name_blocks = template_page.get_text("dict")["blocks"]
    base_name_size = 30
    base_reg_size = 14
    for b in name_blocks:
        if "lines" in b:
            for l in b["lines"]:
                for s in l["spans"]:
                    if "{{NAME}}" in s["text"]:
                        base_name_size = s["size"]
                    if "{{REG_NO}}" in s["text"]:
                        base_reg_size = s["size"]
    doc_template.close()
    
    with open(report_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Reg No', 'Name', 'Status', 'Output File'])
        
        for s in students:
            try:
                doc = fitz.open(pdf_path)
                page = doc[0]
                
                name_rects = page.search_for("{{NAME}}")
                reg_rects = page.search_for("{{REG_NO}}")
                
                if not name_rects or not reg_rects:
                    raise Exception("Placeholders not found on page")
                    
                name_rect = name_rects[0]
                reg_rect = reg_rects[0]
                
                page.add_redact_annot(name_rect)
                page.add_redact_annot(reg_rect)
                page.apply_redactions()
                
                fontname = "helv"
                fontsize = base_name_size
                
                max_width = page.rect.width - 100
                while True:
                    text_len = fitz.get_text_length(s['name'], fontname=fontname, fontsize=fontsize)
                    if text_len <= max_width or fontsize <= 10:
                        break
                    fontsize -= 1
                    
                center_x = (name_rect.x0 + name_rect.x1) / 2
                x_pos = center_x - text_len / 2
                if x_pos < 50:
                    x_pos = 50
                    
                # y1 is the bottom of the bounding box. Text insertion baseline.
                # Usually text bounding box y1 is slightly below the baseline.
                y_pos_name = name_rect.y1 - (name_rect.height * 0.15)
                page.insert_text((x_pos, y_pos_name), s['name'], fontsize=fontsize, fontname=fontname, color=(0,0,0))
                
                y_pos_reg = reg_rect.y1 - (reg_rect.height * 0.15)
                page.insert_text((reg_rect.x0, y_pos_reg), s['reg'], fontsize=base_reg_size, fontname=fontname, color=(0,0,0))
                
                fname = sanitize_filename(f"{s['reg']}_{s['name']}.pdf")
                out_path = os.path.join(out_dir, fname)
                doc.save(out_path)
                doc.close()
                
                # Validation
                if not os.path.exists(out_path) or os.path.getsize(out_path) == 0:
                    raise Exception("File empty or not created")
                    
                vdoc = fitz.open(out_path)
                if vdoc.page_count != 1:
                    raise Exception("Incorrect page count")
                vpage = vdoc[0]
                if abs(vpage.rect.width - template_dims[0]) > 1 or abs(vpage.rect.height - template_dims[1]) > 1:
                    raise Exception("Page dimensions changed")
                    
                vtext = vpage.get_text()
                if s['name'].split()[0] not in vtext or s['reg'] not in vtext:
                    raise Exception("Inserted text not found in extracted PDF text")
                    
                vdoc.close()
                
                writer.writerow([s['reg'], s['name'], 'Success', fname])
                success_count += 1
                
            except Exception as e:
                writer.writerow([s['reg'], s['name'], f'Failed: {e}', ''])
                failed_students.append((s, str(e)))
                
    print("\nCertificate generation completed.")
    print(f"Students found:       {len(students)}")
    print(f"Certificates created: {success_count}")
    print(f"Failed:               {len(failed_students)}")
    print(f"\nOutput directory:\n{out_dir}")
    
    if failed_students:
        print("\nFailed students:")
        for s, err in failed_students:
            print(f"- {s['reg']} ({s['name']}): {err}")
            
    if not failed_students:
        print("\nAll certificates successfully generated.")
        
    # Visual check: render images for first, normal, longest
    if success_count > 0:
        print("\nGenerating preview images for visual QA...")
        students.sort(key=lambda x: len(x['name']))
        samples = [students[0], students[len(students)//2], students[-1]]
        for i, s in enumerate(samples):
            fname = sanitize_filename(f"{s['reg']}_{s['name']}.pdf")
            out_path = os.path.join(out_dir, fname)
            if os.path.exists(out_path):
                vdoc = fitz.open(out_path)
                pix = vdoc[0].get_pixmap(dpi=150)
                pix.save(os.path.join(out_dir, f"preview_{i}.png"))
                vdoc.close()
        print("Previews saved in certificates/ folder.")

if __name__ == "__main__":
    process_certificates()
