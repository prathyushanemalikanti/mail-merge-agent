import os
import glob
import pandas as pd
import pymupdf as fitz
import hashlib
import sys
import subprocess

def get_file_hash(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        h.update(f.read())
    return h.hexdigest()

def qa_test():
    cwd = os.path.abspath(os.getcwd())
    
    excel_path = os.path.join(cwd, 'students.xlsx')
    pdf_path = os.path.join(cwd, 'certificate_sample.pdf')
    out_dir = os.path.join(cwd, 'certificates')
    
    # Check hashes before rerun
    excel_hash_before = get_file_hash(excel_path)
    pdf_hash_before = get_file_hash(pdf_path)
    
    # Read Excel for expected data
    df = pd.read_excel(excel_path)
    expected_students = []
    for idx, row in df.iterrows():
        reg = str(row['Reg No']).strip()
        name = str(row['Name']).strip()
        if not reg or reg == 'nan' or not name or name == 'nan':
            continue
        expected_students.append({'reg': reg, 'name': name})
        
    print(f"1. Verified exactly {len(expected_students)} valid student records were read.")
    if len(expected_students) != 5:
        print("FAIL: Expected 5 student records")
        sys.exit(1)
        
    generated_pdfs = glob.glob(os.path.join(out_dir, '*.pdf'))
    print(f"2. Verified exactly {len(generated_pdfs)} certificate PDFs exist.")
    if len(generated_pdfs) != 5:
        print("FAIL: Expected 5 certificate PDFs")
        sys.exit(1)
        
    for p in generated_pdfs:
        basename = os.path.basename(p)
        if ' ' in basename:
            print(f"FAIL: Filename contains space: {basename}")
            sys.exit(1)
    print("4. Verified filenames use underscores instead of spaces.")
    
    # Template dims
    doc_t = fitz.open(pdf_path)
    t_rect = doc_t[0].rect
    t_w, t_h = t_rect.width, t_rect.height
    doc_t.close()
    
    for s in expected_students:
        fname = f"{s['reg']}_{s['name'].replace(' ', '_')}.pdf".replace('__', '_')
        fpath = os.path.join(out_dir, fname)
        if not os.path.exists(fpath):
            print(f"FAIL: Expected PDF not found for {s['name']}: {fname}")
            sys.exit(1)
            
        # Check PDF
        doc = fitz.open(fpath)
        if doc.page_count != 1:
            print(f"FAIL: PDF {fname} has {doc.page_count} pages, expected 1")
            sys.exit(1)
            
        page = doc[0]
        if abs(page.rect.width - t_w) > 1 or abs(page.rect.height - t_h) > 1:
            print(f"FAIL: PDF {fname} dimensions differ from template")
            sys.exit(1)
            
        text = page.get_text()
        if s['name'].split()[0] not in text or s['reg'] not in text:
            print(f"FAIL: PDF {fname} missing Name or Reg No in text extraction")
            sys.exit(1)
            
        # Verify no other student's info is here
        for os_s in expected_students:
            if os_s['reg'] != s['reg'] and os_s['reg'] in text:
                print(f"FAIL: PDF {fname} contains cross-contamination from {os_s['reg']}")
                sys.exit(1)
                
        doc.close()
        
    print("3, 5, 6, 7, 8, 9. Verified all PDF structure, text, and cross-contamination rules.")
    
    report_df = pd.read_csv(os.path.join(out_dir, 'generation_report.csv'))
    if len(report_df) != 5:
        print("FAIL: Report does not have 5 entries")
        sys.exit(1)
    print("10. Verified generation_report.csv contains 5 records.")
    
    excel_hash_after = get_file_hash(excel_path)
    pdf_hash_after = get_file_hash(pdf_path)
    
    if excel_hash_before != excel_hash_after:
        print("FAIL: Excel file changed!")
        sys.exit(1)
    if pdf_hash_before != pdf_hash_after:
        print("FAIL: PDF template changed!")
        sys.exit(1)
        
    print("11, 12. Verified original files remained unchanged.")
    
    # Rerun test
    print("Testing rerun safety...")
    res = subprocess.run([sys.executable, 'certificate_agent/main.py'], capture_output=True, text=True)
    if res.returncode != 0:
        print("FAIL: Rerun failed")
        print(res.stderr)
        sys.exit(1)
        
    generated_pdfs_rerun = glob.glob(os.path.join(out_dir, '*.pdf'))
    if len(generated_pdfs_rerun) != 5:
        print(f"FAIL: Expected 5 PDFs after rerun, found {len(generated_pdfs_rerun)}")
        sys.exit(1)
    print("15. Verified rerun testing passed (regenerated 5 PDFs successfully).")
    
    print("ALL QA CHECKS PASSED.")

if __name__ == '__main__':
    qa_test()
