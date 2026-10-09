import sys
import os
import re
import subprocess
import urllib.parse
from datetime import datetime

log_file = r"C:\Users\DELL\.gemini\antigravity\brain\f0678079-bf49-4bdd-be79-d5e6298486ad\scratch\open_log.txt"

def log(msg):
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass

def open_pdf_at_page(uri_input):
    log(f"Received URI: {uri_input}")
    base_folder = r"d:\Cong ty\DuToanChuyenNghiep\TT 38\TT 38"
    
    pdf_map = {
        'pl1': '38.2026.tt-bxd-pl1dmdtkhaosat_signed1.pdf',
        'pl2': '38.2026.tt-bxd-pl2dmdtxaydung_signed1.pdf',
        'pl3': '38.2026.tt-bxd-pl3dmdtlapdathtkt_signed1.pdf',
        'pl4': '38.2026.tt-bxd-pl4dmdtldthietbi_signed1.pdf',
        'pl5': '38.2026.tt-bxd-pl5dmdtthinghiem_signed1.pdf',
        'pl6': '38.2026.tt-bxd-pl6dmdtsuachua_signed1.pdf',
        'pl7': '38.2026.tt-bxd-pl7dmdtvatlieu_signed1.pdf',
        'pl8': '38.2026.tt-bxd-pl8dmqldavatvxd_signed1.pdf'
    }

    raw = uri_input.strip()
    if raw.lower().startswith('tt38://'):
        raw = raw[7:]
    elif raw.lower().startswith('tt38:'):
        raw = raw[5:]
        
    raw = urllib.parse.unquote(raw).rstrip('/')
    
    pdf_key = ""
    page = "1"
    
    if '?' in raw:
        query = raw.split('?', 1)[1]
        params = urllib.parse.parse_qs(query)
        pdf_key = params.get('file', params.get('pdf', ['']))[0]
        page = params.get('page', ['1'])[0]
    else:
        parts = raw.split('/')
        if len(parts) >= 2:
            pdf_key = parts[0]
            page = parts[1]
        elif len(parts) == 1:
            pdf_key = parts[0]

    pdf_key_lower = pdf_key.lower().strip()
    target_file = pdf_map.get(pdf_key_lower, None)
    
    if not target_file:
        for k, v in pdf_map.items():
            if k in pdf_key_lower or v.lower() in pdf_key_lower:
                target_file = v
                break
                
    if not target_file:
        target_file = pdf_map['pl2']

    full_pdf_path = os.path.join(base_folder, target_file)
    log(f"Target PDF: {full_pdf_path}, Page: {page}")
    
    # Ưu tiên mở bằng Microsoft Edge hoặc Chrome vì hỗ trợ chuyển thẳng đến trang PDF #page=... cực kỳ mượt mà
    # 1. Edge
    edge_exe = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if os.path.exists(edge_exe):
        try:
            file_url = f"file:///{full_pdf_path.replace(os.sep, '/')}#page={page}"
            log(f"Launching Edge with {file_url}")
            subprocess.Popen([edge_exe, file_url])
            return True
        except Exception as e:
            log(f"Edge error: {e}")

    # 2. Chrome
    chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    if os.path.exists(chrome_exe):
        try:
            file_url = f"file:///{full_pdf_path.replace(os.sep, '/')}#page={page}"
            log(f"Launching Chrome with {file_url}")
            subprocess.Popen([chrome_exe, file_url])
            return True
        except Exception as e:
            log(f"Chrome error: {e}")

    # 3. Foxit Reader
    foxit_exe = r"C:\PROGRAM FILES (X86)\FOXIT SOFTWARE\FOXIT READER\FOXITREADER.EXE"
    if os.path.exists(foxit_exe):
        try:
            log(f"Launching Foxit Reader /A page={page}")
            subprocess.Popen([foxit_exe, '/A', f'page={page}', full_pdf_path])
            return True
        except Exception as e:
            log(f"Foxit error: {e}")

    os.startfile(full_pdf_path)
    return True

if __name__ == "__main__":
    if len(sys.argv) > 1:
        open_pdf_at_page(sys.argv[1])
    else:
        open_pdf_at_page("tt38://pl2/577")
