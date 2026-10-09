import os
import sys
import time
import re
import socket
import subprocess
import webbrowser
import atexit

UV_PATH = r"C:\Users\DELL\.local\bin\uv.exe"
CLOUDFLARED_PATH = r"C:\Program Files (x86)\cloudflared\cloudflared.exe"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_SCRIPT = os.path.join(BASE_DIR, "app_web.py")
LINK_FILE = os.path.join(BASE_DIR, "LINK_TRUY_CAP_ONLINE.txt")

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def main():
    print("=" * 80)
    print("   HE THONG THAM TRA DU TOAN - DANG KHOI DONG DICH VU TRUY CAP ONLINE...")
    print("=" * 80)
    print()

    # 1. Khoi dong Web Server neu chua chay
    web_proc = None
    if not is_port_in_use(5000):
        print("[1/3] Dang khoi dong Web Server tai cong 5000...")
        web_cmd = [
            UV_PATH, "run",
            "--with", "flask",
            "--with", "python-calamine",
            "--with", "openpyxl",
            "--with", "python-docx",
            "python", APP_SCRIPT
        ]
        web_proc = subprocess.Popen(
            web_cmd,
            cwd=BASE_DIR,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        # Cho web server khoi dong
        for _ in range(15):
            if is_port_in_use(5000):
                break
            time.sleep(1)
        print("  -> Web Server da san sang tren cong 5000!")
    else:
        print("[1/3] Web Server da dang chay tren cong 5000.")

    # 2. Khoi dong Cloudflare Tunnel
    if not os.path.exists(CLOUDFLARED_PATH):
        print(f"[LOI] Khong tim thay cloudflared.exe tai: {CLOUDFLARED_PATH}")
        sys.exit(1)

    print("[2/3] Dang ket noi Cloudflare de lay duong link Online Public toan cau...")
    cf_proc = subprocess.Popen(
        [CLOUDFLARED_PATH, "tunnel", "--url", "http://127.0.0.1:5000"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1
    )

    def cleanup():
        if cf_proc and cf_proc.poll() is None:
            cf_proc.terminate()
        if web_proc and web_proc.poll() is None:
            web_proc.terminate()

    atexit.register(cleanup)

    tunnel_url = None
    for line in cf_proc.stdout:
        match = re.search(r'https://[a-zA-Z0-9-]+\.trycloudflare\.com', line)
        if match:
            tunnel_url = match.group(0)
            break

    lan_ip = get_local_ip()

    if tunnel_url:
        # Ghi vao file de nguoi dung tien copy
        content_txt = f"""================================================================================
   THONG TIN DUONG LINK TRUY CAP HE THONG THAM TRA DU TOAN
================================================================================

1. DUONG LINK ONLINE CHO CAC MAY KHAC / DIEN THOAI (TRUY CAP TU XA TOAN CAU):
   👉 {tunnel_url}
   (Chi can gui link nay qua Zalo/Email la may khac vao duoc ngay)

2. DUONG LINK NOI BO CHO CAC MAY CUNG WI-FI / MANG LAN CONG TY:
   👉 http://{lan_ip}:5000

3. DUONG LINK MANG NOI BO TREN MAY NAY:
   👉 http://localhost:5000

Luu y: Vui long giu cua so den dang chay de duy tri ket noi online!
================================================================================
"""
        with open(LINK_FILE, "w", encoding="utf-8") as f:
            f.write(content_txt)

        print()
        print("=" * 80)
        print("   DA TAO DUONG LINK TRUY CAP ONLINE THANH CONG!")
        print("=" * 80)
        print()
        print("🌐 1. LINK ONLINE CHO MAY KHAC HOAC DIEN THOAI (TRUY CAP TU BAT KY DAU):")
        print(f"      👉  {tunnel_url}")
        print()
        print("🏠 2. LINK CHO CAC MAY CHUNG MANG WI-FI CONG TY:")
        print(f"      👉  http://{lan_ip}:5000")
        print()
        print("💻 3. LINK TREN CHINH MAY NAY:")
        print("      👉  http://localhost:5000")
        print()
        print("=" * 80)
        print(f"[*] Da luu link vao file text: LINK_TRUY_CAP_ONLINE.txt (trong thu muc nay)")
        print("[*] Dang tu dong mo trinh duyet voi link online...")
        print("[*] LUU Y: GIU CUA SO NAY DANG CHAY. Khong bam tat [X].")
        print("    (Nhan Ctrl+C neu muon dung lai)")
        print("=" * 80)
        print()

        try:
            webbrowser.open(tunnel_url)
        except Exception:
            pass

        # Duy tri tien trinh
        try:
            while True:
                time.sleep(1)
                if cf_proc.poll() is not None:
                    print("\n[THONG BAO] Cloudflare tunnel da dung lai.")
                    break
        except KeyboardInterrupt:
            print("\nDang tat cac dich vu...")
            cleanup()
    else:
        print("[LOI] Khong the lay duoc duong link tu Cloudflare.")
        cleanup()

if __name__ == "__main__":
    main()
