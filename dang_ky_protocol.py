import winreg
import sys
import os

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def register_tt38_protocol():
    script_path = os.path.join(r"d:\Cong ty\DuToanChuyenNghiep", "mo_trang_tt38.py")
    uv_path = r"C:\Users\DELL\.local\bin\uv.exe"
    
    # Command to run when tt38:// is clicked
    # uv run python "d:\Cong ty\DuToanChuyenNghiep\mo_trang_tt38.py" "%1"
    command_str = f'"{uv_path}" run python "{script_path}" "%1"'
    
    key_path = r"Software\Classes\tt38"
    
    try:
        # Create HKCU\Software\Classes\tt38
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "URL:TT38 Protocol")
            winreg.SetValueEx(key, "URL Protocol", 0, winreg.REG_SZ, "")
            
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r"\shell\open\command") as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command_str)
            
        print("Đăng ký thành công giao thức tt38:// vào Windows (HKCU)!")
        print("Command:", command_str)
        return True
    except Exception as e:
        print("Lỗi khi đăng ký Registry:", e)
        return False

if __name__ == "__main__":
    register_tt38_protocol()
