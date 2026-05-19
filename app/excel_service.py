import sys
import subprocess
import openpyxl
from .config import EXCEL_PATH


def get_portal_name(choice):
    mapping = {
        "1": "Linkedin",
        "2": "Direct",
        "3": "Website",
        "4": "Email",
        "5": "Linkedin Easy Apply"
    }
    return mapping.get(choice, "Unknown")


def append_to_excel(row_data):
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
        ws.append(row_data)
        wb.save(EXCEL_PATH)
        return True
    except Exception as e:
        print(f"✗ Error updating Excel: {e}")
        return False


def close_excel_tracker_if_open():
    try:
        ps_command = r"""
        Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                if ($_.MainWindowTitle -like "*Job List Mustafa*") {
                    Write-Output "Closing Excel workbook window: $($_.MainWindowTitle)"
                    Stop-Process -Id $_.Id -Force
                }
            }
            catch {}
        }
        """
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
            capture_output=True,
            text=True
        )

        if result.stdout.strip():
            print(result.stdout.strip())
        elif result.returncode == 0:
            print("✓ Excel tracker was not open.")
        else:
            print(f"⚠ PowerShell returned code {result.returncode}")
            if result.stderr.strip():
                print(result.stderr.strip())

    except Exception as e:
        print(f"⚠ Could not check/close Excel tracker: {e}")
