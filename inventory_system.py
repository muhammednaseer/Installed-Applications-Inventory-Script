import os
import json
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

 
# CONFIG
 

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PC_FILE = os.path.join(SCRIPT_DIR, "pc_namelist.txt")

OUTPUT_FILE = # give the output location where data excel sheet should be saved.
MAX_THREADS = 10


 
# CLEAN FUNCTION
 

def clean(value):
    if value is None:
        return ""
    return ILLEGAL_CHARACTERS_RE.sub("", str(value))


 
# LOAD PCS
 

def load_pcs():
    if not os.path.exists(PC_FILE):
        print(f"PC file not found: {PC_FILE}")
        return []

    with open(PC_FILE, "r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

  
# GET SOFTWARE FROM PC
 

def get_software(pc):

    # Ping check
    ping = subprocess.run(
        ["ping", "-n", "1", "-w", "1000", pc],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    if ping.returncode != 0:
        return "OFFLINE", [], "Ping failed"

    ps_script = rf"""
    Invoke-Command -ComputerName "{pc}" -ScriptBlock {{

        $ErrorActionPreference = "SilentlyContinue"

        $apps = Get-ItemProperty `
            HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
            HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
            Where-Object {{ $_.DisplayName }} |
            Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

        $apps | ConvertTo-Json -Depth 3
    }}
    """

    try:
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-Command", ps_script,
            ],
            capture_output=True,
            text=True,
            timeout=45,
        )

        output = result.stdout.strip()

        if result.returncode != 0:
            return "FAILED", [], (result.stderr.strip() or "PowerShell error")

        if not output:
            return "FAILED", [], "No data returned"

        try:
            data = json.loads(output)
        except json.JSONDecodeError:
            return "FAILED", [], "Invalid JSON output"

        if isinstance(data, dict):
            data = [data]

        rows = []
        for app in data:
            rows.append([
                pc,
                app.get("DisplayName", ""),
                app.get("DisplayVersion", ""),
                app.get("Publisher", ""),
                app.get("InstallDate", "")
            ])

        return "SUCCESS", rows, ""

    except subprocess.TimeoutExpired:
        return "TIMEOUT", [], "PowerShell timeout (45s)"

    except Exception as e:
        return "ERROR", [], str(e)
 
# SAVE EXCEL


def save_excel(rows, pcs, success_pcs, failed_pcs):

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventory"

    # ---------------- Inventory ----------------
    ws.append([
        "Computer",
        "Application",
        "Version",
        "Publisher",
        "Install Date"
    ])

    for row in rows:
        ws.append([clean(x) for x in row])

    # ---------------- REPORT SECTION ----------------
    
    rs = wb.create_sheet("Report")

    rs.append(["INVENTORY REPORT"])
    rs.append([])

    rs.append(["Total PCs", len(pcs)])
    rs.append(["Successful PCs", len(success_pcs)])
    rs.append(["Failed PCs", len(failed_pcs)])
    rs.append(["Software Records", len(rows)])

    rs.append([])
    rs.append(["FAILED PCS"])
    rs.append(["Computer", "Status", "Reason"])

    for pc, status, reason in failed_pcs:
        rs.append([clean(pc), clean(status), clean(reason)])

    wb.save(OUTPUT_FILE)


 
# MAIN
 

def main():

    pcs = load_pcs()
    if not pcs:
        print("No PCs found.")
        return

    print(f"Loaded {len(pcs)} computers\n")

    all_rows = []
    success_pcs = []
    failed_pcs = []

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(get_software, pc): pc for pc in pcs}

        completed = 0

        for future in as_completed(futures):
            pc = futures[future]
            completed += 1

            try:
                status, rows, reason = future.result()

                if status == "SUCCESS":
                    success_pcs.append(pc)
                    all_rows.extend(rows)
                else:
                    failed_pcs.append((pc, status, reason))

                print(f"[{completed}/{len(pcs)}] {pc:<20} {status}")

            except Exception as e:
                failed_pcs.append((pc, "ERROR", str(e)))
                print(f"[{completed}/{len(pcs)}] {pc:<20} ERROR")

    save_excel(all_rows, pcs, success_pcs, failed_pcs)

    # ---------------- REPORT ----------------
    print("\n===== INVENTORY REPORT =====")
    print(f"Total PCs     : {len(pcs)}")
    print(f"Success       : {len(success_pcs)}")
    print(f"Failed        : {len(failed_pcs)}")
    print(f"Records       : {len(all_rows)}")
    print(f"Output File   : {OUTPUT_FILE}")

    print("\nFAILED PCS")
    for pc, status, reason in failed_pcs:
        print(f"{pc:<20} {status:<10} {reason}")


if __name__ == "__main__":
    main()