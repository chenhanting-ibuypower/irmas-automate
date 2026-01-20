# 🚀 **Project Initialization & Packaging Guide**

This guide explains how to set up the Python environment, manage dependencies, configure environment variables, install Playwright browsers, and package your project into a standalone executable.

---

# 1️⃣ **Create a Virtual Environment**

Choose **one** command depending on your system:

```sh
python -m venv venv
# or
python3 -m venv venv
```

---

# 2️⃣ **Activate the Virtual Environment**

### **Windows**

```sh
venv\Scripts\activate
```

### **macOS / Linux**

```sh
source venv/bin/activate
```

---

# 3️⃣ **Upgrade pip & Tooling**

```sh
python -m pip install --upgrade pip setuptools wheel
```

---

# 4️⃣ **Install Required Packages**

Example:

```sh
pip install python-dotenv
pip install playwright
```

Or install from an existing requirements file:

```sh
pip install -r requirements.txt
```

---

# 5️⃣ **Update `requirements.txt` (Best Practice)**

Always run this **after installing new packages**:

```sh
pip freeze > requirements.txt
```

---

# 6️⃣ **Install Playwright Browsers**

### Install all browsers

```sh
playwright install
```

### Install only Chrome

```sh
playwright install chrome
```

### Install Chromium into the project directory (required for PyInstaller)

**Windows PowerShell**

```powershell
$env:PLAYWRIGHT_BROWSERS_PATH=0; playwright install --force chromium
```

---

# 7️⃣ **Build a Standalone Executable with PyInstaller**

### Basic one-file build

```sh
pyinstaller --onefile --distpath "D:\rpa" hello.py
```

### Example using another script

```sh
pyinstaller -F --distpath "D:\rpa" .\visit_sites.py
```

---

# 8️⃣ **Bundle Playwright Browser Files (Important for Packaging)**

Playwright stores browser binaries here:

```
C:\Users\<user>\AppData\Local\ms-playwright\
```

Include these binaries when building:

```sh
pyinstaller --onefile --add-data "C:/Users/user/AppData/Local/ms-playwright/*;ms-playwright" visit_sites.py
```

This ensures your packaged EXE can run Chromium.

---

# 9️⃣ **Environment Variable Setup (`.env` / `.env.example`)**

This project uses environment variables to securely store credentials.

### File Structure

```
project/
│── .env            # Your local secrets (NOT committed)
│── .env.example    # Template for developers
│── app.py          # Main script
└── ...
```

### `.env.example`

```
EMS_ACCOUNT=
EMS_PASSWORD=
EMS_CARD_PASSWORD=
```

Create your actual `.env` file:

```sh
cp .env.example .env
```

Then fill in your values.

---

# 🔟 **Check Which Python Version Your venv Uses**

Run:

```sh
venv\Scripts\python.exe --version
```

Expected:

```
Python 3.12.x
```

---

# 1️⃣1️⃣ **Fixing Version Mismatches (If Needed)**

1. Activate the correct venv:

```powershell
.\venv\Scripts\activate
```

2. Verify the version:

```sh
python --version
```

Expected:

```
Python 3.12.x
```

---

# 🎯 **Recommended Daily Workflow**

1. Install a new package

   ```sh
   pip install SOME_PACKAGE
   ```

2. Update requirements

   ```sh
   pip freeze > requirements.txt
   ```

### ⚠ Excluded Banned Software Terms

Some software names such as **"Steam"** and **"Tor"** are not included in `banned_software.json` because they produce false positives when matched as substrings.

- `"Steam"` may match `"MSteams"`
- `"Tor"` may match `"store"`

To avoid incorrect detection, these terms are handled in code using **exact match**, **regex word boundaries**, or stricter matching logic. They are intentionally excluded from the configuration file to ensure accurate scanning results.

py -3.12 -m pip install -r requirements.txt

rmdir /s /q build
rmdir /s /q dist
py -3.12 -m PyInstaller IRMAS-AUTOMATE.spec --clean --noconfirm

---

# 1️⃣2️⃣ **Clean Build & Run (One-Liner)**

Use this when you want a fresh build and to run the resulting EXE:

```powershell
Remove-Item -Path "build","dist" -Recurse -Force -ErrorAction SilentlyContinue; py -3.12 -m PyInstaller IRMAS-AUTOMATE.spec --clean --noconfirm
```

Then run the executable:

```powershell
.\dist\IRMAS-AUTOMATE\IRMAS-AUTOMATE.exe
```

**Offline build note:** The build does **not** need internet as long as your venv already has all dependencies installed and the bundled `chromium/` folder is present. Network is only required when initially running `pip install` / `playwright install`.

```powershell
python -m PyInstaller --onedir --name "IRMAS-AUTOMATE" `
--collect-all playwright `
--hidden-import playwright `
main.py
```

(backtick後面不能有空白)

```powershell
# Stop any running processes that might be locking the directories
Get-Process | Where-Object {$_.Path -like "*\dist\*" -or $_.Path -like "*\build\*"} | Stop-Process -Force -ErrorAction SilentlyContinue

# Wait a moment for processes to fully release
Start-Sleep -Seconds 1

# Remove directories with force
if (Test-Path "build") { Remove-Item -Path "build" -Recurse -Force -ErrorAction SilentlyContinue }
if (Test-Path "dist") { Remove-Item -Path "dist" -Recurse -Force -ErrorAction SilentlyContinue }

# Wait another moment
Start-Sleep -Seconds 1

# Run PyInstaller
py -3.12 -m PyInstaller IRMAS-AUTOMATE.spec --clean --noconfirm
```

---

## 🔧 **Troubleshooting: Failed to Load Python DLL**

If you encounter `Failed to load Python DLL 'python312.dll'`:

### **Solution 1: Clean Rebuild with --noupx**

```powershell
# Stop processes
Get-Process | Where-Object {$_.Path -like "*\dist\*" -or $_.Path -like "*\build\*"} | Stop-Process -Force -ErrorAction SilentlyContinue

# Clean completely
Remove-Item -Path "build","dist" -Recurse -Force -ErrorAction SilentlyContinue

# Rebuild WITHOUT UPX compression (prevents DLL corruption)
py -3.12 -m PyInstaller IRMAS-AUTOMATE.spec --clean --noconfirm --noupx
```

### **Solution 2: Install Visual C++ Redistributables**

Download and install: [VC++ Redistributable 2015-2022 (x64)](https://aka.ms/vs/17/release/vc_redist.x64.exe)

### **Solution 3: Temporarily Disable Antivirus**

Some antivirus software quarantines or corrupts Python DLLs during build. Temporarily disable it, rebuild, then re-enable.

### **Solution 4: Verify DLL After Build**

```powershell
Test-Path ".\build\IRMAS-AUTOMATE\_internal\python312.dll"
# Should return: True
```
