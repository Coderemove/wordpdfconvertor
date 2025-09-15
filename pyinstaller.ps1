# PyInstaller build script for PDF/DOCX Converter GUI
# Run this script in PowerShell to build a standalone Windows .exe

$python = "C:/Users/Code/Desktop/CV/.venv312/Scripts/python.exe"

# Activate virtual environment
& .venv312\Scripts\Activate.ps1

# Kill any existing convert.exe processes
Write-Host "Killing any existing convert.exe processes..."
Get-Process convert* | Stop-Process -Force -ErrorAction SilentlyContinue

# Install dependencies if not already installed
Write-Host "Installing PyInstaller and project dependencies..."
& $python -m pip install --upgrade pip
& $python -m pip install pyinstaller pdf2docx docx2pdf

$script = "convert.py"

# Build command: onedir for folder with exe, windowed to hide console, with icon and hidden imports for COM
& pyinstaller --onedir --windowed --icon=icon.ico --hidden-import pywintypes --hidden-import win32com --distpath dist $script