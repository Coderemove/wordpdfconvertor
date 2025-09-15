# PyInstaller build script for Windows

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Always work from the script directory
Set-Location -Path $PSScriptRoot

# Paths
$venvPath = Join-Path $PSScriptRoot '.venv312'
$python   = Join-Path $venvPath 'Scripts\python.exe'
$script   = Join-Path $PSScriptRoot 'convert.py'
$distDir  = Join-Path $PSScriptRoot 'dist'
$iconPath = Join-Path $PSScriptRoot 'icon.ico'

function RunPy {
    param([Parameter(Mandatory=$true)][string[]]$Args)
    & $python @Args
    if ($LASTEXITCODE -ne 0) { throw "python $($Args -join ' ') failed with exit code $LASTEXITCODE" }
}

function RunPyCode {
    param([Parameter(Mandatory=$true)][string]$Code)
    $tmp = [System.IO.Path]::GetTempFileName()
    Set-Content -Path $tmp -Value $Code -Encoding UTF8
    & $python $tmp
    $rc = $LASTEXITCODE
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    if ($rc -ne 0) { throw "inline python failed with exit code $rc" }
}

# Ensure venv exists
if (-not (Test-Path $python)) {
    Write-Host "Creating virtual environment at $venvPath ..."
    if (Get-Command py -ErrorAction SilentlyContinue) {
        & py -3 -m venv $venvPath
    } else {
        & python -m venv $venvPath
    }
    if (-not (Test-Path $python)) { throw "Failed to create venv at $venvPath" }
}

# Kill any running convert.exe
Write-Host "Stopping any running convert.exe ..."
Get-Process convert -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# Ensure pip and tools
Write-Host "Upgrading pip/setuptools/wheel ..."
RunPy @('-m','ensurepip','--upgrade')
RunPy @('-m','pip','install','--upgrade','pip','setuptools','wheel')

# Install build/runtime deps in venv
Write-Host "Installing dependencies (PyInstaller, pdf2docx, docx2pdf, pywin32) ..."
RunPy @('-m','pip','install','pyinstaller','pdf2docx','docx2pdf','pywin32')

# Verify PyInstaller is importable in this venv without -c quoting issues
RunPyCode 'import PyInstaller, sys; print("PyInstaller OK", PyInstaller.__version__)'

# Build
Write-Host "Building executable ..."
$pyiArgs = @('-m','PyInstaller','--onedir','--windowed','--distpath', $distDir,'--hidden-import','pywintypes','--hidden-import','win32com')
if (Test-Path $iconPath) { $pyiArgs += @('--icon', $iconPath) }
$pyiArgs += @($script)
RunPy $pyiArgs

Write-Host "Build complete. See: $distDir"