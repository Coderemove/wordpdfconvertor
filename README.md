# CV Converter

This project provides a script/exe to convert CVs or related documents. It is designed for Windows usage.

## Version

**1.0.0**

## Features
- Converts CV files (details depend on `convert.py` implementation)
- Windows compatible

## Dependencies
- Python 3.7+
- Microsoft Word (for document handling)
- [PyInstaller](https://www.pyinstaller.org/) (for building executable)
- Any additional dependencies listed in `convert.py` (see code for details)

## Installation
1. Install Python 3.7 or newer from [python.org](https://www.python.org/downloads/windows/).
2. Install required packages:
   ```powershell
   pip install pyinstaller
   # Add any other dependencies here
   ```

## Building for Windows
To create a standalone Windows executable:

1. Open PowerShell in the project directory.
2. Run the provided build script:
   ```powershell
   .\pyinstaller.ps1
   ```
   Or, manually run:
   ```powershell
   pyinstaller --onefile --icon=icon.ico convert.py
   ```
3. The executable will be in the `dist` folder.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
