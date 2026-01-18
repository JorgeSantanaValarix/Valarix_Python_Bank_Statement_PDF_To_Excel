# Configuration and Installation - PDF to Excel

This document explains how to configure and install all necessary dependencies to run `pdf_to_excel.py` on Windows Server.

## What does this script do?

`pdf_to_excel.py` converts bank statements from PDF format to Excel (.xlsx) files with multiple organized sheets.

## System Requirements

- Windows Server or Windows 10/11
- Python 3.7 or higher
- Internet connection (to download dependencies)
- Administrator permissions (optional, only for installing Tesseract OCR)

---

## Prerequisites

### Python

- **Minimum version:** Python 3.7
- **Download:** https://www.python.org/downloads/
- **IMPORTANT:** During installation, check the option **"Add Python to PATH"**

### Permissions

- Read/write permissions in the project folder
- Administrator permissions (optional, only if you want to install Tesseract automatically with Chocolatey)

---

## Quick Installation (Recommended)

### Step 1: Clone the repository

```bash
git clone <repository-url>
cd Valarix_Python_Bank_Statement_PDF_To_Excel
```

### Step 2: Run the installation script

```bash
setup.bat
```

The script will do everything automatically:
1. ✅ Verify that Python is installed
2. ✅ Update pip
3. ✅ Install all Python dependencies
4. ✅ Attempt to install Tesseract OCR (if Chocolatey is available)
5. ✅ Download English and Spanish language packs for Tesseract
6. ✅ Verify that everything is correct

**Estimated time:** 5-10 minutes

---

## Manual Installation

If the automatic script fails or you prefer to install manually, follow these steps:

### Step 1: Install Python

1. Download Python from: https://www.python.org/downloads/
2. Run the installer
3. **IMPORTANT:** Check "Add Python to PATH"
4. Complete the installation
5. Verify the installation:
   ```bash
   python --version
   ```

### Step 2: Install Python Dependencies

```bash
# Update pip
python -m pip install --upgrade pip

# Install dependencies
python -m pip install -r requirements.txt
```

**Dependencies that will be installed:**
- `pdfplumber` - PDF text extraction
- `pandas` - Data manipulation
- `openpyxl` - Excel file writing
- `pymupdf>=1.23.0` - PDF to image conversion for OCR
- `pytesseract>=0.3.10` - Python wrapper for Tesseract OCR
- `Pillow>=10.0.0` - Image processing

### Step 3: Install Tesseract OCR

#### Option A: With Chocolatey (Recommended)

If you have Chocolatey installed:

```powershell
choco install tesseract
```

#### Option B: Windows Installer

1. Go to: https://github.com/UB-Mannheim/tesseract/wiki
2. Download the latest installer:
   - **64-bit:** `tesseract-ocr-w64-setup-5.x.x.exe`
   - **32-bit:** `tesseract-ocr-w32-setup-5.x.x.exe`
3. Run the installer
4. **IMPORTANT:** During installation:
   - ✅ Check **"Add Tesseract to PATH"**
   - ✅ Select **"Spanish"** in additional languages
   - ✅ Select **"English"** as well
5. Complete the installation
6. **CLOSE and REOPEN** PowerShell/CMD so PATH updates
7. Verify the installation:
   ```bash
   tesseract --version
   ```

You should see something like:
```
tesseract 5.3.0
 leptonica-1.83.0
```

---

## Installation Verification

### Automatic Verification

The `setup.bat` script includes automatic verification. If you want to verify manually:

```bash
python -c "import pdfplumber, pandas, openpyxl, fitz, pytesseract, PIL; print('✅ All dependencies installed')"
```

### Verify Tesseract OCR

```bash
tesseract --version
```

### Verify Tesseract Languages

```python
import pytesseract
langs = pytesseract.get_languages()
print(f"Available languages: {langs}")
print(f"Spanish available: {'spa' in langs}")
print(f"English available: {'eng' in langs}")
```

---

## Script Usage

### Integration Example

Call `pdf_to_excel.py` from C#:

```csharp
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

// TODO: Get PDF path from database
string pdfPath = GetPdfPathFromDatabase();

// Execute Python script
var startInfo = new ProcessStartInfo
{
    FileName = "python",
    Arguments = $"\"pdf_to_excel.py\" \"{pdfPath}\"",
    RedirectStandardOutput = true,
    RedirectStandardError = true,
    UseShellExecute = false,
    CreateNoWindow = true
};

using (var process = Process.Start(startInfo))
{
    string output = process.StandardOutput.ReadToEnd();
    string error = process.StandardError.ReadToEnd();
    process.WaitForExit();
    
    if (process.ExitCode == 0)
    {
        // Extract Excel path from output: "✅ Excel file created successfully -> {path}"
        var match = Regex.Match(output, @"Excel file created successfully -> (.+)");
        string excelPath = match.Success ? match.Groups[1].Value.Trim() : Path.ChangeExtension(pdfPath, ".xlsx");
        
        // TODO: Update database with excelPath and status 'Done'
        UpdateDatabaseWithExcelPath(pdfPath, excelPath, "Done", null);
    }
    else
    {
        // TODO: Update database with status 'Failed' and error message
        UpdateDatabaseWithExcelPath(pdfPath, null, "Failed", error);
    }
}

// Dummy method: Get PDF path from database
static string GetPdfPathFromDatabase()
{
    // TODO: Replace with actual database query
    // Example SQL: SELECT PDFTOOLPATH FROM BANK_STTEMENT_SEARCH_HISTORY_FORWINDOWSERVICE WHERE Status = 'Pending'
    return @"Test\Bank Statement\BBVA.pdf";
}

// Dummy method: Update database with Excel path and status
static void UpdateDatabaseWithExcelPath(string pdfPath, string excelPath, string status, string errorMessage)
{
    // TODO: Replace with actual database update
    // Example SQL: UPDATE BANK_STTEMENT_SEARCH_HISTORY_FORWINDOWSERVICE 
    // SET FILEPATH = @excelPath, Status = @status, ErrorMessage = @errorMessage
    // WHERE PDFTOOLPATH = @pdfPath
}
```

### Script Output Format

The Python script prints the following to stdout, which C# can capture:

- `Reading PDF...`
- `🏦 Bank detected: [bank name]`
- `📊 Exporting to Excel...`
- `✅ VALIDATION: ALL CORRECT` (or `THERE ARE DIFFERENCES`)
- `✅ Excel file created successfully -> [full path to Excel file]`

The last line is the most important, as it contains the Excel file path that C# extracts using regex.

### Exit Codes

The script returns the following exit codes:
- `0` = Success (Excel created correctly)
- `1` = Error (check error messages in stderr)

### Direct Command Line Usage (Optional)

You can also run the script directly from command line for testing:

```bash
python pdf_to_excel.py "path\to\file.pdf"
```

The script will generate an Excel file with the same name as the PDF but with `.xlsx` extension.

**Example:**
```bash
python pdf_to_excel.py "Test\Bank Statement\BBVA.pdf"
```

This will generate: `Test\Bank Statement\BBVA.xlsx`

---

## Troubleshooting

### Error: "Python not found"

**Problem:** Python is not installed or not in PATH.

**Solution:**
1. Install Python from https://www.python.org/downloads/
2. Make sure to check "Add Python to PATH" during installation
3. Close and reopen PowerShell/CMD
4. Verify with: `python --version`

### Error: "pip not found"

**Problem:** pip is not installed or not in PATH.

**Solution:**
```bash
python -m ensurepip --upgrade
```

### Error: "Tesseract not found"

**Problem:** Tesseract OCR is not installed or not in PATH.

**Solution:**
1. Install Tesseract following the instructions in the "Manual Installation" section
2. Make sure to check "Add Tesseract to PATH" during installation
3. **CLOSE and REOPEN** PowerShell/CMD
4. Verify with: `tesseract --version`

### Error: "Spanish (spa) NOT available"

**Problem:** The Spanish language is not installed in Tesseract.

**Solution:**
1. Reinstall Tesseract
2. During installation, select "Spanish" in additional languages
3. Or run `setup.bat` again - it will automatically download missing language packs

### Error installing Python dependencies

**Problem:** Connection error or permissions.

**Solution:**
1. Verify your Internet connection
2. Try updating pip: `python -m pip install --upgrade pip`
3. If you use a proxy, configure pip:
   ```bash
   pip config set global.proxy http://proxy:port
   ```
4. Run PowerShell/CMD as administrator

### Script works but without OCR

**Problem:** Tesseract is not installed or not in PATH.

**Solution:**
- The script will work but only process PDFs with legible text
- For full OCR support, install Tesseract following the instructions

### Error: "requirements.txt not found"

**Problem:** The script was run from an incorrect directory.

**Solution:**
1. Make sure to run `setup.bat` from the project root
2. Verify that `requirements.txt` is in the same directory

---

## Frequently Asked Questions

### Do I need to install Tesseract OCR?

**Answer:** Not strictly necessary, but highly recommended. Without Tesseract:
- ✅ The script will work for PDFs with legible text
- ❌ It will not be able to process PDFs with scanned or illegible text

### Can I use the script without Internet connection?

**Answer:** Once all dependencies are installed, yes. But you need Internet for:
- Installing Python dependencies (first time)
- Installing Tesseract OCR (first time)

### Does it work on Windows Server?

**Answer:** Yes, the script is designed to work on Windows Server and regular Windows.

### Do I need administrator permissions?

**Answer:** 
- **For Python dependencies:** No
- **For Tesseract OCR with Chocolatey:** Yes
- **For manual Tesseract OCR:** No (but you need installation permissions)

### Which banks are supported?

**Answer:** The script supports multiple Mexican banks:
- BBVA
- Banamex
- HSBC
- Santander
- Scotiabank
- Inbursa
- Konfio
- Clara
- Banregio
- Banorte
- Banbajío
- Base

### Can I run the script from Windows Service?

**Answer:** Yes, the script is designed to run from command line:
```bash
python pdf_to_excel.py "path\to\file.pdf"
```

This makes it compatible with Windows Services that execute commands.

**Exit codes:**
- `0` = Success (Excel created correctly)
- `1` = Error (check console messages)

Example for Windows Service:
```batch
python pdf_to_excel.py "\\network\path\file.pdf"
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Processing failed
)
```

### Does the generated Excel have any branding?

**Answer:** Yes, the generated Excel has "CONTAAYUDA" as the author in the file properties. You can see it in:
- Windows: File Properties → Details → Author
- Excel: File → Info → Properties → Author

---

## Useful Links

- **Python:** https://www.python.org/downloads/
- **Tesseract OCR Windows:** https://github.com/UB-Mannheim/tesseract/wiki
- **Chocolatey:** https://chocolatey.org/
- **openpyxl documentation:** https://openpyxl.readthedocs.io/
- **pandas documentation:** https://pandas.pydata.org/

---

## Support

If you encounter problems during installation:

1. Review the **Troubleshooting** section above
2. Verify that all prerequisites are met
3. Make sure you have closed and reopened PowerShell/CMD after installing Python or Tesseract
4. Run `setup.bat` again to verify the installation

---

## Next Steps

Once installation is complete:

1. Test the script with a test PDF:
   ```bash
   python pdf_to_excel.py "Test\Bank Statement\BBVA.pdf"
   ```

2. Verify that the Excel was generated correctly

3. Review the Excel sheets:
   - Summary
   - Bank Statement Report
   - Data Validation

Ready to use! 🎉
