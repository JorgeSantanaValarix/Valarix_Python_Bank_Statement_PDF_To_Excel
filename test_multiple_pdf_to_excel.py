import os
import sys
import subprocess
import glob
from pathlib import Path
import re
import time


def find_pdf_files(folder_path: str, recursive: bool = False) -> list:
    """
    Find all PDF files in the specified folder.
    
    Args:
        folder_path: Path to the folder to search
        recursive: If True, search in subdirectories too
    
    Returns:
        List of absolute paths to PDF files
    """
    folder_path = os.path.abspath(folder_path)
    if not os.path.isdir(folder_path):
        raise ValueError(f"Folder does not exist: {folder_path}")
    
    if recursive:
        pattern = os.path.join(folder_path, "**", "*.pdf")
    else:
        pattern = os.path.join(folder_path, "*.pdf")
    
    pdf_files = []
    for pdf_path in glob.glob(pattern, recursive=recursive):
        if os.path.isfile(pdf_path):
            pdf_files.append(os.path.abspath(pdf_path))
    
    return sorted(pdf_files)


def detect_error_from_output(stderr_text: str, stdout_text: str, return_code: int) -> tuple:
    """
    Detect errors from pdf_to_excel.py output.
    
    Args:
        stderr_text: Captured stderr text from execution
        stdout_text: Captured stdout text from execution
        return_code: Return code from subprocess
    
    Returns:
        Tuple of (has_error: bool, error_type: str, error_message: str)
    """
    # Combine stdout and stderr for analysis
    combined_output = stdout_text + stderr_text
    
    if return_code != 0:
        # Check for specific error messages
        if "❌ VALIDATION: THERE ARE DIFFERENCES" in combined_output:
            return (True, "Validation error", "❌ VALIDATION: THERE ARE DIFFERENCES")
        elif "❌ Excel file not created" in combined_output:
            # Extract error message if available
            error_match = re.search(r"❌ Excel file not created.*?\n.*?Error: (.+)", combined_output, re.DOTALL)
            if error_match:
                error_msg = error_match.group(1).strip()
                # Truncate if too long
                if len(error_msg) > 200:
                    error_msg = error_msg[:200] + "..."
            else:
                error_msg = "Excel file not created"
            return (True, "Excel creation error", error_msg)
        elif "❌ Error:" in combined_output:
            # Extract first error message
            error_match = re.search(r"❌ Error: (.+?)(?:\n|$)", combined_output)
            if error_match:
                error_msg = error_match.group(1).strip()
                if len(error_msg) > 200:
                    error_msg = error_msg[:200] + "..."
            else:
                error_msg = "Unknown error"
            return (True, "General error", error_msg)
        else:
            return (True, "Unknown error", f"Process exited with code {return_code}")
    
    # Also check for validation errors even if return code is 0
    if "❌ VALIDATION: THERE ARE DIFFERENCES" in combined_output:
        return (True, "Validation error", "❌ VALIDATION: THERE ARE DIFFERENCES")
    
    return (False, "", "")


def process_single_pdf(pdf_path: str, script_path: str = "pdf_to_excel.py") -> tuple:
    """
    Process a single PDF file by executing pdf_to_excel.py.
    Shows output in real-time and captures errors.
    
    Args:
        pdf_path: Path to the PDF file to process
        script_path: Path to pdf_to_excel.py script
    
    Returns:
        Tuple of (success: bool, error_type: str, error_message: str, elapsed_time: float)
    """
    start_time = time.time()
    
    try:
        # Get absolute paths
        pdf_path = os.path.abspath(pdf_path)
        script_path = os.path.abspath(script_path)
        
        # Validate script exists
        if not os.path.isfile(script_path):
            elapsed_time = time.time() - start_time
            return (False, "Script not found", f"pdf_to_excel.py not found at: {script_path}", elapsed_time)
        
        # Build command
        cmd = [sys.executable, script_path, pdf_path]
        
        # Execute with real-time output
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,  # Merge stderr into stdout
            text=True,
            bufsize=1,  # Line buffered
            encoding='utf-8',
            errors='replace',
            cwd=os.path.dirname(script_path) or os.getcwd()
        )
        
        # Read output line by line and print in real-time
        stdout_lines = []
        for line in process.stdout:
            print(line, end='', flush=True)  # Print in real-time
            stdout_lines.append(line)
        
        # Wait for process to complete
        return_code = process.wait()
        
        # Calculate elapsed time
        elapsed_time = time.time() - start_time
        
        # Get all output
        stdout_text = ''.join(stdout_lines)
        stderr_text = ""  # Already merged into stdout
        
        # Detect errors
        has_error, error_type, error_message = detect_error_from_output(stderr_text, stdout_text, return_code)
        
        if has_error:
            return (False, error_type, error_message, elapsed_time)
        else:
            return (True, "", "", elapsed_time)
    
    except Exception as e:
        elapsed_time = time.time() - start_time
        return (False, "Execution error", str(e), elapsed_time)


def format_time(seconds: float) -> str:
    """
    Format time in seconds to a human-readable string.
    
    Args:
        seconds: Time in seconds
    
    Returns:
        Formatted time string (e.g., "2m 30.5s" or "45.2s")
    """
    if seconds < 60:
        return f"{seconds:.2f}s"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        secs = seconds % 60
        return f"{minutes}m {secs:.2f}s"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = seconds % 60
        return f"{hours}h {minutes}m {secs:.2f}s"


def process_folder(folder_path: str, recursive: bool = False) -> dict:
    """
    Process all PDF files in a folder.
    
    Args:
        folder_path: Path to folder containing PDFs
        recursive: If True, process PDFs in subdirectories too
    
    Returns:
        Dictionary with statistics including failed list and total time
    """
    stats = {
        'total': 0,
        'successful': 0,
        'failed': 0,
        'failed_list': [],  # List of dicts: {'file': str, 'error_type': str, 'error_message': str}
        'total_time': 0.0
    }
    
    # Find PDFs
    pdf_files = find_pdf_files(folder_path, recursive=recursive)
    stats['total'] = len(pdf_files)
    
    if stats['total'] == 0:
        print(f"⚠️  No PDF files found in: {folder_path}")
        return stats
    
    print(f"📁 Processing folder: {folder_path}")
    print(f"📄 Found {stats['total']} PDF file(s)\n")
    
    # Start total timer
    total_start_time = time.time()
    
    # Process each PDF
    for idx, pdf_path in enumerate(pdf_files, 1):
        pdf_name = os.path.basename(pdf_path)
        
        # Separator
        print("=" * 60)
        print(f"[{idx}/{stats['total']}] Processing: {pdf_name}")
        print("=" * 60)
        
        success, error_type, error_message, elapsed_time = process_single_pdf(pdf_path)
        
        # Separator after processing
        print("=" * 60)
        
        # Print time for this PDF
        time_str = format_time(elapsed_time)
        print(f"⏱️  Time: {time_str}")
        
        if success:
            print(f"✅ Success: {pdf_name}\n")
            stats['successful'] += 1
        else:
            print(f"❌ Failed: {pdf_name} ({error_type})\n")
            stats['failed'] += 1
            stats['failed_list'].append({
                'file': pdf_name,
                'error_type': error_type,
                'error_message': error_message
            })
    
    # Calculate total time
    stats['total_time'] = time.time() - total_start_time
    
    return stats


def print_summary(stats: dict):
    """
    Print final summary with statistics.
    
    Args:
        stats: Dictionary with processing statistics
    """
    total = stats['total']
    successful = stats['successful']
    failed = stats['failed']
    total_time = stats.get('total_time', 0.0)
    
    # Calculate success rate
    if total > 0:
        success_rate = (successful / total) * 100
    else:
        success_rate = 0.0
    
    print("\n" + "=" * 60)
    print("📊 FINAL SUMMARY")
    print("=" * 60)
    print(f"Total PDFs processed: {total}")
    print(f"✅ Successful: {successful}")
    print(f"❌ Failed: {failed}")
    print(f"📈 Success rate: {success_rate:.1f}%")
    print(f"⏱️  Total time: {format_time(total_time)}")
    
    if stats['failed_list']:
        print(f"\n❌ Failed PDFs:")
        for idx, failed_item in enumerate(stats['failed_list'], 1):
            print(f"   {idx}. {failed_item['file']}")
            print(f"      Error: {failed_item['error_type']}")
            if failed_item['error_message']:
                # Truncate long error messages
                error_msg = failed_item['error_message']
                if len(error_msg) > 100:
                    error_msg = error_msg[:100] + "..."
                print(f"      Details: {error_msg}")
    
    print("=" * 60)


def main():
    """Main function to handle command line arguments and execute processing."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Process all PDF files in a folder using pdf_to_excel.py'
    )
    parser.add_argument(
        'folder_path',
        nargs='?',  # Make it optional
        help='Path to folder containing PDF files'
    )
    parser.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='Process PDFs in subdirectories too'
    )
    
    args = parser.parse_args()
    
    # Get folder path
    if args.folder_path:
        folder_path = args.folder_path
    else:
        folder_path = input("Enter folder path: ").strip()
    
    # Validate folder
    folder_path = os.path.abspath(folder_path)
    if not os.path.isdir(folder_path):
        print(f"❌ Error: Folder does not exist: {folder_path}")
        sys.exit(1)
    
    # Process folder
    try:
        stats = process_folder(folder_path, recursive=args.recursive)
        
        # Print summary
        print_summary(stats)
        
        # Exit with error code if any failed
        if stats['failed'] > 0:
            sys.exit(1)
    
    except KeyboardInterrupt:
        print("\n\n⚠️  Process interrupted by user")
        sys.exit(130)
    except Exception as e:
        print(f"❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
