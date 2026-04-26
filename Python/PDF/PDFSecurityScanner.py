import os
import datetime
import shutil
from pathlib import Path

try:
    from PyPDF2 import PdfReader, PdfWriter
    # PyPDF2 3.0+ uses different permission constants
    PERMISSIONS_ALL = -1  # Allow all permissions
except ImportError:
    print("PyPDF2 not found. Install with: pip install PyPDF2")
    exit(1)

def get_user_inputs():
    """Get configuration from user input"""
    print("=== PDF Security Tool ===\n")
    
    # Get directory to scan
    while True:
        root_folder = input("Enter the directory path to scan for PDFs: ").strip()
        if os.path.exists(root_folder):
            break
        print("❌ Directory not found. Please try again.")
    
    # Get password options
    print("\nPassword Configuration:")
    user_password = input("Enter user password (leave empty for no password): ").strip()
    
    while True:
        owner_password = input("Enter owner password (required): ").strip()
        if owner_password:
            break
        print("❌ Owner password is required.")
    
    # Backup option
    backup = input("\nCreate backup copies? (y/n): ").lower().startswith('y')
    
    return root_folder, user_password, owner_password, backup

def setup_logging(root_folder):
    """Setup logging configuration"""
    log_dir = os.path.join(root_folder, "pdf_security_logs")
    os.makedirs(log_dir, exist_ok=True)
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_path = os.path.join(log_dir, f"pdf_security_log_{timestamp}.txt")
    
    return log_file_path

def log_message(log_file_path, message):
    """Log message to file and console"""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"{timestamp} {message}"
    
    with open(log_file_path, "a", encoding="utf-8") as log_file:
        log_file.write(log_entry + "\n")
    print(log_entry)

def create_backup(file_path, backup_dir):
    """Create backup of original file"""
    try:
        os.makedirs(backup_dir, exist_ok=True)
        backup_path = os.path.join(backup_dir, f"backup_{os.path.basename(file_path)}")
        shutil.copy2(file_path, backup_path)
        return backup_path
    except Exception as e:
        return None

def secure_pdf(file_path, user_password, owner_password, log_file_path, backup_dir=None):
    """Apply security to a single PDF file"""
    temp_path = file_path + ".temp_secured.pdf"
    
    try:
        # Create backup if requested
        if backup_dir:
            backup_path = create_backup(file_path, backup_dir)
            if backup_path:
                log_message(log_file_path, f"📁 Backup created: {backup_path}")
        
        # Read original PDF
        reader = PdfReader(file_path)
        writer = PdfWriter()
        
        # Copy all pages
        for page in reader.pages:
            writer.add_page(page)
        
        # Apply security settings - PyPDF2 3.0+ syntax
        writer.encrypt(
            user_password=user_password if user_password else "",
            owner_password=owner_password,
            permissions_flag=PERMISSIONS_ALL  # Allow all permissions except changing security
        )
        
        # Write secured PDF to temp file
        with open(temp_path, "wb") as temp_file:
            writer.write(temp_file)
        
        # Replace original with secured version
        os.remove(file_path)
        os.rename(temp_path, file_path)
        
        log_message(log_file_path, f"✅ Secured: {file_path}")
        return True
        
    except Exception as e:
        # Clean up temp file if it exists
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass
        
        log_message(log_file_path, f"❌ Failed to secure {file_path}: {str(e)}")
        return False

def scan_and_secure_pdfs(root_folder, user_password, owner_password, backup_enabled, log_file_path):
    """Main function to scan directory and secure all PDFs"""
    
    # Setup backup directory if enabled
    backup_dir = None
    if backup_enabled:
        backup_dir = os.path.join(root_folder, "pdf_backups")
    
    # Statistics
    total_files = 0
    secured_files = 0
    failed_files = 0
    
    log_message(log_file_path, f"🔍 Starting PDF security scan in: {root_folder}")
    log_message(log_file_path, f"🔒 User password: {'Set' if user_password else 'None'}")
    log_message(log_file_path, f"🔐 Owner password: Set")
    log_message(log_file_path, f"💾 Backup enabled: {backup_enabled}")
    log_message(log_file_path, "=" * 50)
    
    # Walk through directory
    for dirpath, dirnames, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith('.pdf'):
                file_path = os.path.join(dirpath, filename)
                total_files += 1
                
                log_message(log_file_path, f"📄 Processing: {file_path}")
                
                if secure_pdf(file_path, user_password, owner_password, log_file_path, backup_dir):
                    secured_files += 1
                else:
                    failed_files += 1
    
    # Summary
    log_message(log_file_path, "=" * 50)
    log_message(log_file_path, "📊 SUMMARY:")
    log_message(log_file_path, f"   Total PDFs found: {total_files}")
    log_message(log_file_path, f"   Successfully secured: {secured_files}")
    log_message(log_file_path, f"   Failed: {failed_files}")
    log_message(log_file_path, f"   Success rate: {(secured_files/total_files*100):.1f}%" if total_files > 0 else "   Success rate: N/A")
    
    return total_files, secured_files, failed_files

def test_pypdf2():
    """Test PyPDF2 functionality"""
    try:
        reader = PdfReader
        writer = PdfWriter
        print("✅ PyPDF2 imports successfully!")
        return True
    except Exception as e:
        print(f"❌ PyPDF2 test failed: {e}")
        return False

def main():
    """Main execution function"""
    try:
        # Test PyPDF2 first
        print("🧪 Testing PyPDF2 installation...")
        if not test_pypdf2():
            print("❌ PyPDF2 is not working properly. Try reinstalling with:")
            print("   pip uninstall PyPDF2")
            print("   pip install PyPDF2")
            return
        
        # Get user inputs
        root_folder, user_password, owner_password, backup_enabled = get_user_inputs()
        
        # Setup logging
        log_file_path = setup_logging(root_folder)
        
        # Confirm before proceeding
        print(f"\n📋 Configuration Summary:")
        print(f"   Directory: {root_folder}")
        print(f"   User password: {'Set' if user_password else 'None (no password to open)'}")
        print(f"   Owner password: Set")
        print(f"   Create backups: {backup_enabled}")
        print(f"   Log file: {log_file_path}")
        
        confirm = input("\nProceed with securing PDFs? (y/n): ").lower()
        if not confirm.startswith('y'):
            print("Operation cancelled.")
            return
        
        # Execute the security process
        print("\n🚀 Starting PDF security process...\n")
        total, secured, failed = scan_and_secure_pdfs(
            root_folder, user_password, owner_password, backup_enabled, log_file_path
        )
        
        print(f"\n✨ Process completed!")
        print(f"📊 Results: {secured}/{total} PDFs secured successfully")
        if failed > 0:
            print(f"⚠️  {failed} files failed - check log for details")
        print(f"📝 Full log available at: {log_file_path}")
        
    except KeyboardInterrupt:
        print("\n\n⛔ Operation cancelled by user.")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")

if __name__ == "__main__":
    main()