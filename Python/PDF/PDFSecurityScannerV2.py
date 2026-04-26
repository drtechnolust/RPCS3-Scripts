import os
import datetime
import shutil
from pathlib import Path

try:
    from PyPDF2 import PdfReader, PdfWriter
except ImportError:
    print("PyPDF2 not found. Install with: pip install PyPDF2")
    exit(1)

def get_permission_settings():
    """Get detailed permission settings from user"""
    print("\n🔒 PDF Permission Settings:")
    print("Choose what users can do with the secured PDFs:")
    
    permissions = {}
    
    # Define permission options with descriptions
    permission_options = [
        ("printing", "Allow printing", True),
        ("copying", "Allow text/content copying", True),
        ("commenting", "Allow adding comments/annotations", True),
        ("form_filling", "Allow filling form fields", True),
        ("signing", "Allow digital signatures", True),
        ("accessibility", "Allow screen reader access", True),
        ("document_assembly", "Allow page insertion/deletion/rotation", False),
        ("high_quality_printing", "Allow high-quality printing", True),
        ("page_extraction", "Allow page extraction", False),
        ("document_changes", "Allow document editing/changes", False)
    ]
    
    for perm_key, description, default in permission_options:
        while True:
            default_text = "Y" if default else "N"
            response = input(f"   {description}? (y/n) [default: {default_text}]: ").strip().lower()
            
            if response == "":
                permissions[perm_key] = default
                break
            elif response in ['y', 'yes']:
                permissions[perm_key] = True
                break
            elif response in ['n', 'no']:
                permissions[perm_key] = False
                break
            else:
                print("     Please enter 'y' or 'n'")
    
    return permissions

def get_security_level():
    """Quick security level presets"""
    print("\n🛡️  Security Level Presets:")
    print("1. 📖 Read-Only (view and print only)")
    print("2. 🖊️  Limited (read, print, copy, forms)")
    print("3. 🔓 Permissive (most actions allowed)")
    print("4. 🔧 Custom (choose individual permissions)")
    
    while True:
        choice = input("\nSelect security level (1-4): ").strip()
        
        if choice == "1":  # Read-only
            return {
                "printing": True,
                "copying": False,
                "commenting": False,
                "form_filling": False,
                "signing": False,
                "accessibility": True,
                "document_assembly": False,
                "high_quality_printing": True,
                "page_extraction": False,
                "document_changes": False
            }
        elif choice == "2":  # Limited
            return {
                "printing": True,
                "copying": True,
                "commenting": False,
                "form_filling": True,
                "signing": True,
                "accessibility": True,
                "document_assembly": False,
                "high_quality_printing": True,
                "page_extraction": False,
                "document_changes": False
            }
        elif choice == "3":  # Permissive
            return {
                "printing": True,
                "copying": True,
                "commenting": True,
                "form_filling": True,
                "signing": True,
                "accessibility": True,
                "document_assembly": True,
                "high_quality_printing": True,
                "page_extraction": False,  # Still restrict this
                "document_changes": False  # Still restrict this
            }
        elif choice == "4":  # Custom
            return get_permission_settings()
        else:
            print("Please enter 1, 2, 3, or 4")

def permissions_to_flag(permissions):
    """Convert permission dictionary to PyPDF2 flag"""
    # PyPDF2 uses bitwise flags for permissions
    # Start with no permissions (0)
    flag = 0
    
    # Add permissions based on user choices
    if permissions.get("printing", False):
        flag |= 4  # Printing
    if permissions.get("document_changes", False):
        flag |= 8  # Modifying contents
    if permissions.get("copying", False):
        flag |= 16  # Copying/extracting text and graphics
    if permissions.get("commenting", False):
        flag |= 32  # Adding/modifying annotations and form fields
    if permissions.get("form_filling", False):
        flag |= 256  # Filling form fields
    if permissions.get("accessibility", False):
        flag |= 512  # Extracting text/graphics for accessibility
    if permissions.get("document_assembly", False):
        flag |= 1024  # Assembling the document
    if permissions.get("high_quality_printing", False):
        flag |= 2048  # High-quality printing
    
    return flag

def get_user_inputs():
    """Get configuration from user input"""
    print("=== Enhanced PDF Security Tool ===\n")
    
    # Get directory to scan
    while True:
        root_folder = input("Enter the directory path to scan for PDFs: ").strip()
        if os.path.exists(root_folder):
            break
        print("❌ Directory not found. Please try again.")
    
    # Get security settings
    permissions = get_security_level()
    
    # Get password options
    print("\n🔑 Password Configuration:")
    user_password = input("Enter user password (leave empty for no password): ").strip()
    
    while True:
        owner_password = input("Enter owner password (required): ").strip()
        if owner_password:
            break
        print("❌ Owner password is required.")
    
    # Additional options
    print("\n⚙️  Additional Options:")
    backup = input("Create backup copies? (y/n): ").lower().startswith('y')
    encrypt_metadata = input("Encrypt document metadata? (y/n): ").lower().startswith('y')
    
    return root_folder, user_password, owner_password, permissions, backup, encrypt_metadata

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

def secure_pdf(file_path, user_password, owner_password, permissions, encrypt_metadata, log_file_path, backup_dir=None):
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
        
        # Convert permissions to flag
        permissions_flag = permissions_to_flag(permissions)
        
        # Apply security settings
        writer.encrypt(
            user_password=user_password if user_password else "",
            owner_password=owner_password,
            permissions_flag=permissions_flag
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

def log_permission_summary(log_file_path, permissions):
    """Log the permission settings being applied"""
    log_message(log_file_path, "🔐 Permission Settings:")
    for perm, enabled in permissions.items():
        status = "✅ Allowed" if enabled else "❌ Denied"
        perm_name = perm.replace("_", " ").title()
        log_message(log_file_path, f"   {perm_name}: {status}")

def scan_and_secure_pdfs(root_folder, user_password, owner_password, permissions, encrypt_metadata, backup_enabled, log_file_path):
    """Main function to scan directory and secure all PDFs"""
    
    # Setup backup directory if enabled
    backup_dir = None
    if backup_enabled:
        backup_dir = os.path.join(root_folder, "pdf_backups")
    
    # Statistics
    total_files = 0
    secured_files = 0
    failed_files = 0
    
    log_message(log_file_path, f"🔍 Starting Enhanced PDF Security Scan")
    log_message(log_file_path, f"📁 Directory: {root_folder}")
    log_message(log_file_path, f"🔒 User password: {'Set' if user_password else 'None'}")
    log_message(log_file_path, f"🔐 Owner password: Set")
    log_message(log_file_path, f"🗃️  Encrypt metadata: {encrypt_metadata}")
    log_message(log_file_path, f"💾 Backup enabled: {backup_enabled}")
    
    # Log permission settings
    log_permission_summary(log_file_path, permissions)
    log_message(log_file_path, "=" * 60)
    
    # Walk through directory
    for dirpath, dirnames, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith('.pdf'):
                file_path = os.path.join(dirpath, filename)
                total_files += 1
                
                log_message(log_file_path, f"📄 Processing: {file_path}")
                
                if secure_pdf(file_path, user_password, owner_password, permissions, encrypt_metadata, log_file_path, backup_dir):
                    secured_files += 1
                else:
                    failed_files += 1
    
    # Summary
    log_message(log_file_path, "=" * 60)
    log_message(log_file_path, "📊 FINAL SUMMARY:")
    log_message(log_file_path, f"   Total PDFs found: {total_files}")
    log_message(log_file_path, f"   Successfully secured: {secured_files}")
    log_message(log_file_path, f"   Failed: {failed_files}")
    log_message(log_file_path, f"   Success rate: {(secured_files/total_files*100):.1f}%" if total_files > 0 else "   Success rate: N/A")
    
    return total_files, secured_files, failed_files

def main():
    """Main execution function"""
    try:
        # Get user inputs
        root_folder, user_password, owner_password, permissions, backup_enabled, encrypt_metadata = get_user_inputs()
        
        # Setup logging
        log_file_path = setup_logging(root_folder)
        
        # Display configuration summary
        print(f"\n📋 Configuration Summary:")
        print(f"   Directory: {root_folder}")
        print(f"   User password: {'Set' if user_password else 'None (no password to open)'}")
        print(f"   Owner password: Set")
        print(f"   Create backups: {backup_enabled}")
        print(f"   Encrypt metadata: {encrypt_metadata}")
        print(f"   Log file: {log_file_path}")
        
        print(f"\n🔐 Permission Summary:")
        for perm, enabled in permissions.items():
            status = "✅" if enabled else "❌"
            perm_name = perm.replace("_", " ").title()
            print(f"   {status} {perm_name}")
        
        confirm = input("\nProceed with securing PDFs? (y/n): ").lower()
        if not confirm.startswith('y'):
            print("Operation cancelled.")
            return
        
        # Execute the security process
        print("\n🚀 Starting Enhanced PDF Security Process...\n")
        total, secured, failed = scan_and_secure_pdfs(
            root_folder, user_password, owner_password, permissions, 
            encrypt_metadata, backup_enabled, log_file_path
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