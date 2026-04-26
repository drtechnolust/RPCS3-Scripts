import os
import shutil
import re
import logging
from pathlib import Path
from datetime import datetime

def setup_logging(dest_dir):
    """Setup logging to track all operations"""
    log_file = os.path.join(dest_dir, f"organization_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()  # Also print to console
        ]
    )
    return log_file

def parse_version(name):
    """Extract version number from filename/folder name"""
    # Look for patterns like v1.23, V1.23, 1.23, etc.
    patterns = [
        r'v(\d+\.\d+)',
        r'V(\d+\.\d+)', 
        r'(\d+\.\d+)',
        r'v(\d+)',
        r'V(\d+)'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, name)
        if match:
            return float(match.group(1))
    return 0.0

def get_region(name):
    """Determine region from filename"""
    name_upper = name.upper()
    
    # USA/US indicators
    if any(indicator in name_upper for indicator in ['USA', '_US_', 'NTSC-U', 'NTSC_U']):
        return 'USA'
    
    # European indicators  
    elif any(indicator in name_upper for indicator in ['EUR', '_EU_', 'EUROPE', 'PAL']):
        return 'EUR'
    
    # Japanese indicators
    elif any(indicator in name_upper for indicator in ['JPN', '_JP_', 'JAPAN', 'NTSC-J']):
        return 'JPN'
    
    # Asian indicators
    elif any(indicator in name_upper for indicator in ['ASIA', 'ASIAN']):
        return 'ASIA'
    
    return 'Unknown'

def get_content_type(name):
    """Determine what type of content this is"""
    name_lower = name.lower()
    
    # DLC indicators
    if any(indicator in name_lower for indicator in ['dlc', 'addon', 'adddlc', 'add-on']):
        return 'DLC'
    
    # Update/Patch indicators
    elif any(indicator in name_lower for indicator in ['update', 'patch', 'fix']):
        return 'UPDATE'
    
    # File type indicators
    elif name_lower.endswith('.pkg'):
        return 'PKG'
    elif name_lower.endswith(('.rar', '.zip', '.7z')):
        return 'ARCHIVE'
    
    # Base game
    else:
        return 'GAME'

def extract_game_id(name):
    """Extract CUSA ID or use folder name as ID"""
    cusa_match = re.search(r'CUSA\d+', name.upper())
    if cusa_match:
        return cusa_match.group(0)
    
    # If no CUSA, use first part of name (before version/region indicators)
    clean_name = re.split(r'[_\-\.](v\d|V\d|\d+\.\d+|USA|EUR|JPN)', name)[0]
    return clean_name.strip()

def safe_move(src, dest_dir, dry_run=False):
    """Safely move file/folder with error handling"""
    try:
        if not os.path.exists(src):
            logging.warning(f"Source does not exist: {src}")
            return False
            
        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, os.path.basename(src))
        
        # Handle name conflicts
        counter = 1
        original_dest = dest_path
        while os.path.exists(dest_path):
            name, ext = os.path.splitext(original_dest)
            dest_path = f"{name}_({counter}){ext}"
            counter += 1
        
        if dry_run:
            logging.info(f"[DRY RUN] Would move: {src} -> {dest_path}")
            return True
        else:
            shutil.move(src, dest_path)
            logging.info(f"Moved: {src} -> {dest_path}")
            return True
            
    except Exception as e:
        logging.error(f"Failed to move {src}: {str(e)}")
        return False

def clean_empty_folders(path, dry_run=False):
    """Remove empty folders after organization"""
    removed_count = 0
    for root, dirs, files in os.walk(path, topdown=False):
        try:
            if not dirs and not files and root != path:
                if dry_run:
                    logging.info(f"[DRY RUN] Would remove empty folder: {root}")
                else:
                    os.rmdir(root)
                    logging.info(f"Removed empty folder: {root}")
                removed_count += 1
        except Exception as e:
            logging.warning(f"Could not remove folder {root}: {str(e)}")
    
    logging.info(f"{'Would remove' if dry_run else 'Removed'} {removed_count} empty folders")

def validate_paths(root_dir, dest_dir):
    """Validate input and output paths"""
    if not os.path.exists(root_dir):
        raise ValueError(f"Source directory does not exist: {root_dir}")
    
    if not os.path.isdir(root_dir):
        raise ValueError(f"Source path is not a directory: {root_dir}")
    
    # Create destination if it doesn't exist
    try:
        os.makedirs(dest_dir, exist_ok=True)
    except Exception as e:
        raise ValueError(f"Cannot create destination directory {dest_dir}: {str(e)}")
    
    # Check if dest is inside source (would cause infinite loop)
    if os.path.commonpath([root_dir, dest_dir]) == root_dir:
        raise ValueError("Destination directory cannot be inside source directory")

def show_preview(items_to_process):
    """Show what will be organized before processing"""
    print("\n" + "="*80)
    print("PREVIEW - Items to be organized:")
    print("="*80)
    
    categories = {}
    for item in items_to_process:
        category = item['category']
        if category not in categories:
            categories[category] = []
        categories[category].append(item['name'])
    
    for category, items in categories.items():
        print(f"\n{category}: {len(items)} items")
        for item in items[:5]:  # Show first 5 items
            print(f"  - {item}")
        if len(items) > 5:
            print(f"  ... and {len(items) - 5} more")
    
    print("\n" + "="*80)
    return input("Continue with organization? (y/n): ").lower().startswith('y')

def main():
    print("PlayStation Game Organizer v2.0")
    print("="*50)
    
    # Get user input
    ROOT_DIR = input("Enter the source root directory to scan: ").strip().strip('"')
    DEST_DIR = input("Enter the destination base directory: ").strip().strip('"')
    
    try:
        validate_paths(ROOT_DIR, DEST_DIR)
    except ValueError as e:
        print(f"Error: {e}")
        return
    
    # Setup logging
    log_file = setup_logging(DEST_DIR)
    logging.info(f"Starting organization from {ROOT_DIR} to {DEST_DIR}")
    
    # Ask for dry run
    dry_run = input("Run in preview mode first? (y/n): ").lower().startswith('y')
    
    # Define destination folders
    folders = {
        'DLC': os.path.join(DEST_DIR, "DLC"),
        'UPDATE': os.path.join(DEST_DIR, "Updates"), 
        'PKG': os.path.join(DEST_DIR, "PKG Files"),
        'ARCHIVE': os.path.join(DEST_DIR, "Archive Files"),
        'OLD_VERSION': os.path.join(DEST_DIR, "Old Versions"),
        'OTHER_REGIONS': os.path.join(DEST_DIR, "Other Regions"),
        'USA_GAMES': os.path.join(DEST_DIR, "USA Games"),
        'UNKNOWN': os.path.join(DEST_DIR, "Unknown Content")
    }
    
    # Track seen games for version management
    seen_games = {}
    items_to_process = []
    processed_count = 0
    
    print("Scanning directories...")
    
    # Scan all files and folders
    for root, dirs, files in os.walk(ROOT_DIR):
        # Limit depth to prevent excessive scanning
        depth = len(Path(root).relative_to(ROOT_DIR).parts)
        if depth > 5:
            continue
        
        # Process files
        for file in files:
            if file.lower().endswith(('.pkg', '.rar', '.zip', '.7z')):
                full_path = os.path.join(root, file)
                content_type = get_content_type(file)
                region = get_region(file)
                version = parse_version(file)
                
                items_to_process.append({
                    'path': full_path,
                    'name': file,
                    'type': 'file',
                    'content_type': content_type,
                    'region': region,
                    'version': version,
                    'category': content_type
                })
        
        # Process directories
        for directory in dirs:
            full_path = os.path.join(root, directory)
            content_type = get_content_type(directory)
            region = get_region(directory)
            version = parse_version(directory)
            game_id = extract_game_id(directory)
            
            items_to_process.append({
                'path': full_path,
                'name': directory,
                'type': 'folder',
                'content_type': content_type,
                'region': region,
                'version': version,
                'game_id': game_id,
                'category': content_type
            })
    
    print(f"Found {len(items_to_process)} items to organize")
    
    # Show preview if requested
    if dry_run and items_to_process:
        if not show_preview(items_to_process):
            print("Organization cancelled by user")
            return
        
        # Ask if they want to do actual run now
        actual_run = input("Run actual organization now? (y/n): ").lower().startswith('y')
        if not actual_run:
            print("Stopping after preview")
            return
        dry_run = False
    
    # Process items
    print(f"{'Previewing' if dry_run else 'Processing'} items...")
    
    for item in items_to_process:
        processed_count += 1
        if processed_count % 10 == 0:
            print(f"Progress: {processed_count}/{len(items_to_process)}")
        
        path = item['path']
        content_type = item['content_type']
        region = item['region']
        version = item['version']
        
        # Determine destination based on content type and region
        if content_type == 'DLC':
            dest = folders['DLC']
            
        elif content_type == 'UPDATE':
            dest = folders['UPDATE']
            
        elif content_type in ['PKG']:
            dest = folders['PKG']
            
        elif content_type == 'ARCHIVE':
            dest = folders['ARCHIVE']
            
        elif region != 'USA' and region != 'Unknown':
            dest = folders['OTHER_REGIONS']
            
        elif content_type == 'GAME' and region in ['USA', 'Unknown']:
            # Handle version management for games
            if item['type'] == 'folder':
                game_id = item['game_id']
                
                if game_id in seen_games:
                    if version > seen_games[game_id]['version']:
                        # Move old version to old versions folder
                        old_path = seen_games[game_id]['path']
                        safe_move(old_path, folders['OLD_VERSION'], dry_run)
                        
                        # Update tracking and move new version to main folder
                        seen_games[game_id] = {'version': version, 'path': path}
                        dest = folders['USA_GAMES']
                    else:
                        # This is an older version
                        dest = folders['OLD_VERSION']
                else:
                    # First time seeing this game
                    seen_games[game_id] = {'version': version, 'path': path}
                    dest = folders['USA_GAMES']
            else:
                dest = folders['USA_GAMES']
        else:
            dest = folders['UNKNOWN']
        
        # Move the item
        safe_move(path, dest, dry_run)
    
    # Clean up empty folders
    if not dry_run:
        print("Cleaning up empty folders...")
        clean_empty_folders(ROOT_DIR, dry_run)
    
    # Summary
    print(f"\n{'Preview' if dry_run else 'Organization'} complete!")
    print(f"Processed {processed_count} items")
    print(f"Log file: {log_file}")
    
    if not dry_run:
        print("\nOrganized folders:")
        for category, folder_path in folders.items():
            if os.path.exists(folder_path):
                item_count = len(os.listdir(folder_path))
                print(f"  {category}: {item_count} items in {folder_path}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        logging.error(f"Unexpected error: {str(e)}")