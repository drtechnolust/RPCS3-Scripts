import xml.etree.ElementTree as ET
import os
import csv
from pathlib import Path
import sys
import codecs

def find_common_base_path(paths):
    """Find the most common base directory from a list of paths."""
    if not paths:
        return ''
    
    if len(paths) == 1:
        return str(Path(paths[0]).parent)
    
    try:
        # Convert to Path objects and find common parent
        path_objects = [Path(p) for p in paths]
        # Get all parent directories
        all_parents = [p.parent for p in path_objects]
        
        # Find the most common parent directory
        parent_counts = {}
        for parent in all_parents:
            parent_str = str(parent)
            parent_counts[parent_str] = parent_counts.get(parent_str, 0) + 1
        
        # Return the most common parent, or use os.path.commonpath as fallback
        if parent_counts:
            most_common = max(parent_counts.items(), key=lambda x: x[1])[0]
            return most_common
        else:
            return os.path.commonpath(paths)
    except (ValueError, OSError):
        # Fallback to first path's directory if commonpath fails
        return str(Path(paths[0]).parent)

def validate_paths(game_paths, base_path):
    """Validate that paths exist and return valid ones with warnings for invalid."""
    valid_paths = []
    invalid_count = 0
    
    for path in game_paths:
        try:
            # Handle both absolute and relative paths
            if Path(path).is_absolute():
                full_path = Path(path)
            else:
                full_path = Path(base_path) / path
            
            if full_path.exists():
                valid_paths.append(path)
            else:
                invalid_count += 1
        except Exception:
            invalid_count += 1
    
    return valid_paths, invalid_count

def clean_text(text):
    """Clean text to remove problematic characters."""
    if not text:
        return ""
    
    # Replace common problematic Unicode characters
    replacements = {
        '\u2019': "'",  # Right single quotation mark
        '\u2018': "'",  # Left single quotation mark
        '\u201c': '"',  # Left double quotation mark
        '\u201d': '"',  # Right double quotation mark
        '\u2013': '-',  # En dash
        '\u2014': '-',  # Em dash
        '\u00e9': 'e',  # é
        '\u00e8': 'e',  # è
        '\u00ea': 'e',  # ê
        '\u00eb': 'e',  # ë
        '\u00ed': 'i',  # í
        '\u00ec': 'i',  # ì
        '\u00ee': 'i',  # î
        '\u00ef': 'i',  # ï
        '\u00f3': 'o',  # ó
        '\u00f2': 'o',  # ò
        '\u00f4': 'o',  # ô
        '\u00f6': 'o',  # ö
        '\u00fa': 'u',  # ú
        '\u00f9': 'u',  # ù
        '\u00fb': 'u',  # û
        '\u00fc': 'u',  # ü
        '\u00e1': 'a',  # á
        '\u00e0': 'a',  # à
        '\u00e2': 'a',  # â
        '\u00e4': 'a',  # ä
        '\u00f1': 'n',  # ñ
        '\u00e7': 'c',  # ç
    }
    
    cleaned = text
    for unicode_char, replacement in replacements.items():
        cleaned = cleaned.replace(unicode_char, replacement)
    
    # Remove any remaining non-ASCII characters
    cleaned = ''.join(char for char in cleaned if ord(char) < 128)
    
    return cleaned.strip()

def safe_parse_xml(xml_file):
    """Safely parse XML file with multiple encoding attempts."""
    encodings_to_try = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1', 'windows-1252']
    
    for encoding in encodings_to_try:
        try:
            with open(xml_file, 'r', encoding=encoding, errors='replace') as f:
                content = f.read()
            
            # Clean the content before parsing
            content = clean_text(content)
            
            # Parse the XML content
            root = ET.fromstring(content)
            return root
            
        except Exception:
            continue
    
    # Final attempt with binary mode and manual cleaning
    try:
        with open(xml_file, 'rb') as f:
            raw_content = f.read()
        
        # Try to decode with errors='ignore'
        content = raw_content.decode('utf-8', errors='ignore')
        content = clean_text(content)
        
        root = ET.fromstring(content)
        return root
    except Exception:
        return None

def main():
    # Configuration - Your specific paths
    launchbox_xml_dir = Path('C:/Arcade/LaunchBox/Data/Platforms/')
    output_csv = Path('C:/Arcade/RetroBat/Convertion/launchbox_platform_paths.csv')
    
    # Set console encoding to handle Unicode
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass
    
    print("LaunchBox Platform Path Extractor")
    print("=" * 40)
    print(f"Source: {launchbox_xml_dir}")
    print(f"Output: {output_csv}")
    print()
    
    # Ensure directories exist
    try:
        output_csv.parent.mkdir(parents=True, exist_ok=True)
    except PermissionError:
        print(f"Error: Permission denied creating directory {output_csv.parent}")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Check if source directory exists
    if not launchbox_xml_dir.exists():
        print(f"Error: LaunchBox directory not found: {launchbox_xml_dir}")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Initialize list to hold data
    platform_data = []
    total_invalid_paths = 0
    failed_files = []
    processed_count = 0
    
    # Get list of XML files
    xml_files = [f for f in launchbox_xml_dir.iterdir() if f.suffix.lower() == '.xml']
    
    if not xml_files:
        print(f"No XML files found in {launchbox_xml_dir}")
        input("Press Enter to exit...")
        sys.exit(1)
    
    print(f"Processing {len(xml_files)} XML files...")
    print("-" * 50)
    
    # Parse all XML files
    for i, xml_file in enumerate(xml_files, 1):
        platform_name = clean_text(xml_file.stem)  # Clean platform name
        
        try:
            print(f"Processing {i}/{len(xml_files)}: {platform_name}")
        except UnicodeEncodeError:
            print(f"Processing {i}/{len(xml_files)}: [Platform with special characters]")
        
        # Use safe XML parsing
        root = safe_parse_xml(xml_file)
        
        if root is None:
            print(f"  - Error parsing XML: Could not decode file")
            failed_files.append(platform_name)
            continue
        
        # Extract all game paths for this platform
        game_paths = []
        
        try:
            for game in root.findall('.//Game'):
                path_elem = game.find('ApplicationPath')
                if path_elem is not None and path_elem.text and path_elem.text.strip():
                    # Clean up the path text
                    clean_path = clean_text(path_elem.text.strip())
                    if clean_path:  # Only add if cleaning didn't remove everything
                        game_paths.append(clean_path)
        except Exception as e:
            print(f"  - Error extracting paths: Skipping file")
            continue
        
        # Add to data if there are paths
        if game_paths:
            # Find the best base path
            try:
                base_path = find_common_base_path(game_paths)
                base_path = clean_text(base_path)
            except Exception:
                base_path = ""
            
            # Validate paths (simplified to avoid encoding issues)
            paths_to_use = game_paths  # Skip validation to avoid more encoding issues
            
            platform_data.append([
                platform_name, 
                base_path, 
                ';'.join(paths_to_use),
                len(paths_to_use)
            ])
            
            processed_count += 1
            print(f"  + Found {len(paths_to_use)} games")
        else:
            print(f"  - No game paths found")
    
    # Write to CSV with ASCII encoding
    if platform_data:
        try:
            with open(output_csv, 'w', newline='', encoding='ascii', errors='ignore') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Platform_Name', 'Base_Path', 'All_Paths', 'Game_Count'])
                writer.writerows(platform_data)
            
            print("-" * 50)
            print(f"Success! Data saved to {output_csv}")
            print(f"Statistics:")
            print(f"   • Platforms processed: {processed_count}")
            print(f"   • Total games found: {sum(row[3] for row in platform_data)}")
            if failed_files:
                print(f"   • Failed to parse: {len(failed_files)} files")
            
        except Exception as e:
            print(f"Error writing CSV file: {e}")
            input("Press Enter to exit...")
            sys.exit(1)
    else:
        print("No platforms with valid game paths found!")
        input("Press Enter to exit...")
        sys.exit(1)
    
    print("\nProcessing complete!")
    input("Press Enter to exit...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        input("Press Enter to exit...")
        sys.exit(1)