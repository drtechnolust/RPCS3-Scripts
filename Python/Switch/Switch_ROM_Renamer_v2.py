import os
import re
import csv
from pathlib import Path

SUPPORTED_EXTENSIONS = ['.xci', '.nsp']

def extract_region_and_version(filename):
    """
    Extract region and version info from filename
    Returns: (region, version) tuple
    """
    # Get the base name without extension
    name_without_ext = Path(filename).stem
    
    # Pattern to match region: __[REGION]_ or [REGION] or (REGION)
    region_patterns = [
        r'__([A-Z]{2,3})_',  # __US_ format
        r'\[([A-Z]{2,3})\]',  # [US] format
        r'\(([A-Z]{2,3})\)'   # (US) format
    ]
    
    region = None
    for pattern in region_patterns:
        region_match = re.search(pattern, name_without_ext)
        if region_match:
            region = region_match.group(1)
            break
    
    # Pattern to match version: _v[NUMBER]_ or _v[NUMBER] or [v[NUMBER]] or (v[NUMBER])
    version_patterns = [
        r'_v(\d+(?:\.\d+)*(?:\.\d+)?)_?',  # _v1.0.0_ or _v1_
        r'\[v(\d+(?:\.\d+)*(?:\.\d+)?)\]',  # [v1.0.0]
        r'\(v(\d+(?:\.\d+)*(?:\.\d+)?)\)'   # (v1.0.0)
    ]
    
    version = None
    for pattern in version_patterns:
        version_match = re.search(pattern, name_without_ext)
        if version_match:
            version = version_match.group(1)
            break
    
    return region, version

def clean_filename(filename):
    """
    Extract and clean the game name from the filename, preserving region info
    """
    # Get the file extension and base name
    file_path = Path(filename)
    extension = file_path.suffix.lower()
    name_without_ext = file_path.stem
    
    # Extract region and version first
    region, version = extract_region_and_version(filename)
    
    # Pattern to match: everything before _[HEX_CODE] or other unwanted suffixes
    # Hex codes appear to be like _010044B1E8000 or similar
    patterns_to_remove = [
        r'^(.+?)_[0-9A-Fa-f]{12,}.*$',  # Hex codes
        r'^(.+?)__[A-Z]{2,3}_.*$',      # Region markers
        r'^(.+?)\s*\[.*\].*$',          # Anything in brackets
        r'^(.+?)\s*\(.*\).*$'           # Anything in parentheses
    ]
    
    game_name = name_without_ext
    for pattern in patterns_to_remove:
        match = re.match(pattern, game_name)
        if match:
            game_name = match.group(1)
            break
    
    # If no pattern matched, try splitting on common separators
    if game_name == name_without_ext:
        # Split on underscores and take the first substantial part
        parts = game_name.split('_')
        if len(parts) > 1:
            # Take parts until we hit something that looks like metadata
            clean_parts = []
            for part in parts:
                if re.match(r'^[0-9A-Fa-f]{8,}$', part) or part.lower() in ['nsw2u', 'patched', 'base', 'update']:
                    break
                clean_parts.append(part)
            if clean_parts:
                game_name = '_'.join(clean_parts)
    
    # Clean up the game name
    cleaned_name = clean_game_name(game_name)
    
    # Add region and version info to the cleaned name
    final_name = build_final_name(cleaned_name, region, version)
    
    return final_name, region, version, extension

def clean_game_name(name):
    """
    Clean up the game name formatting
    """
    # Replace common separators with spaces
    name = re.sub(r'[_\-:]+', ' ', name)
    
    # Remove common unwanted suffixes that might have leaked through
    unwanted_suffixes = [
        r'\s+(nsw2u|patched|base|update|dlc).*$',
        r'\s+v\d+.*$',
        r'\s+\d{4}-\d{2}-\d{2}.*$'  # Dates
    ]
    
    for suffix_pattern in unwanted_suffixes:
        name = re.sub(suffix_pattern, '', name, flags=re.IGNORECASE)
    
    # Handle special cases for numbers at start
    if re.match(r'^\d+\s*', name):
        # If it starts with numbers, keep them but ensure proper spacing
        name = re.sub(r'^(\d+)\s*(.*)$', r'\1 \2', name).strip()
    
    # Clean up spacing
    name = ' '.join(name.split())
    
    # Proper title case, but preserve all-caps words and handle special cases
    words = name.split()
    cleaned_words = []
    
    for word in words:
        if word.isupper() and len(word) > 2:
            # Keep words that are all uppercase (like acronyms)
            cleaned_words.append(word)
        elif word.lower() in ['and', 'or', 'the', 'of', 'in', 'on', 'at', 'to', 'for', 'with']:
            # Keep common words lowercase unless they're the first word
            if len(cleaned_words) == 0:
                cleaned_words.append(word.title())
            else:
                cleaned_words.append(word.lower())
        elif word.islower() or (word[0].islower() and len(word) > 1):
            # Title case for lowercase words
            cleaned_words.append(word.title())
        else:
            # Keep words that already have mixed case
            cleaned_words.append(word)
    
    return ' '.join(cleaned_words)

def build_final_name(game_name, region, version):
    """
    Build the final filename with region and version info
    """
    final_name = game_name
    
    # Add region if available
    if region:
        final_name += f" ({region})"
    
    # Add version if available
    if version:
        final_name += f" [v{version}]"
    
    return final_name

def process_files(directory_path, output_csv='game_name_mapping.csv', rename_files=False):
    """
    Process all supported ROM files in the directory and create a mapping
    """
    directory = Path(directory_path)
    mappings = []
    
    # Get all supported files
    rom_files = []
    for extension in SUPPORTED_EXTENSIONS:
        rom_files.extend(directory.glob(f'*{extension}'))
    
    print(f"Found {len(rom_files)} ROM files to process")
    
    for file_path in sorted(rom_files):
        original_name = file_path.name
        cleaned_name, region, version, extension = clean_filename(original_name)
        
        # Create new filename
        new_filename = f"{cleaned_name}{extension}"
        
        mappings.append({
            'original_filename': original_name,
            'cleaned_game_name': cleaned_name,
            'new_filename': new_filename,
            'region': region or 'Unknown',
            'version': version or 'Unknown',
            'file_type': extension.upper()[1:]  # Remove the dot and uppercase
        })
        
        print(f"Original: {original_name}")
        print(f"Cleaned:  {new_filename}")
        print(f"Type:     {extension.upper()[1:]}")
        print(f"Region:   {region or 'Unknown'}")
        print(f"Version:  {version or 'Unknown'}")
        print("-" * 80)
        
        # Optionally rename the files
        if rename_files:
            new_path = file_path.parent / new_filename
            if not new_path.exists():
                file_path.rename(new_path)
                print(f"Renamed: {original_name} -> {new_filename}")
            else:
                print(f"Warning: {new_filename} already exists, skipping rename")
    
    # Write mapping to CSV
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['original_filename', 'cleaned_game_name', 'new_filename', 'region', 'version', 'file_type']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for mapping in mappings:
            writer.writerow(mapping)
    
    print(f"\nMapping saved to: {output_csv}")
    print(f"Processed {len(mappings)} files")
    
    # Print summary statistics
    print_summary_stats(mappings)
    
    return mappings

def print_summary_stats(mappings):
    """
    Print summary statistics about the processed files
    """
    # Count by file type
    type_counts = {}
    region_counts = {}
    version_counts = {}
    
    for mapping in mappings:
        # File type counts
        file_type = mapping['file_type']
        type_counts[file_type] = type_counts.get(file_type, 0) + 1
        
        # Region counts
        region = mapping['region']
        region_counts[region] = region_counts.get(region, 0) + 1
        
        # Version counts
        version = mapping['version']
        if version != 'Unknown':
            version_counts[version] = version_counts.get(version, 0) + 1
    
    print("\nFile type distribution:")
    for file_type, count in sorted(type_counts.items()):
        print(f"  {file_type}: {count} files")
    
    print("\nRegion distribution:")
    for region, count in sorted(region_counts.items()):
        print(f"  {region}: {count} files")
    
    if version_counts:
        print("\nVersion distribution:")
        for version, count in sorted(version_counts.items()):
            print(f"  v{version}: {count} files")

def analyze_patterns(directory_path):
    """
    Analyze the filename patterns to understand the data better
    """
    directory = Path(directory_path)
    
    # Get all supported files
    rom_files = []
    for extension in SUPPORTED_EXTENSIONS:
        rom_files.extend(directory.glob(f'*{extension}'))
    
    regions = set()
    versions = set()
    extensions = set()
    
    print("Analyzing filename patterns...")
    print("=" * 60)
    
    sample_size = min(20, len(rom_files))
    print(f"Analyzing first {sample_size} files out of {len(rom_files)} total...")
    
    for file_path in sorted(rom_files)[:sample_size]:
        filename = file_path.name
        region, version = extract_region_and_version(filename)
        extension = file_path.suffix.lower()
        
        extensions.add(extension)
        if region:
            regions.add(region)
        if version:
            versions.add(version)
        
        print(f"  {filename[:60]}{'...' if len(filename) > 60 else ''}")
    
    print("=" * 60)
    print(f"Found file types: {sorted(extensions)}")
    print(f"Found regions: {sorted(regions)}")
    print(f"Found versions: {sorted(versions)}")
    print("=" * 60)

# Example usage
if __name__ == "__main__":
    # Set your directory path here
    directory_path = "."  # Current directory, change as needed
    
    # First, analyze the patterns
    analyze_patterns(directory_path)
    
    print("\nPress Enter to continue with processing, or Ctrl+C to exit...")
    input()
    
    # Process files (set rename_files=True to actually rename them)
    mappings = process_files(directory_path, rename_files=False)
    
    # Preview some results
    print("\nSample mappings:")
    for i, mapping in enumerate(mappings[:10]):  # Show first 10
        print(f"{mapping['original_filename']} -> {mapping['new_filename']}")