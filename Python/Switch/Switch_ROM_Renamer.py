import os
import re
import csv
from pathlib import Path

def extract_region_and_version(filename):
    """
    Extract region and version info from filename
    Returns: (region, version) tuple
    """
    # Remove .mp extension
    name_without_ext = filename.replace('.mp', '')
    
    # Pattern to match region: __[REGION]_ or __[REGION]_v[VERSION]_
    region_pattern = r'__([A-Z]{2,3})_'
    region_match = re.search(region_pattern, name_without_ext)
    region = region_match.group(1) if region_match else None
    
    # Pattern to match version: _v[NUMBER]_ or _v[NUMBER]
    version_pattern = r'_v(\d+(?:\.\d+)?)_?'
    version_match = re.search(version_pattern, name_without_ext)
    version = version_match.group(1) if version_match else None
    
    return region, version

def clean_filename(filename):
    """
    Extract and clean the game name from the filename, preserving region info
    """
    # Remove the .mp extension
    name_without_ext = filename.replace('.mp', '')
    
    # Extract region and version first
    region, version = extract_region_and_version(filename)
    
    # Pattern to match: everything before _[HEX_CODE]
    # Hex codes appear to be like _010044B1E8000 or similar
    pattern = r'^(.+?)_[0-9A-Fa-f]{12,}.*$'
    match = re.match(pattern, name_without_ext)
    
    if match:
        game_name = match.group(1)
    else:
        # Fallback: take everything before the first underscore followed by numbers/hex
        parts = name_without_ext.split('_')
        if len(parts) > 1:
            game_name = parts[0]
        else:
            game_name = name_without_ext
    
    # Clean up the game name
    cleaned_name = clean_game_name(game_name)
    
    # Add region and version info to the cleaned name
    final_name = build_final_name(cleaned_name, region, version)
    
    return final_name, region, version

def clean_game_name(name):
    """
    Clean up the game name formatting
    """
    # Replace common separators with spaces
    name = re.sub(r'[_\-:]+', ' ', name)
    
    # Handle special cases for numbers at start
    if re.match(r'^\d+\s*', name):
        # If it starts with numbers, keep them but ensure proper spacing
        name = re.sub(r'^(\d+)\s*(.*)$', r'\1 \2', name).strip()
    
    # Handle version info that might have leaked through
    name = re.sub(r'\s+v\d+.*$', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s+(base|app|update).*$', '', name, flags=re.IGNORECASE)
    
    # Clean up spacing
    name = ' '.join(name.split())
    
    # Proper title case, but preserve all-caps words like "STAR WARS"
    words = name.split()
    cleaned_words = []
    
    for word in words:
        if word.isupper() and len(word) > 2:
            # Keep words that are all uppercase (like acronyms)
            cleaned_words.append(word)
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
        final_name += f" [{region}]"
    
    # Add version if available
    if version:
        final_name += f" [v{version}]"
    
    return final_name

def process_files(directory_path, output_csv='game_name_mapping.csv', rename_files=False):
    """
    Process all .mp files in the directory and create a mapping
    """
    directory = Path(directory_path)
    mappings = []
    
    # Get all .mp files
    mp_files = list(directory.glob('*.mp'))
    
    for file_path in mp_files:
        original_name = file_path.name
        cleaned_name, region, version = clean_filename(original_name)
        
        # Create new filename
        new_filename = f"{cleaned_name}.mp"
        
        mappings.append({
            'original_filename': original_name,
            'cleaned_game_name': cleaned_name,
            'new_filename': new_filename,
            'region': region or 'Unknown',
            'version': version or 'Unknown'
        })
        
        print(f"Original: {original_name}")
        print(f"Cleaned:  {new_filename}")
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
        fieldnames = ['original_filename', 'cleaned_game_name', 'new_filename', 'region', 'version']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()
        for mapping in mappings:
            writer.writerow(mapping)
    
    print(f"\nMapping saved to: {output_csv}")
    print(f"Processed {len(mappings)} files")
    
    # Print summary by region
    region_counts = {}
    for mapping in mappings:
        region = mapping['region']
        region_counts[region] = region_counts.get(region, 0) + 1
    
    print("\nRegion distribution:")
    for region, count in sorted(region_counts.items()):
        print(f"  {region}: {count} files")
    
    return mappings

# Additional function to analyze patterns before processing
def analyze_patterns(directory_path):
    """
    Analyze the filename patterns to understand the data better
    """
    directory = Path(directory_path)
    mp_files = list(directory.glob('*.mp'))
    
    regions = set()
    versions = set()
    
    print("Analyzing filename patterns...")
    print("=" * 60)
    
    for file_path in mp_files[:20]:  # Analyze first 20 files
        filename = file_path.name
        region, version = extract_region_and_version(filename)
        
        if region:
            regions.add(region)
        if version:
            versions.add(version)
    
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