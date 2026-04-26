import os
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict

def extract_launchbox_rom_paths(launchbox_data_dir, output_file="rom_paths.txt"):
    """
    Extract ROM directory paths from LaunchBox platform XML files.
    """
    platforms_dir = Path(launchbox_data_dir)
    
    if not platforms_dir.exists():
        print(f"Directory does not exist: {platforms_dir}")
        return
    
    # Find all XML files in the platforms directory
    xml_files = list(platforms_dir.glob("*.xml"))
    
    if not xml_files:
        print("No XML files found in the platforms directory.")
        return
    
    print(f"Found {len(xml_files)} platform XML files:")
    for xml_file in xml_files:
        print(f"  - {xml_file.name}")
    
    # Dictionary to track ROM paths by platform
    platform_rom_paths = defaultdict(set)
    all_rom_paths = set()
    
    for xml_file in xml_files:
        platform_name = xml_file.stem  # Filename without extension
        print(f"\nProcessing: {platform_name}")
        
        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()
            
            game_count = 0
            # Process main Game elements
            for game in root.findall(".//Game"):
                app_path_elem = game.find("ApplicationPath")
                if app_path_elem is not None and app_path_elem.text:
                    # Convert relative path to absolute and get directory
                    rom_path = app_path_elem.text.strip()
                    rom_dir = os.path.dirname(rom_path)
                    
                    if rom_dir:
                        # Remove leading ".." or ".\" and normalize
                        rom_dir = rom_dir.replace("..\\", "").replace("../", "").replace(".\\", "").replace("./", "")
                        
                        platform_rom_paths[platform_name].add(rom_dir)
                        all_rom_paths.add(rom_dir)
                        game_count += 1
            
            # Process AdditionalApplication elements
            additional_count = 0
            for additional in root.findall(".//AdditionalApplication"):
                app_path_elem = additional.find("ApplicationPath")
                if app_path_elem is not None and app_path_elem.text:
                    rom_path = app_path_elem.text.strip()
                    rom_dir = os.path.dirname(rom_path)
                    
                    if rom_dir:
                        rom_dir = rom_dir.replace("..\\", "").replace("../", "").replace(".\\", "").replace("./", "")
                        platform_rom_paths[platform_name].add(rom_dir)
                        all_rom_paths.add(rom_dir)
                        additional_count += 1
            
            print(f"  Found {game_count} games, {additional_count} additional applications")
            
        except ET.ParseError as e:
            print(f"Error parsing {xml_file.name}: {e}")
        except Exception as e:
            print(f"Error processing {xml_file.name}: {e}")
    
    # Write results to file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("# ROM Directory Paths extracted from LaunchBox\n")
        f.write("# Generated for Retrobat migration\n")
        f.write(f"# Source: {launchbox_data_dir}\n")
        f.write(f"# Processed {len(xml_files)} platform XML files\n\n")
        
        # Write paths organized by platform
        for platform, paths in sorted(platform_rom_paths.items()):
            if paths:
                f.write(f"# Platform: {platform} ({len(paths)} directories)\n")
                for path in sorted(paths):
                    f.write(f"{path}\n")
                f.write("\n")
        
        # Write summary
        f.write("# === ALL UNIQUE ROM DIRECTORIES ===\n")
        for path in sorted(all_rom_paths):
            f.write(f"{path}\n")
    
    print(f"\n✅ Extraction complete!")
    print(f"📁 Platforms processed: {len(platform_rom_paths)}")
    print(f"📂 Total unique ROM directories: {len(all_rom_paths)}")
    print(f"💾 Results saved to: {output_file}")
    
    # Show preview
    print(f"\n📋 Preview of ROM directories found:")
    for i, path in enumerate(sorted(all_rom_paths)[:10]):
        print(f"   {path}")
    if len(all_rom_paths) > 10:
        print(f"   ... and {len(all_rom_paths) - 10} more")

# Pre-configured for your LaunchBox installation
LAUNCHBOX_PLATFORMS_DIR = r"C:\Arcade\launchbox\data\platforms"

def main():
    print("LaunchBox ROM Path Extractor for Retrobat Migration")
    print("=" * 55)
    print(f"LaunchBox Directory: {LAUNCHBOX_PLATFORMS_DIR}")
    
    output_file = input("Enter output filename (default: rom_paths.txt): ").strip()
    if not output_file:
        output_file = "rom_paths.txt"
    
    extract_launchbox_rom_paths(LAUNCHBOX_PLATFORMS_DIR, output_file)

if __name__ == "__main__":
    main()