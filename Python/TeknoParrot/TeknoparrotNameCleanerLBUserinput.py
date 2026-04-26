import xml.etree.ElementTree as ET
import os
import shutil
from pathlib import Path

def get_user_paths():
    """Get directory paths from user input with validation"""
    
    print("🎮 TeknoParrot LaunchBox Title Cleaner v3.0")
    print("=" * 60)
    print("Please provide the following directories:\n")
    
    # Get LaunchBox Platforms folder
    while True:
        print("📁 LaunchBox Platforms Folder:")
        print("   Example: C:\\LaunchBox\\Data\\Platforms")
        print("   (This folder contains your platform XML files)")
        
        launchbox_platforms = input("\nEnter LaunchBox Platforms path: ").strip().strip('"')
        
        if not launchbox_platforms:
            print("❌ Please enter a path")
            continue
            
        platforms_path = Path(launchbox_platforms)
        if not platforms_path.exists():
            print(f"❌ Directory not found: {launchbox_platforms}")
            retry = input("Try again? (y/n): ").strip().lower()
            if retry != 'y':
                return None, None, None
            continue
            
        # Look for TeknoParrot XML file
        teknoparrot_xml = None
        xml_files = list(platforms_path.glob("*.xml"))
        
        if xml_files:
            print(f"\n📄 Found {len(xml_files)} XML files:")
            for i, xml_file in enumerate(xml_files, 1):
                print(f"   {i}. {xml_file.name}")
            
            # Try to find TeknoParrot.xml automatically
            for xml_file in xml_files:
                if 'teknoparrot' in xml_file.name.lower():
                    teknoparrot_xml = xml_file
                    print(f"\n✅ Auto-detected: {xml_file.name}")
                    break
            
            if not teknoparrot_xml:
                print(f"\nWhich file is your TeknoParrot platform?")
                try:
                    choice = int(input(f"Choose (1-{len(xml_files)}): ")) - 1
                    if 0 <= choice < len(xml_files):
                        teknoparrot_xml = xml_files[choice]
                    else:
                        print("❌ Invalid choice")
                        continue
                except ValueError:
                    print("❌ Please enter a number")
                    continue
        else:
            print("❌ No XML files found in this directory")
            continue
        
        break
    
    # Get TeknoParrot UserProfiles folder
    while True:
        print(f"\n📁 TeknoParrot UserProfiles Folder:")
        print("   Example: C:\\TeknoParrot\\UserProfiles")
        print("   (This folder contains your game profile XML files)")
        
        teknoparrot_profiles = input("\nEnter TeknoParrot UserProfiles path: ").strip().strip('"')
        
        if not teknoparrot_profiles:
            print("❌ Please enter a path")
            continue
            
        profiles_path = Path(teknoparrot_profiles)
        if not profiles_path.exists():
            print(f"❌ Directory not found: {teknoparrot_profiles}")
            retry = input("Try again? (y/n): ").strip().lower()
            if retry != 'y':
                return None, None, None
            continue
            
        # Check for XML files
        xml_files = list(profiles_path.glob("*.xml"))
        if xml_files:
            print(f"✅ Found {len(xml_files)} XML files in UserProfiles")
            break
        else:
            print("⚠️  No XML files found in this directory")
            retry = input("Continue anyway? (y/n): ").strip().lower()
            if retry == 'y':
                break
    
    return str(teknoparrot_xml), teknoparrot_profiles, str(platforms_path)

def save_paths_config(launchbox_xml, teknoparrot_profiles, platforms_folder):
    """Save paths to a config file for future use"""
    try:
        config_file = Path("teknoparrot_cleaner_config.txt")
        with open(config_file, 'w') as f:
            f.write(f"launchbox_xml={launchbox_xml}\n")
            f.write(f"teknoparrot_profiles={teknoparrot_profiles}\n")
            f.write(f"platforms_folder={platforms_folder}\n")
        print(f"💾 Paths saved to: {config_file}")
        return True
    except Exception as e:
        print(f"⚠️  Could not save config: {e}")
        return False

def load_paths_config():
    """Load paths from config file"""
    try:
        config_file = Path("teknoparrot_cleaner_config.txt")
        if not config_file.exists():
            return None, None, None
            
        config = {}
        with open(config_file, 'r') as f:
            for line in f:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    config[key] = value
        
        return config.get('launchbox_xml'), config.get('teknoparrot_profiles'), config.get('platforms_folder')
    except Exception as e:
        print(f"⚠️  Could not load config: {e}")
        return None, None, None

def clean_teknoparrot_titles(launchbox_xml_path, teknoparrot_profiles_path):
    """
    Clean up TeknoParrot game titles in LaunchBox XML file using official game names
    from TeknoParrot UserProfiles XML files.
    """
    
    # Validate paths first
    launchbox_path = Path(launchbox_xml_path)
    profiles_path = Path(teknoparrot_profiles_path)
    
    if not launchbox_path.exists():
        print(f"❌ LaunchBox XML file not found: {launchbox_xml_path}")
        return
        
    if not profiles_path.exists():
        print(f"❌ TeknoParrot UserProfiles folder not found: {teknoparrot_profiles_path}")
        return
    
    # Load LaunchBox XML
    try:
        tree = ET.parse(launchbox_xml_path)
        root = tree.getroot()
        print(f"✅ Loaded LaunchBox XML: {Path(launchbox_xml_path).name}")
    except ET.ParseError as e:
        print(f"❌ Error parsing LaunchBox XML: {e}")
        return
    except Exception as e:
        print(f"❌ Error loading LaunchBox XML: {e}")
        return
    
    # Debug: Show LaunchBox XML structure
    game_elements = root.findall('.//Game')
    print(f"📋 Found {len(game_elements)} games in LaunchBox XML")
    
    # Create mapping of TeknoParrot games
    tp_game_names = {}
    xml_files_found = list(profiles_path.glob("*.xml"))
    print(f"📁 Found {len(xml_files_found)} XML files in UserProfiles folder")
    
    if len(xml_files_found) == 0:
        print("❌ No XML files found in UserProfiles folder!")
        return
    
    print(f"\n🔍 Scanning TeknoParrot profiles...")
    
    # Scan TeknoParrot UserProfiles folder
    for xml_file in xml_files_found:
        try:
            # Parse XML with namespace handling
            tp_tree = ET.parse(xml_file)
            tp_root = tp_tree.getroot()
            
            # Try multiple ways to find the game name
            game_name = None
            
            # Method 1: Look for GameNameInternal
            for elem in tp_root.iter():
                if elem.tag.endswith('GameNameInternal') and elem.text:
                    game_name = elem.text.strip()
                    break
            
            # Method 2: Look for GameName (alternative tag)
            if not game_name:
                for elem in tp_root.iter():
                    if elem.tag.endswith('GameName') and elem.text:
                        game_name = elem.text.strip()
                        break
            
            # Method 3: Look for ProfileName as fallback
            if not game_name:
                for elem in tp_root.iter():
                    if elem.tag.endswith('ProfileName') and elem.text:
                        game_name = elem.text.strip()
                        break
            
            if game_name:
                key = xml_file.stem  # filename without extension
                tp_game_names[key] = game_name
                print(f"✅ {key:<20} -> {game_name}")
            else:
                print(f"⚠️  {xml_file.stem:<20} -> [No game name found]")
                
        except ET.ParseError as e:
            print(f"❌ XML parse error in {xml_file.name}: {e}")
        except Exception as e:
            print(f"❌ Error reading {xml_file.name}: {e}")
    
    print(f"\n📊 Successfully mapped {len(tp_game_names)} TeknoParrot games")
    
    if len(tp_game_names) == 0:
        print("❌ No game names found in TeknoParrot XML files!")
        return
    
    # Update LaunchBox XML
    games_updated = 0
    games_not_found = []
    
    print(f"\n🔍 Checking LaunchBox games for updates...")
    
    for game in game_elements:
        title_elem = game.find('Title')
        if title_elem is not None and title_elem.text:
            current_title = title_elem.text.strip()
            
            # Try exact match first
            if current_title in tp_game_names:
                new_title = tp_game_names[current_title]
                if new_title != current_title:
                    title_elem.text = new_title
                    print(f"✅ '{current_title}' -> '{new_title}'")
                    games_updated += 1
                else:
                    print(f"⚪ '{current_title}' already correct")
            else:
                # Try case-insensitive match
                found_match = False
                for tp_key, tp_name in tp_game_names.items():
                    if current_title.lower() == tp_key.lower():
                        title_elem.text = tp_name
                        print(f"✅ '{current_title}' -> '{tp_name}' (case-insensitive)")
                        games_updated += 1
                        found_match = True
                        break
                
                if not found_match:
                    games_not_found.append(current_title)
    
    # Show unmatched games
    if games_not_found:
        print(f"\n⚠️  Could not find TeknoParrot matches for {len(games_not_found)} games:")
        for game in sorted(games_not_found):
            print(f"   📝 {game}")
        print(f"\n💡 Tip: These might be custom titles or games not in your UserProfiles")
    
    # Save updated XML if changes were made
    if games_updated > 0:
        try:
            # Create backup with timestamp
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_path = f"{launchbox_xml_path}.backup_{timestamp}"
            shutil.copy2(launchbox_xml_path, backup_path)
            print(f"\n💾 Backup created: {Path(backup_path).name}")
            
            # Save updated XML with proper formatting
            tree.write(launchbox_xml_path, encoding='utf-8', xml_declaration=True)
            print(f"✅ LaunchBox XML updated successfully!")
            print(f"📈 Total games updated: {games_updated}")
            
        except PermissionError:
            print(f"❌ Permission denied. Make sure LaunchBox is closed and run as administrator.")
        except Exception as e:
            print(f"❌ Error saving file: {e}")
            
    else:
        print(f"\n✨ No updates needed - all titles are already correct!")

def show_available_games(teknoparrot_profiles_path):
    """Show all available TeknoParrot games and their proper names"""
    profiles_path = Path(teknoparrot_profiles_path)
    
    if not profiles_path.exists():
        print(f"❌ TeknoParrot UserProfiles path not found: {teknoparrot_profiles_path}")
        return
    
    print("\n📋 Available TeknoParrot Games (from UserProfiles):")
    print("=" * 80)
    print(f"{'XML Filename':<25} | {'Game Name'}")
    print("-" * 80)
    
    xml_files = sorted(profiles_path.glob("*.xml"))
    games_found = 0
    
    for xml_file in xml_files:
        try:
            tp_tree = ET.parse(xml_file)
            tp_root = tp_tree.getroot()
            
            # Try to find game name using multiple methods
            game_name = None
            
            # Search for GameNameInternal
            for elem in tp_root.iter():
                if elem.tag.endswith('GameNameInternal') and elem.text:
                    game_name = elem.text.strip()
                    break
            
            # Search for GameName
            if not game_name:
                for elem in tp_root.iter():
                    if elem.tag.endswith('GameName') and elem.text:
                        game_name = elem.text.strip()
                        break
            
            # Fallback to ProfileName
            if not game_name:
                for elem in tp_root.iter():
                    if elem.tag.endswith('ProfileName') and elem.text:
                        game_name = f"{elem.text.strip()} (ProfileName)"
                        break
            
            if game_name:
                print(f"{xml_file.stem:<25} | {game_name}")
                games_found += 1
            else:
                print(f"{xml_file.stem:<25} | [No name found]")
                
        except Exception as e:
            print(f"{xml_file.stem:<25} | [Error: {e}]")
    
    print("-" * 80)
    print(f"Total games found: {games_found}")

# Main execution
if __name__ == "__main__":
    print("🎮 TeknoParrot LaunchBox Title Cleaner v3.0")
    print("=" * 60)
    print("📁 Directory Input Version")
    print()
    
    # Try to load saved config first
    saved_xml, saved_profiles, saved_platforms = load_paths_config()
    
    if saved_xml and saved_profiles and saved_platforms:
        print("💾 Found saved configuration:")
        print(f"   LaunchBox XML: {Path(saved_xml).name}")
        print(f"   UserProfiles: {saved_profiles}")
        print()
        
        use_saved = input("Use saved paths? (y/n): ").strip().lower()
        if use_saved == 'y':
            launchbox_xml = saved_xml
            teknoparrot_profiles = saved_profiles
            
            # Verify paths still exist
            if not Path(launchbox_xml).exists():
                print(f"❌ Saved LaunchBox XML not found: {launchbox_xml}")
                launchbox_xml, teknoparrot_profiles, platforms_folder = get_user_paths()
            elif not Path(teknoparrot_profiles).exists():
                print(f"❌ Saved UserProfiles folder not found: {teknoparrot_profiles}")
                launchbox_xml, teknoparrot_profiles, platforms_folder = get_user_paths()
            else:
                print("✅ Using saved paths")
                platforms_folder = saved_platforms
        else:
            launchbox_xml, teknoparrot_profiles, platforms_folder = get_user_paths()
    else:
        launchbox_xml, teknoparrot_profiles, platforms_folder = get_user_paths()
    
    if not launchbox_xml or not teknoparrot_profiles:
        print("❌ Required paths not provided. Exiting.")
        input("Press Enter to exit...")
        exit()
    
    # Ask to save paths for future use
    if not (saved_xml and saved_profiles):
        save_config = input("\n💾 Save these paths for future use? (y/n): ").strip().lower()
        if save_config == 'y':
            save_paths_config(launchbox_xml, teknoparrot_profiles, platforms_folder)
    
    print(f"\n📋 Configuration:")
    print(f"   LaunchBox XML: {Path(launchbox_xml).name}")
    print(f"   UserProfiles: {teknoparrot_profiles}")
    print()
    
    print("Available options:")
    print("1. Show available TeknoParrot games")
    print("2. Clean up LaunchBox titles")
    
    choice = input("\nChoice (1 or 2): ").strip()
    
    if choice == "1":
        show_available_games(teknoparrot_profiles)
    elif choice == "2":
        print("\n⚠️  IMPORTANT: Make sure LaunchBox is closed before proceeding!")
        confirm = input("LaunchBox is closed and you want to continue? (y/N): ").strip().lower()
        if confirm == 'y':
            clean_teknoparrot_titles(launchbox_xml, teknoparrot_profiles)
        else:
            print("❌ Operation cancelled")
    else:
        print("❌ Invalid choice")
    
    input("\nPress Enter to exit...")