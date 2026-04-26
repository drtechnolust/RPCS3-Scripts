import xml.etree.ElementTree as ET
import os
import shutil
from pathlib import Path

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
        print(f"✅ Loaded LaunchBox XML: {launchbox_xml_path}")
    except ET.ParseError as e:
        print(f"❌ Error parsing LaunchBox XML: {e}")
        return
    except Exception as e:
        print(f"❌ Error loading LaunchBox XML: {e}")
        return
    
    # Debug: Show LaunchBox XML structure
    print(f"📋 LaunchBox XML root element: <{root.tag}>")
    game_elements = root.findall('.//Game')
    print(f"📋 Found {len(game_elements)} games in LaunchBox XML")
    
    # Create mapping of TeknoParrot games
    tp_game_names = {}
    xml_files_found = list(profiles_path.glob("*.xml"))
    print(f"📁 Found {len(xml_files_found)} XML files in UserProfiles folder")
    
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
        print("❌ No game names found in TeknoParrot XML files. Check your UserProfiles folder.")
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
            print(f"\n💾 Backup created: {backup_path}")
            
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
    
    print("📋 Available TeknoParrot Games (from UserProfiles):")
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

def debug_xml_structure(xml_path):
    """Debug function to show XML structure"""
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        print(f"Root element: <{root.tag}>")
        print(f"Root attributes: {root.attrib}")
        
        print("\nFirst few child elements:")
        for i, child in enumerate(root):
            if i < 10:  # Show first 10 elements
                print(f"  <{child.tag}> = '{child.text[:50]}...'" if child.text and len(child.text) > 50 else f"  <{child.tag}> = '{child.text}'" if child.text else f"  <{child.tag}>")
            
        print(f"\nAll unique element names found:")
        elements = set()
        for elem in root.iter():
            elements.add(elem.tag)
        for elem in sorted(elements):
            print(f"  {elem}")
            
    except Exception as e:
        print(f"Error: {e}")

# Main execution
if __name__ == "__main__":
    print("🎮 TeknoParrot LaunchBox Title Cleaner v2.2")
    print("=" * 60)
    
    # YOUR ACTUAL PATHS
    launchbox_platforms_folder = r"C:\Arcade\LaunchBox\Data\Platforms"
    teknoparrot_userprofiles = r"C:\Arcade\LaunchBox\Emulators\TeknoParrot\UserProfiles"
    
    # Assume TeknoParrot.xml is the platform file
    launchbox_xml = os.path.join(launchbox_platforms_folder, "TeknoParrot.xml")
    
    print(f"📁 LaunchBox XML: {launchbox_xml}")
    print(f"📁 TeknoParrot UserProfiles: {teknoparrot_userprofiles}")
    print()
    
    # Check if files exist
    if not os.path.exists(launchbox_xml):
        print("❌ TeknoParrot.xml not found in Platforms folder.")
        print("Available XML files in Platforms folder:")
        platforms_path = Path(launchbox_platforms_folder)
        if platforms_path.exists():
            for xml_file in platforms_path.glob("*.xml"):
                print(f"   📄 {xml_file.name}")
        print("\nPlease update the script with the correct platform XML filename.")
        exit()
    
    if not os.path.exists(teknoparrot_userprofiles):
        print(f"❌ UserProfiles folder not found: {teknoparrot_userprofiles}")
        exit()
    
    print("Available options:")
    print("1. Show available TeknoParrot games")
    print("2. Clean up LaunchBox titles")
    print("3. Debug LaunchBox XML structure")
    print("4. Debug TeknoParrot XML structure")
    
    choice = input("\nChoice (1/2/3/4): ").strip()
    
    if choice == "1":
        show_available_games(teknoparrot_userprofiles)
    elif choice == "2":
        print("\n⚠️  IMPORTANT: Make sure LaunchBox is closed before proceeding!")
        confirm = input("LaunchBox is closed and you want to continue? (y/N): ").strip().lower()
        if confirm == 'y':
            clean_teknoparrot_titles(launchbox_xml, teknoparrot_userprofiles)
        else:
            print("❌ Operation cancelled")
    elif choice == "3":
        debug_xml_structure(launchbox_xml)
    elif choice == "4":
        userprofiles_path = Path(teknoparrot_userprofiles)
        xml_files = list(userprofiles_path.glob("*.xml"))
        if xml_files:
            print("Available UserProfile XML files:")
            for i, xml_file in enumerate(xml_files[:10]):  # Show first 10
                print(f"  {i+1}. {xml_file.name}")
            
            try:
                choice_idx = int(input(f"\nChoose file to debug (1-{min(10, len(xml_files))}): ")) - 1
                if 0 <= choice_idx < len(xml_files):
                    debug_xml_structure(xml_files[choice_idx])
                else:
                    print("❌ Invalid choice")
            except ValueError:
                print("❌ Invalid number")
        else:
            print("❌ No XML files found in UserProfiles folder")
    else:
        print("❌ Invalid choice")