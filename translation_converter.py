import pandas as pd
import json
import os
from pathlib import Path
import re

def read_excel_file(file_path):
    """Read the Excel file and return a DataFrame."""
    try:
        # Read the Excel file with the first row as header
        df = pd.read_excel(file_path)
        print("\nRaw Excel data preview:")
        print(df.head(20))  # Show more rows for debugging
        
        # Keep all columns
        df = df.iloc[:, [0, 1]]  # Only keep the first two columns
        df.columns = ['English', 'Translation']
        
        # Remove rows where English is NaN or empty
        df = df.dropna(subset=['English'])
        df = df[df['English'].astype(str).str.strip() != '']
        
        # Convert to string and strip whitespace
        df['English'] = df['English'].astype(str).str.strip()
        df['Translation'] = df['Translation'].fillna('').astype(str).str.strip()
        
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def load_source_json(file_path):
    """Load the source JSON file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            print("\nSource JSON values:")
            for key, value in data.items():
                print(f"{key}: {value}")
            return data
    except Exception as e:
        print(f"Error reading source JSON file: {e}")
        return None

def find_matching_key(english_text, source_json):
    """Find the matching key in source JSON for the given English text."""
    if pd.isna(english_text):
        return None
        
    english_text = str(english_text).strip()
    if not english_text or english_text.lower() in ['[video]', 'enter copy', 'content/copy en']:
        return None
    
    # First try exact match with values (replacing &nbsp; with space and removing <span> tags)
    for key, value in source_json.items():
        if isinstance(value, str):
            normalized_value = value.replace('&nbsp;', ' ').replace('<span>', '').replace('</span>', '')
            normalized_english = english_text.replace('<span>', '').replace('</span>', '')
            if normalized_value == normalized_english:
                print(f"Found exact match for '{english_text}' with key '{key}'")
                return key
    
    # If no exact match, try case-insensitive match
    for key, value in source_json.items():
        if isinstance(value, str):
            normalized_value = value.replace('&nbsp;', ' ').replace('<span>', '').replace('</span>', '').lower()
            normalized_english = english_text.replace('<span>', '').replace('</span>', '').lower()
            if normalized_value == normalized_english:
                print(f"Found case-insensitive match for '{english_text}' with key '{key}'")
                return key
    
    print(f"No match found for '{english_text}'")
    return None

def process_translations(df, source_json):
    """Process the DataFrame and create a translation dictionary."""
    translations = source_json.copy()  # Start with the source JSON structure
    updated_count = 0
    
    for _, row in df.iterrows():
        english_text = row['English']
        translation = row['Translation']
        
        # Skip if either field is empty
        if pd.isna(english_text) or pd.isna(translation):
            continue
            
        english_text = str(english_text).strip()
        translation = str(translation).strip()
        
        if not english_text or not translation or translation.lower() == 'nan':
            continue
            
        # Find the matching key in source JSON
        key = find_matching_key(english_text, source_json)
        if key:
            # Replace spaces with &nbsp; in the translation if the original had &nbsp;
            if '&nbsp;' in source_json[key]:
                translation = translation.replace(' ', '&nbsp;')
            
            translations[key] = translation
            print(f"Updated translation for '{english_text}' to '{translation}'")
            updated_count += 1
    
    print(f"\nTotal translations updated: {updated_count}")
    return translations

def save_to_json(translations, output_path):
    """Save translations to a JSON file."""
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(translations, f, ensure_ascii=False, indent=4)
        print(f"Translations saved to {output_path}")
    except Exception as e:
        print(f"Error saving JSON file: {e}")


def read_remove_slots():
    """Read the list of slot IDs to remove from removeSlot.txt"""
    try:
        print("\nTrying to read removeSlot.txt...")
        with open('sourceCode/removeSlot.txt', 'r', encoding='utf-8') as f:
            slot_ids = [line.strip() for line in f if line.strip()]
            print(f"Found slot IDs to remove: {slot_ids}")
            return slot_ids
    except FileNotFoundError:
        print("removeSlot.txt not found in sourceCode directory")
        return []
    except Exception as e:
        print(f"Error reading removeSlot.txt: {e}")
        return []

def remove_slots(html_content, slot_ids):
    """Remove code blocks and styles containing the specified slot IDs"""
    if not slot_ids:
        return html_content
    
    # First, remove the articles with matching IDs
    for slot_id in slot_ids:
        # Only match articles where the ID is an exact match
        pattern = rf'<article[^>]*\bid="{slot_id}"[^>]*>.*?</article>'
        html_content = re.sub(pattern, '', html_content, flags=re.DOTALL)
        
        # Only match style blocks that specifically target this slot ID
        # Look for #{slot_id} followed by a space, comma, opening brace, or end of line
        style_pattern = rf'<style[^>]*>([^<]*#{slot_id}[\s,{{][^<]*)</style>'
        html_content = re.sub(style_pattern, '', html_content)
        
        # Same for noscript blocks
        noscript_pattern = rf'<noscript[^>]*>([^<]*#{slot_id}[\s,{{][^<]*)</noscript>'
        html_content = re.sub(noscript_pattern, '', html_content)
    
    return html_content

def remove_slots_from_tagging(tagging_path, slot_ids):
    """Remove specified slots from tagging.json file."""
    try:
        # Read the tagging.json file
        with open(tagging_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # Fix the JSON format
        content = content.replace("'", '"')
        
        # 2. Add quotes around property names
        content = re.sub(r'(\w+):', r'"\1":', content)
        
        # 3. Remove trailing commas
        content = re.sub(r',\s*}', '}', content)
        content = re.sub(r',\s*]', ']', content)
        
        # 4. Remove empty lines
        content = '\n'.join(line for line in content.split('\n') if line.strip())
        
        print("\nTransformed JSON content:")
        print(content)
        
        # Parse the JSON
        tagging_data = json.loads(content)
        
        # Remove the specified slots from slotOrder
        if 'slotOrder' in tagging_data:
            original_slots = tagging_data['slotOrder']
            # Filter out the slots to remove and empty strings
            tagging_data['slotOrder'] = [slot for slot in original_slots if slot and slot not in slot_ids]
            removed_count = len(original_slots) - len(tagging_data['slotOrder'])
            print(f"Removed {removed_count} slots from tagging.json: {slot_ids}")
            
            # Print the updated slotOrder for verification
            print("\nUpdated slotOrder:")
            print(json.dumps(tagging_data['slotOrder'], indent=2))
        
        # Save the updated tagging.json to exportCode directory
        export_path = Path(tagging_path).parent.parent / "exportCode" / "tagging.json"
        with open(export_path, 'w', encoding='utf-8') as f:
            json.dump(tagging_data, f, indent=2, ensure_ascii=False)
        
        print(f"Updated tagging.json saved to {export_path}")
        
        # Print the updated content for verification
        print("\nUpdated tagging.json content:")
        print(json.dumps(tagging_data, indent=2))
        
        # Also update the source tagging.json file
        with open(tagging_path, 'w', encoding='utf-8') as f:
            json.dump(tagging_data, f, indent=2, ensure_ascii=False)
        print(f"Updated source tagging.json saved to {tagging_path}")
    except Exception as e:
        print(f"Error updating tagging.json: {e}")

def read_region_and_update_prefix():
    """Read region from region.txt and update prefix in tagging.json"""
    try:
        # Read region from sourceCode/region.txt
        with open('sourceCode/region.txt', 'r', encoding='utf-8') as f:
            region = f.read().strip()
            print(f"\nRegion from sourceCode/region.txt: {region}")
        
        if not region:
            print("Warning: sourceCode/region.txt is empty")
            return False
        
        # Read tagging.json
        with open('sourceCode/tagging.json', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Fix JSON format
        content = content.replace("'", '"')
        content = re.sub(r'(\w+):', r'"\1":', content)
        content = re.sub(r',\s*}', '}', content)
        content = re.sub(r',\s*]', ']', content)
        content = '\n'.join(line for line in content.split('\n') if line.strip())
        
        # Parse JSON
        tagging_data = json.loads(content)
        
        # Update prefix with exact region value
        tagging_data['prefix'] = region
        print(f"Updated prefix to: {region}")
        
        # Remove empty strings from slotOrder
        if 'slotOrder' in tagging_data:
            tagging_data['slotOrder'] = [slot for slot in tagging_data['slotOrder'] if slot]
        
        # Save updated tagging.json to both sourceCode and exportCode
        with open('sourceCode/tagging.json', 'w', encoding='utf-8') as f:
            json.dump(tagging_data, f, indent=2, ensure_ascii=False)
        print(f"Updated source tagging.json saved to sourceCode/tagging.json")
        
        with open('exportCode/tagging.json', 'w', encoding='utf-8') as f:
            json.dump(tagging_data, f, indent=2, ensure_ascii=False)
        print(f"Updated tagging.json saved to exportCode/tagging.json")
        
        return True
        
    except Exception as e:
        print(f"Error reading region and updating prefix: {e}")
        return False

def main():
    # Get the current directory
    current_dir = Path.cwd()
    
    # Set up file paths
    excel_path = current_dir / 'TB.xlsx'
    image_path_file = current_dir / 'new_image_path.txt'
    source_json_path = current_dir / 'sourceCode' / 'index.json'
    source_html_path = current_dir / 'sourceCode' / 'index.html'
    tagging_path = current_dir / 'sourceCode' / 'tagging.json'
    region_path = current_dir / 'sourceCode' / 'region.txt'
    
    # Create exportCode directory if it doesn't exist
    export_dir = current_dir / 'exportCode'
    export_dir.mkdir(exist_ok=True)
    
    # Clean up exportCode directory
    print("\nCleaning up exportCode directory...")
    for file in export_dir.glob('*'):
        if file.is_file():
            file.unlink()
            print(f"Removed {file.name}")
    
    # Set up output paths
    output_json_path = export_dir / 'translations.json'
    output_html_path = export_dir / 'index.html'
    output_tagging_path = export_dir / 'tagging.json'
    output_index_json_path = export_dir / 'index.json'
    
    # Check if required files exist
    if not excel_path.exists():
        print(f"Error: Excel file not found at {excel_path}")
        return
    
    if not image_path_file.exists():
        print(f"Error: Image path file not found at {image_path_file}")
        return
    
    if not source_json_path.exists():
        print(f"Error: Source JSON file not found at {source_json_path}")
        return
    
    if not source_html_path.exists():
        print(f"Error: HTML file not found at {source_html_path}")
        return
    
    # First, read the source HTML file
    with open(source_html_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # Read slot IDs to remove
    print("\nReading slot IDs to remove...")
    slot_ids = read_remove_slots()
    if slot_ids:
        print(f"Found slot IDs to remove: {slot_ids}")
        html_content = remove_slots(html_content, slot_ids)
        print("Removed specified slots from HTML")
    
    # Read region and update prefix
    region = None
    if region_path.exists():
        with open(region_path, 'r', encoding='utf-8') as f:
            region = f.read().strip()
            print(f"\nRegion from sourceCode/region.txt: {region}")
            
        if region and tagging_path.exists():
            # Read the latest tagging.json from sourceCode
            with open(tagging_path, 'r', encoding='utf-8') as f:
                content = f.read()
                
            # Fix the JSON format
            content = content.replace("'", '"')  # Replace single quotes with double quotes
            content = re.sub(r'(\w+)\s*:', r'"\1":', content)  # Add quotes around property names
            content = re.sub(r',\s*}', '}', content)  # Remove trailing commas
            content = re.sub(r',\s*]', ']', content)  # Remove trailing commas
            content = '\n'.join(line for line in content.split('\n') if line.strip())  # Remove empty lines
            
            print("\nTransformed JSON content:")
            print(content)
            
            # Parse and update the JSON
            tagging_data = json.loads(content)
            tagging_data['prefix'] = region
            print(f"Updated prefix to: {region}")
            
            # Remove slots from slotOrder if they exist
            if 'slotOrder' in tagging_data and slot_ids:
                original_slots = tagging_data['slotOrder']
                tagging_data['slotOrder'] = [slot for slot in original_slots if slot and slot not in slot_ids]
                removed_count = len(original_slots) - len(tagging_data['slotOrder'])
                print(f"Removed {removed_count} slots from tagging.json slotOrder")
                print("Updated slotOrder:", tagging_data['slotOrder'])
            
            # Save the updated tagging.json to exportCode directory
            with open(output_tagging_path, 'w', encoding='utf-8') as f:
                json.dump(tagging_data, f, indent=2, ensure_ascii=False)
            print(f"Updated tagging.json saved to {output_tagging_path}")
    
    # Copy index.json from sourceCode to exportCode
    try:
        with open(source_json_path, 'r', encoding='utf-8') as f:
            index_data = json.load(f)
        
        # Update slotOrder in index.json if needed
        if slot_ids and 'slotOrder' in index_data:
            original_slots = index_data['slotOrder']
            index_data['slotOrder'] = [slot for slot in original_slots if slot and slot not in slot_ids]
            removed_count = len(original_slots) - len(index_data['slotOrder'])
            print(f"Removed {removed_count} slots from index.json slotOrder")
            print("Updated index.json slotOrder:", index_data['slotOrder'])
        
        # Save the updated index.json to exportCode directory
        with open(output_index_json_path, 'w', encoding='utf-8') as f:
            json.dump(index_data, f, indent=2, ensure_ascii=False)
        print(f"Updated index.json saved to {output_index_json_path}")
    except Exception as e:
        print(f"Error copying index.json: {e}")
    
    # Skip translations for SEA and AU regions
    skip_translations = region and region.startswith(('SEA', 'AU'))
    if skip_translations:
        print(f"\nSkipping translations for region: {region}")
    
    # Read image path from text file
    try:
        with open(image_path_file, 'r', encoding='utf-8') as f:
            new_image_path = f.read().strip()
            print(f"\nNew image path: {new_image_path}")
            
        if not new_image_path:
            print("Error: Empty image path found in text file")
            return
            
    except Exception as e:
        print(f"Error reading image path from text file: {e}")
        return
    
    # Load source JSON
    source_json = load_source_json(source_json_path)
    if source_json is None:
        return
    
    # Read and process the Excel file for translations only if not SEA or AU
    if not skip_translations:
        df = read_excel_file(excel_path)
        if df is not None:
            translations = process_translations(df, source_json)
            save_to_json(translations, output_json_path)
    
    # First save the HTML file
    with open(output_html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    print(f"Initial HTML saved to {output_html_path}")
    
    # Then update image paths in the saved HTML file
    # update_image_paths(output_html_path, new_image_path, output_html_path)
    # print(f"Updated HTML with image paths saved to {output_html_path}")

if __name__ == "__main__":
    main() 