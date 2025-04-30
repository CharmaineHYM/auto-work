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

def main():
    # Get the current directory
    current_dir = Path.cwd()
    
    # Set up file paths
    excel_path = current_dir / 'TB.xlsx'
    source_json_path = current_dir / 'sourceCode' / 'index.json'
    
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
    
    # Check if required files exist
    if not excel_path.exists():
        print(f"Error: Excel file not found at {excel_path}")
        return
    
  
    if not source_json_path.exists():
        print(f"Error: Source JSON file not found at {source_json_path}")
        return
    
  
   
    
    # Load source JSON
    source_json = load_source_json(source_json_path)
    if source_json is None:
        return
    
    # Read and process the Excel file for translations only if not SEA or AU
    else:
        df = read_excel_file(excel_path)
        if df is not None:
            translations = process_translations(df, source_json)
            save_to_json(translations, output_json_path)
  

if __name__ == "__main__":
    main() 