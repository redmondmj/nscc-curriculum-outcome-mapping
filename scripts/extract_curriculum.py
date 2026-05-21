import os
import json
import glob
from bs4 import BeautifulSoup
import re

def extract_course_data(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        soup = BeautifulSoup(f, 'html.parser')
    
    title_text = soup.title.string if soup.title else ""
    course_code = ""
    course_name = ""
    
    if ":" in title_text:
        course_code, course_name = [x.strip() for x in title_text.split(":", 1)]
    else:
        # fallback
        course_code = os.path.basename(filepath).split("_")[0].strip()
        course_name = title_text
        
    text = soup.get_text(separator='\n', strip=True)
    lines = text.split('\n')
    
    outcomes = []
    current_outcome = None
    parsing_objectives = False
    
    for i in range(len(lines)):
        line = lines[i].strip()
        
        # Stop parsing if we hit the end of outcomes section
        if line.startswith("Learning Outcomes Display") or line.startswith("Other Course Notes"):
            break
            
        if line == "Outcome" and i + 1 < len(lines):
            if current_outcome:
                outcomes.append(current_outcome)
            current_outcome = {
                "outcome_text": lines[i+1].strip(),
                "objectives": []
            }
            parsing_objectives = False
            
        elif line == "Objectives":
            parsing_objectives = True
            
        elif parsing_objectives and current_outcome:
            # Add objective lines if they aren't "Outcome" or empty
            if line and line != "Outcome" and line != "Objectives":
                current_outcome["objectives"].append(line)
                
    if current_outcome:
        outcomes.append(current_outcome)
        
    return {
        "course_code": course_code,
        "course_name": course_name,
        "outcomes": outcomes
    }

def main():
    base_dir      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    raw_dir       = os.path.join(base_dir, "data", "raw", "itsm_curriculum")
    processed_dir = os.path.join(base_dir, "data", "processed")
    
    all_courses = []
    
    # Process RTF files which are actually HTML
    rtf_files = glob.glob(os.path.join(raw_dir, "*.rtf"))
    print(f"Found {len(rtf_files)} .rtf curriculum files.")
    
    for filepath in rtf_files:
        try:
            data = extract_course_data(filepath)
            all_courses.append(data)
            print(f"Extracted: {data['course_code']} - {data['course_name']} ({len(data['outcomes'])} outcomes)")
        except Exception as e:
            print(f"Error parsing {filepath}: {e}")
            
    output_file = os.path.join(processed_dir, "curriculum_extracted.json")
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(all_courses, f, indent=4)
        
    print(f"Successfully saved {len(all_courses)} courses to {output_file}")

if __name__ == "__main__":
    main()
