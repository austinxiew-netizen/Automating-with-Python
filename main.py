from pathlib import Path
import pandas as pd
import sys


# 1 set path
def receive_path(input):
    """
    Receive the path and return the compatible standardized path
    """

    input_path = Path(input).resolve()
    # Determine whether the input path exists
    if not input_path.exists():
        raise SystemExit(f'Error: {input_path} The path does not exist')
        
    return input_path

# 2. Extract the file list. Not considering sub-files (no recursion)
def find_file(input_path):
    '''
    Receive a path. If it is an exlce file, store it in the list and return it.
    If it is a folder, find all the exlce files in it and store them in the list.
    If it is not exlce and not a folder, an error will be reported and you will exit.
    '''
    file_names = []
    # Retain the exists() check to ensure that the function can be called independently.
    if not input_path.exists():
        raise SystemExit(f'Error: {input_path} The path does not exist.')
    # find a single xlsx file
    if input_path.is_file():
        if input_path.suffix.lower() != '.xlsx':
            raise SystemExit(f"{input_path}It's not a directory or an excel file. Please check.")
        else:
            file_names.append(input_path)
            return file_names
    # Find the xlsx file in the folder
    else:        
        for file in input_path.iterdir():            
            if file.suffix.lower() == '.xlsx':
                file_names.append(file)
        if not file_names:
            raise SystemExit(f'{input_path} No excel files were found')
    return file_names

# 3. Rules: Header mapping & Unit conversion
def set_mapping_rules():
    rules = {}

    # Standardize column headers
    rules["field_mapping"] = {
        "Vacancy": "vacancy_rate",
        "Vacancy %": "vacancy_rate",
        "Vacancy Rate": "vacancy_rate",

        "Rent": "asking_rent",
        "Average Rent": "asking_rent",
        "Rent (USD/sqft)": "asking_rent",

        "Absorption": "net_absorption",
        "Take-up": "net_absorption",
        "Net Absorption": "net_absorption"
    }

    # Clean units and special characters
    rules["value_keywords"] = {
        "k": 1000,
        "M": 1000000,
        "%": 0.01,
        "$": "",
        ",": ""
    }

    # Final output column order
    rules["output_fields"] = ["vacancy_rate", "asking_rent", "net_absorption", "source_file"]

    return rules

def read_excel_file(file_path, rules):
    '''
    Read the file and initially process the data header
    Add source column
    '''
    df = pd.read_excel(file_path)
    field_mapping = rules["field_mapping"]
    # Uniform header
    df.columns = [field_mapping.get(col, col) for col in df.columns]
    data = df.to_dict(orient='records')  #  List[Dict]
    # Append the source file name
    for row in data:
        row["source_file"] = file_path.name
    print(f"Processed {file_path.name} ({len(data)} rows)")

    return data 

# 5. Value Normalization 
def clean_value(text, rules):
    """Normalize raw values based on unit and currency rules."""
    if pd.isna(text) or str(text).strip() == "":
        return None

    # Strip whitespace and thousands separators
    text = str(text).strip().replace(',', '')

    # Apply scaling rules (e.g., k, M, %)
    for unit, factor in rules["value_keywords"].items():
        if unit in text:
            try:
                clean_text = text.replace(unit, '').strip()
                return float(clean_text) * factor if isinstance(factor, (int, float)) else clean_text
            except (ValueError, TypeError):
                return text

    # Direct conversion
    try:
        return float(text)
    except:
        return text
    

def clean_data(data, rules):
    """
    Master data cleaning function
    Delete rows suspected of being "duplicate headers" (all values equal to field names or containing keywords)
    - Use field mapping to convert the original header to a standard field
    - Use clean_value to uniformly handle unit and numeric conversions
    """

    cleaned = []  

    # Extract all standard field names to determine if they are duplicate headers
    field_keys = set(rules["field_mapping"].values())

    for row in data:
        # Filter out the file name field and only retain the data field for judgment
        check_values = {k: v for k, v in row.items() if k != "source_file"}       
        values_str = [str(v).strip().lower() for v in check_values.values()]

        if all(val == "" or val == "nan" for val in values_str):
            continue

        #  All values are in the standard field list (possibly duplicate headers)
        condition_all_fields = all(val in field_keys for val in values_str)

        keywords = ["vacancy", "rent", "takeup", "absorption", "area", "sqm", "sqft","summary", "total"]
        condition_contains_keywords = any(
            keyword in val for val in values_str for keyword in keywords
        )

        # If any of the above conditions is met, skip this line
        if condition_all_fields or condition_contains_keywords:
            continue

        # Data cleaning
        new_row = {}          
        for key in field_keys:
            if key in row:
                new_row[key] = clean_value(row[key], rules)

        # Source file name
        if "source_file" in row:
            new_row["source_file"] = row["source_file"]
        cleaned.append(new_row)

    return cleaned  

def save_to_excel(data, output_path, rules):
    """
    Write the cleaned data into an Excel file
    """
    if not data:
        print("No data can be written Excel")
        return

    df = pd.DataFrame(data)  

    # Force column order (automatically fill NaN for missing fields)
    output_fields = rules.get("output_fields", df.columns.tolist())
    df = df.reindex(columns=output_fields)

 
    out_file = output_path / "merged_cleaned.xlsx"
    try:
        df.to_excel(out_file, index=False)
        print(f" The result has been saved to:{out_file}")
    except PermissionError:
        print(f"Write failure: The file may be open or there may be no write permission {out_file}")
    except OSError as e:
        print(f" Write failure (path or system error) :{e}")
    except Exception as e:
        print(f"Write failure (Unknown error)：{e}")

def main(urse_path):

    input_path = receive_path(urse_path)

    if input_path.is_dir():
        output_path = input_path / 'output_files'
    else:
        output_path = input_path.parent / 'output_files'

    output_path.mkdir(exist_ok=True)

    rules = set_mapping_rules()
    file_names = find_file(input_path)

    all_data = []
    for file in file_names:
        print(f"Dealing with documents：{file.name}")
        raw_data = read_excel_file(file, rules)
        
        if not raw_data:
            continue
        cleaned = clean_data(raw_data, rules)
        all_data.extend(cleaned)

    save_to_excel(all_data, output_path, rules)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage:python script.py <Input path>")
        sys.exit(1)
    
    input_arg = sys.argv[1]

    main(input_arg)
