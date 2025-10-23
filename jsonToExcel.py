import os
import json
import pandas as pd
from datetime import datetime

def safe_get(data, path, default=""):
    """
    Safely navigate through nested dictionary/list structures
    """
    try:
        keys = path.split('.')
        current = data
        for key in keys:
            if isinstance(current, dict):
                current = current.get(key, default)
            elif isinstance(current, list) and key.isdigit():
                index = int(key)
                if 0 <= index < len(current):
                    current = current[index]
                else:
                    return default
            else:
                return default
            if current is None:
                return default
        return current
    except:
        return default

def format_date(year, month=None, day=None):
    """
    Format date components into a readable string
    """
    if not year:
        return ""
    
    try:
        if month and day:
            return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
        elif month:
            return f"{year}-{month.zfill(2)}"
        else:
            return str(year)
    except:
        return str(year) if year else ""

def process_json_folder_to_excel(folder_path, output_filename):
    """
    Reads all JSON files from a folder, processes nested data according to mappings,
    and saves the combined data to a single Excel file.
    
    Args:
        folder_path (str): The path to the folder containing JSON files.
        output_filename (str): The name of the output Excel file (e.g., 'output.xlsx').
    """
    all_records = []

    # Check if the folder exists
    if not os.path.isdir(folder_path):
        print(f"Error: Folder '{folder_path}' not found.")
        return

    print(f"Processing JSON files from '{folder_path}'...")
    
    # Counter for processed files
    processed_files = 0
    
    # 1. Loop through each file in the specified folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.json'):
            file_path = os.path.join(folder_path, filename)
            processed_files += 1
            
            if processed_files % 100 == 0:
                print(f"Processed {processed_files} files...")
            
            with open(file_path, 'r', encoding='utf-8') as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError as e:
                    print(f"Warning: Could not decode JSON from {filename}. Error: {e}. Skipping.")
                    continue

                # Extract response data
                response = safe_get(data, 'response', {})
                if not response:
                    print(f"Warning: No response data found in {filename}. Skipping.")
                    continue

                # Get company information
                company_name = safe_get(response, 'organisationName', '')
                company_org_id = safe_get(response, 'OrgId', '')
                
                # Process each officer in the officersList
                officers_list = safe_get(response, 'officersList', [])
                
                for officer_idx, officer in enumerate(officers_list):
                    if not isinstance(officer, dict):
                        continue
                    
                    # Create a record for this officer
                    record = {}
                    
                    # Basic officer information
                    record['CompanyName'] = company_name
                    record['CompanyOrgID'] = company_org_id
                    record['OfficerID'] = safe_get(officer, 'id', '')
                    record['PersonID'] = safe_get(officer, 'person.id', '')
                    record['Status'] = safe_get(officer, 'status', '')
                    
                    # Person information
                    person_info = safe_get(officer, 'PersonInformation', {})
                    name_info = safe_get(person_info, 'Name', {})
                    
                    record['Prefix'] = safe_get(name_info, 'Prefix', '')
                    record['FirstName'] = safe_get(name_info, 'FirstName', '')
                    record['LastName'] = safe_get(name_info, 'LastName', '')
                    record['Initial/Middle'] = safe_get(name_info, 'Middle/Initial', '')
                    record['Suffix'] = safe_get(name_info, 'Suffix', '')
                    record['Age'] = safe_get(name_info, 'Age', '')
                    record['Sex'] = safe_get(name_info, 'Sex', '')
                    
                    # Biography
                    bio_info = safe_get(officer, 'BiographicalInformation', {})
                    bio_text = safe_get(bio_info, 'Text', {})
                    record['Biography'] = safe_get(bio_text, '_', '')
                    
                    # Position History (variable number)
                    position_info = safe_get(officer, 'PositionInformation', {})
                    titles = safe_get(position_info, 'Titles', [])
                    
                    for title_idx, title in enumerate(titles, 1):
                        if isinstance(title, dict):
                            record[f'LongTitle_{title_idx}'] = safe_get(title, 'LongTitle', '')
                            
                            # Start date
                            start_date = safe_get(title, 'Start', {})
                            start_year = safe_get(start_date, 'year', '')
                            start_month = safe_get(start_date, 'month', '')
                            start_day = safe_get(start_date, 'day', '')
                            record[f'StartDate_{title_idx}'] = format_date(start_year, start_month, start_day)
                            
                            # End date
                            end_date = safe_get(title, 'End', {})
                            if end_date:
                                end_year = safe_get(end_date, 'year', '')
                                end_month = safe_get(end_date, 'month', '')
                                end_day = safe_get(end_date, 'day', '')
                                record[f'EndDate_{title_idx}'] = format_date(end_year, end_month, end_day)
                            else:
                                record[f'EndDate_{title_idx}'] = 'Present'
                    
                    # Corporate Affiliations (variable number)
                    corp_affiliations = safe_get(officer, 'CorporateAffiliations', [])
                    
                    for aff_idx, affiliation in enumerate(corp_affiliations, 1):
                        if isinstance(affiliation, dict):
                            company_info = safe_get(affiliation, 'Company', {})
                            officer_info = safe_get(affiliation, 'Officer', {})
                            
                            record[f'AffiliationCompanyName_{aff_idx}'] = safe_get(company_info, 'name', '')
                            record[f'AffiliationCompanyOrgID_{aff_idx}'] = safe_get(company_info, 'orgid', '')
                            record[f'AffiliationTitle_{aff_idx}'] = safe_get(officer_info, 'title', '')
                            record[f'AffiliationActiveStatus_{aff_idx}'] = safe_get(officer_info, 'active', '')
                    
                    # Salary Information (variable number)
                    salary_info = safe_get(officer, 'SalaryInformation', {})
                    compensation_periods = safe_get(salary_info, 'CompensationPeriod', [])
                    
                    for comp_idx, period in enumerate(compensation_periods, 1):
                        if isinstance(period, dict):
                            submission = safe_get(period, 'Submission', {})
                            record[f'SalaryYear_{comp_idx}'] = safe_get(submission, 'year', '')
                            
                            # StandardizedCompensation
                            std_comp = safe_get(period, 'StandardizedCompensation', [])
                            
                            # Find FYT (index 5) and RSA (index 2)
                            for comp_item in std_comp:
                                if isinstance(comp_item, dict):
                                    coa = safe_get(comp_item, 'coa', '')
                                    value = safe_get(comp_item, '_', '')
                                    
                                    if coa == 'FYT':
                                        record[f'FYT_{comp_idx}'] = value
                                    elif coa == 'RSA':
                                        record[f'RSA_{comp_idx}'] = value
                    
                    # Education Information (variable number)
                    education_history = safe_get(person_info, 'EducationHistory', [])
                    
                    for edu_idx, education in enumerate(education_history, 1):
                        if isinstance(education, dict):
                            record[f'College_{edu_idx}'] = safe_get(education, 'College._', '')
                            record[f'Degree_{edu_idx}'] = safe_get(education, 'Degree._', '')
                            record[f'Major_{edu_idx}'] = safe_get(education, 'Major._', '')
                            
                            graduation = safe_get(education, 'Graduation', {})
                            record[f'Graduation_{edu_idx}'] = safe_get(graduation, 'year', '')
                    
                    # Committee Information (variable number)
                    committee_memberships = safe_get(position_info, 'CommitteeMemberships', [])
                    
                    for comm_idx, committee in enumerate(committee_memberships, 1):
                        if isinstance(committee, dict):
                            record[f'CommitteeName_{comm_idx}'] = safe_get(committee, 'CommitteeName', '')
                            record[f'CommitteeTitle_{comm_idx}'] = safe_get(committee, 'Title', '')
                            
                            start_date = safe_get(committee, 'Start', {})
                            record[f'CommitteeStartDate_{comm_idx}'] = safe_get(start_date, 'year', '')
                    
                    all_records.append(record)

    if not all_records:
        print("No valid records were processed.")
        return

    print(f"Processed {processed_files} files and created {len(all_records)} records.")
    print("Creating DataFrame...")

    # 3. Create a pandas DataFrame from the list of dictionaries
    df = pd.DataFrame(all_records)
    
    # Optional: For better organization, sort the columns
    static_cols = sorted([col for col in df.columns if '_' not in col])
    dynamic_cols = sorted([col for col in df.columns if '_' in col])
    df = df[static_cols + dynamic_cols]

    print(f"DataFrame created with {len(df)} rows and {len(df.columns)} columns.")
    print("Saving to Excel...")

    # 4. Save the DataFrame to an Excel file
    try:
        df.to_excel(output_filename, index=False, engine='openpyxl')
        print(f"Successfully created '{output_filename}' with data from {len(all_records)} records.")
        print(f"Excel file contains {len(df)} rows and {len(df.columns)} columns.")
    except Exception as e:
        print(f"An error occurred while saving to Excel: {e}")

if __name__ == "__main__":
    JSON_FOLDER = 'JSON-DATA-ALL'
    OUTPUT_EXCEL_FILE = 'officers_data.xlsx'
    
    print("Starting JSON to Excel conversion...")
    print(f"Input folder: {JSON_FOLDER}")
    print(f"Output file: {OUTPUT_EXCEL_FILE}")
    print("-" * 50)
    
    process_json_folder_to_excel(JSON_FOLDER, OUTPUT_EXCEL_FILE)
    
    print("-" * 50)
    print("Conversion completed!")



