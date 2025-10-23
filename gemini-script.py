# import os
# import json
# import pandas as pd

# def get_nested_value(data_dict, path, default=None):
#     """
#     Safely accesses a nested value in a dictionary using a list of keys (path).
#     Returns a default value if the path is not found.
#     """
#     for key in path:
#         if isinstance(data_dict, dict):
#             data_dict = data_dict.get(key)
#         else:
#             return default
#     return data_dict if data_dict is not None else default

# def process_json_folder_to_excel(folder_path, output_filename):
#     """
#     Processes all JSON files in a folder, extracts officer data according to
#     pre-defined mappings, and saves the combined data to a single Excel file.
#     """
#     all_officers_data = []
#     processed_files = 0

#     # --- 1. Loop through all files in the specified folder ---
#     if not os.path.isdir(folder_path):
#         print(f"❌ Error: Folder '{folder_path}' not found. Please create it and add your JSON files.")
#         return

#     for filename in os.listdir(folder_path):
#         if filename.endswith('.json'):
#             file_path = os.path.join(folder_path, filename)
#             print(f"Processing file: {filename}...")

#             try:
#                 with open(file_path, 'r', encoding='utf-8') as f:
#                     data = json.load(f)
#             except Exception as e:
#                 print(f"  - Warning: Could not read or decode {filename}. Skipping. Error: {e}")
#                 continue

#             # Safely get the main list of officers to iterate through
#             officers_list = get_nested_value(data, ['response', 'officersList'], [])
#             if not officers_list:
#                 print(f"  - Warning: No 'officersList' found in {filename}. Skipping.")
#                 continue

#             # Extract company-level information for the current file
#             company_name = get_nested_value(data, ['response', 'organisationName'])
#             company_org_id = get_nested_value(data, ['response', 'OrgId'])

#             # --- 2. Extract data for each officer in the current file ---
#             for officer in officers_list:
#                 officer_record = {}

#                 # Static Information
#                 officer_record['CompanyName'] = company_name
#                 officer_record['CompanyOrgID'] = company_org_id
#                 officer_record['SourceFile'] = filename # Good practice to track origin
#                 officer_record['OfficerID'] = officer.get('id')
#                 officer_record['PersonID'] = get_nested_value(officer, ['Person', 'id'])
#                 # ... (all other static fields)
#                 officer_record['Status'] = officer.get('status')
#                 officer_record['Prefix'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Prefix'])
#                 officer_record['FirstName'] = get_nested_value(officer, ['PersonInformation', 'Name', 'FirstName'])
#                 officer_record['LastName'] = get_nested_value(officer, ['PersonInformation', 'Name', 'LastName'])
#                 officer_record['Initial/Middle'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Middle/Initial'])
#                 officer_record['Suffix'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Suffix'])
#                 officer_record['Age'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Age'])
#                 officer_record['Sex'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Sex'])
#                 officer_record['Biography'] = get_nested_value(officer, ['BiographicalInformation', 'Text', '_'])


#                 # Variable (List) Information
#                 # Position History
#                 positions = get_nested_value(officer, ['PositionInformation', 'Titles'], [])
#                 if isinstance(positions, dict): positions = [positions]
#                 for i, pos in enumerate(positions, start=1):
#                     officer_record[f'Position_LongTitle_{i}'] = pos.get('LongTitle')
#                     s_year, s_month, s_day = str(get_nested_value(pos, ['Start', 'year'], '')), str(get_nested_value(pos, ['Start', 'month'], '')), str(get_nested_value(pos, ['Start', 'day'], ''))
#                     officer_record[f'Position_StartDate_{i}'] = '-'.join(filter(None, [s_year, s_month, s_day]))
#                     e_year, e_month, e_day = str(get_nested_value(pos, ['End', 'year'], '')), str(get_nested_value(pos, ['End', 'month'], '')), str(get_nested_value(pos, ['End', 'day'], ''))
#                     end_date_str = '-'.join(filter(None, [e_year, e_month, e_day]))
#                     officer_record[f'Position_EndDate_{i}'] = end_date_str if end_date_str else 'Present'
                
#                 # ... (all other variable list processing logic remains the same)
#                 # Corporate Affiliations
#                 affiliations = get_nested_value(officer, ['CorporateAffiliations'], [])
#                 if isinstance(affiliations, dict): affiliations = [affiliations]
#                 for i, aff in enumerate(affiliations, start=1):
#                     officer_record[f'Affiliation_CompanyName_{i}'] = get_nested_value(aff, ['Company', 'name'])
#                     officer_record[f'Affiliation_CompanyOrgID_{i}'] = get_nested_value(aff, ['Company', 'orgid'])
#                     officer_record[f'Affiliation_Title_{i}'] = get_nested_value(aff, ['Officer', 'title'])
#                     officer_record[f'Affiliation_ActiveStatus_{i}'] = get_nested_value(aff, ['Officer', 'active'])

#                 # Salary
#                 salaries = get_nested_value(officer, ['SalaryInformation', 'CompensationPeriod'], [])
#                 if isinstance(salaries, dict): salaries = [salaries]
#                 for i, sal in enumerate(salaries, start=1):
#                     officer_record[f'Salary_Year_{i}'] = get_nested_value(sal, ['Submission', 'year'])
#                     std_comp = sal.get('StandardizedCompensation', [])
#                     if isinstance(std_comp, list):
#                         officer_record[f'Salary_FYT_{i}'] = next((item.get('_') for item in std_comp if item.get('coa') == 'FYT'), None)
#                         officer_record[f'Salary_RSA_{i}'] = next((item.get('_') for item in std_comp if item.get('coa') == 'RSA'), None)
                
#                 # Append the fully processed officer record to the main list
#                 all_officers_data.append(officer_record)
            
#             processed_files += 1

#     # --- 3. Create DataFrame and save to Excel ---
#     if not all_officers_data:
#         print("No officer data was processed. The output file will not be created.")
#         return
        
#     df = pd.DataFrame(all_officers_data)
    
#     try:
#         df.to_excel(output_filename, index=False, engine='openpyxl')
#         print(f"\n✅ Success! Data for {len(all_officers_data)} officers from {processed_files} files has been saved to '{output_filename}'")
#     except Exception as e:
#         print(f"\n❌ Error: Could not save the Excel file. Reason: {e}")


# # --- HOW TO USE ---
# if __name__ == "__main__":
#     # 1. Create a folder named 'json_files' (or any name you like).
#     # 2. Place all your JSON files inside this folder.
#     # 3. Make sure this script is in the same parent directory as your folder.
#     JSON_FOLDER_PATH = 'JSON-DATA-ALL'
#     OUTPUT_EXCEL_FILE = 'combined_officers_output.xlsx'
    
#     process_json_folder_to_excel(JSON_FOLDER_PATH, OUTPUT_EXCEL_FILE)


import os
import json
import pandas as pd
import re

def get_nested_value(data_dict, path, default=None):
    """
    Safely accesses a nested value in a dictionary using a list of keys (path).
    Returns a default value if the path is not found.
    """
    for key in path:
        if isinstance(data_dict, dict):
            data_dict = data_dict.get(key)
        else:
            return default
    return data_dict if data_dict is not None else default

def get_max_index(data_list, prefix):
    """
    Finds the highest index number (e.g., the '5' in 'Position_Title_5')
    for a given prefix across all processed records.
    """
    max_i = 0
    # We use a regex to reliably find the number at the end of the key
    prog = re.compile(rf"^{re.escape(prefix)}(\d+)$")
    for record in data_list:
        for key in record.keys():
            match = prog.match(key)
            if match:
                num = int(match.group(1))
                if num > max_i:
                    max_i = num
    return max_i

def process_json_folder_to_excel(folder_path, output_filename):
    """
    Processes all JSON files in a folder, extracts officer data according to
    pre-defined mappings, and saves the combined data to a single Excel file
    with guaranteed column ordering.
    """
    all_officers_data = []
    processed_files = 0

    # --- 1. Loop through all files in the specified folder ---
    if not os.path.isdir(folder_path):
        print(f"❌ Error: Folder '{folder_path}' not found. Please create it and add your JSON files.")
        return

    for filename in os.listdir(folder_path):
        if filename.endswith('.json'):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing file: {filename}...")

            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception as e:
                print(f"  - Warning: Could not read or decode {filename}. Skipping. Error: {e}")
                continue

            officers_list = get_nested_value(data, ['response', 'officersList'], [])
            if not officers_list:
                print(f"  - Warning: No 'officersList' found in {filename}. Skipping.")
                continue

            company_name = get_nested_value(data, ['response', 'organisationName'])
            company_org_id = get_nested_value(data, ['response', 'OrgId'])

            # --- 2. Extract data for each officer in the current file ---
            for officer in officers_list:
                # Use a dictionary for flexible key addition
                officer_record = {}

                # --- Static Information ---
                officer_record['CompanyName'] = company_name
                officer_record['CompanyOrgID'] = company_org_id
                officer_record['SourceFile'] = filename
                officer_record['OfficerID'] = officer.get('id')
                officer_record['PersonID'] = get_nested_value(officer, ['Person', 'id'])
                officer_record['Status'] = officer.get('status')
                officer_record['Prefix'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Prefix'])
                officer_record['FirstName'] = get_nested_value(officer, ['PersonInformation', 'Name', 'FirstName'])
                officer_record['LastName'] = get_nested_value(officer, ['PersonInformation', 'Name', 'LastName'])
                officer_record['Initial/Middle'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Middle/Initial'])
                officer_record['Suffix'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Suffix'])
                officer_record['Age'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Age'])
                officer_record['Sex'] = get_nested_value(officer, ['PersonInformation', 'Name', 'Sex'])
                officer_record['Biography'] = get_nested_value(officer, ['BiographicalInformation', 'Text', '_'])


                # --- Variable (List) Information ---
                
                # Position History
                positions = get_nested_value(officer, ['PositionInformation', 'Titles'], [])
                if isinstance(positions, dict): positions = [positions]
                for i, pos in enumerate(positions, start=1):
                    officer_record[f'Position_LongTitle_{i}'] = pos.get('LongTitle')
                    s_year, s_month, s_day = str(get_nested_value(pos, ['Start', 'year'], '')), str(get_nested_value(pos, ['Start', 'month'], '')), str(get_nested_value(pos, ['Start', 'day'], ''))
                    officer_record[f'Position_StartDate_{i}'] = '-'.join(filter(None, [s_year, s_month, s_day]))
                    e_year, e_month, e_day = str(get_nested_value(pos, ['End', 'year'], '')), str(get_nested_value(pos, ['End', 'month'], '')), str(get_nested_value(pos, ['End', 'day'], ''))
                    end_date_str = '-'.join(filter(None, [e_year, e_month, e_day]))
                    officer_record[f'Position_EndDate_{i}'] = end_date_str if end_date_str else 'Present'

                # Corporate Affiliations
                affiliations = get_nested_value(officer, ['CorporateAffiliations'], [])
                if isinstance(affiliations, dict): affiliations = [affiliations]
                for i, aff in enumerate(affiliations, start=1):
                    officer_record[f'Affiliation_CompanyName_{i}'] = get_nested_value(aff, ['Company', 'name'])
                    officer_record[f'Affiliation_CompanyOrgID_{i}'] = get_nested_value(aff, ['Company', 'orgid'])
                    officer_record[f'Affiliation_Title_{i}'] = get_nested_value(aff, ['Officer', 'title'])
                    officer_record[f'Affiliation_ActiveStatus_{i}'] = get_nested_value(aff, ['Officer', 'active'])

                # Salary
                salaries = get_nested_value(officer, ['SalaryInformation', 'CompensationPeriod'], [])
                if isinstance(salaries, dict): salaries = [salaries]
                for i, sal in enumerate(salaries, start=1):
                    officer_record[f'Salary_Year_{i}'] = get_nested_value(sal, ['Submission', 'year'])
                    std_comp = sal.get('StandardizedCompensation', [])
                    if isinstance(std_comp, list):
                        # Robustly find FYT and RSA regardless of position
                        officer_record[f'Salary_FYT_{i}'] = next((item.get('_') for item in std_comp if item.get('coa') == 'FYT'), None)
                        officer_record[f'Salary_RSA_{i}'] = next((item.get('_') for item in std_comp if item.get('coa') == 'RSA'), None)
                
                # Education History (NEW)
                educations = get_nested_value(officer, ['PersonInformation', 'EducationHistory'], [])
                if isinstance(educations, dict): educations = [educations]
                for i, edu in enumerate(educations, start=1):
                    officer_record[f'Education_College_{i}'] = get_nested_value(edu, ['College', '_'])
                    officer_record[f'Education_Degree_{i}'] = get_nested_value(edu, ['Degree', '_'])
                    officer_record[f'Education_Major_{i}'] = get_nested_value(edu, ['Major', '_'])
                    officer_record[f'Education_GradYear_{i}'] = get_nested_value(edu, ['Graduation', 'year'])

                # Committee Information (NEW)
                committees = get_nested_value(officer, ['PositionInformation', 'CommitteeMemberships'], [])
                if isinstance(committees, dict): committees = [committees]
                for i, com in enumerate(committees, start=1):
                    officer_record[f'Committee_Name_{i}'] = com.get('CommitteeName')
                    officer_record[f'Committee_Title_{i}'] = com.get('Title')
                    officer_record[f'Committee_StartDate_{i}'] = get_nested_value(com, ['Start', 'year'])
                
                # Append the fully processed officer record to the main list
                all_officers_data.append(officer_record)
            
            processed_files += 1

    # --- 3. Define Column Order and Create DataFrame ---
    if not all_officers_data:
        print("No officer data was processed. The output file will not be created.")
        return
        
    # Define the static, non-looping columns in their desired order
    static_columns = [
        'CompanyName', 'CompanyOrgID', 'SourceFile', 'OfficerID', 'PersonID',
        'Status', 'Prefix', 'FirstName', 'LastName', 'Initial/Middle', 'Suffix',
        'Age', 'Sex', 'Biography'
    ]
    
    # Find the maximum number of items for each variable category
    max_positions = get_max_index(all_officers_data, 'Position_LongTitle_')
    max_affiliations = get_max_index(all_officers_data, 'Affiliation_CompanyName_')
    max_salaries = get_max_index(all_officers_data, 'Salary_Year_')
    max_educations = get_max_index(all_officers_data, 'Education_College_')
    max_committees = get_max_index(all_officers_data, 'Committee_Name_')
    
    # Build the dynamic column lists
    position_cols = [
        item for i in range(1, max_positions + 1)
        for item in [
            f'Position_LongTitle_{i}', f'Position_StartDate_{i}', f'Position_EndDate_{i}'
        ]
    ]
    
    affiliation_cols = [
        item for i in range(1, max_affiliations + 1)
        for item in [
            f'Affiliation_CompanyName_{i}', f'Affiliation_CompanyOrgID_{i}',
            f'Affiliation_Title_{i}', f'Affiliation_ActiveStatus_{i}'
        ]
    ]
    
    salary_cols = [
        item for i in range(1, max_salaries + 1)
        for item in [
            f'Salary_Year_{i}', f'Salary_FYT_{i}', f'Salary_RSA_{i}'
        ]
    ]
    
    education_cols = [
        item for i in range(1, max_educations + 1)
        for item in [
            f'Education_College_{i}', f'Education_Degree_{i}',
            f'Education_Major_{i}', f'Education_GradYear_{i}'
        ]
    ]
    
    committee_cols = [
        item for i in range(1, max_committees + 1)
        for item in [
            f'Committee_Name_{i}', f'Committee_Title_{i}', f'Committee_StartDate_{i}'
        ]
    ]
    
    # Combine all column lists into the final, guaranteed order
    final_column_order = (
        static_columns + 
        position_cols + 
        affiliation_cols + 
        salary_cols + 
        education_cols + 
        committee_cols
    )
    
    # Create DataFrame using the specified column order
    # Any missing keys for a given officer will be filled with 'NaN'
    df = pd.DataFrame(all_officers_data, columns=final_column_order)
    
    # --- 4. Save to Excel ---
    try:
        # df.to_excel(output_filename, index=False, engine='openpyxl')
        df.to_csv(output_filename, index=False, encoding='utf-8-sig') 
        print(f"\n✅ Success! Data for {len(all_officers_data)} officers from {processed_files} files has been saved to '{output_filename}'")
        print(f"Column order has been enforced and categories are grouped.")
    except Exception as e:
        print(f"\n❌ Error: Could not save the Excel file. Reason: {e}")


# --- HOW TO USE ---
if __name__ == "__main__":
    # 1. Create a folder named 'json_files' (or any name you like).
    # 2. Place all your JSON files inside this folder.
    # 3. Make sure this script is in the same parent directory as your folder.
    JSON_FOLDER_PATH = 'JSON-DATA-ALL'
    # OUTPUT_EXCEL_FILE = 'combined_officers_output_ordered.xlsx'
    OUTPUT_FILE = 'combined_officers_output_ordered.csv'
    
    # process_json_folder_to_excel(JSON_FOLDER_PATH, OUTPUT_EXCEL_FILE)
    process_json_folder_to_excel(JSON_FOLDER_PATH, OUTPUT_FILE)
