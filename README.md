# JSON Scrapper Project

A comprehensive Python project for scraping corporate officer and director data from Refinitiv's OfficersDirectors API and converting it to structured formats (Excel/CSV).

## Project Overview

This project automates the process of:
1. **Data Collection**: Fetching corporate officer data from Refinitiv's API using PermIDs
2. **Data Processing**: Converting nested JSON responses into structured tabular data
3. **Data Export**: Saving processed data to Excel and CSV formats
4. **Data Recovery**: Re-fetching failed or missing data entries

## Project Structure

### Core Python Files

#### 1. `request.py` - Main Data Fetching Script
- **Purpose**: Primary script for fetching officer data from Refinitiv API
- **Key Features**:
  - Reads PermIDs from Excel files (`main.xlsx`)
  - Makes authenticated API requests to Refinitiv's OfficersDirectors endpoint
  - Handles authentication with cookies and headers
  - Saves responses as JSON files with company names and IDs
  - Supports resuming from specific row numbers
  - Handles both successful JSON responses and error cases

**Usage**:
```bash
python request.py --start-row 4162
```

#### 2. `api_client.py` - API Client Module
- **Purpose**: Isolated API client for making requests to Refinitiv's OfficersDirectors API
- **Key Features**:
  - Clean separation of API logic from main processing
  - Configurable timeout settings
  - Returns raw HTTP responses for flexible handling

#### 3. `find_and_fetch_remaining.py` - Data Recovery Script
- **Purpose**: Identifies and re-fetches failed data entries
- **Key Features**:
  - Scans for files with "unknown-" prefix (failed API calls)
  - Extracts PermIDs from failed filenames
  - Re-attempts API calls for missing data
  - Saves recovered data to separate directory

**Usage**:
```bash
python find_and_fetch_remaining.py
```

#### 4. `gemini-script.py` - Advanced Data Processing Script
- **Purpose**: Comprehensive JSON-to-CSV conversion with advanced data structuring
- **Key Features**:
  - Processes nested JSON officer data into flat tabular format
  - Handles variable-length data (multiple positions, affiliations, salaries, etc.)
  - Enforces consistent column ordering across all records
  - Supports dynamic column generation based on data complexity
  - Exports to CSV format with UTF-8 encoding

**Data Fields Extracted**:
- **Basic Information**: Company name, Org ID, Officer ID, Person ID, Status
- **Personal Details**: Name components, Age, Sex, Biography
- **Position History**: Job titles, start/end dates (multiple positions)
- **Corporate Affiliations**: Other company roles and titles
- **Compensation**: Salary data (FYT, RSA) by year
- **Education**: College, degree, major, graduation year
- **Committee Memberships**: Committee names, titles, start dates

**Usage**:
```bash
python gemini-script.py
```

#### 5. `jsonToExcel.py` - Alternative Data Processing Script
- **Purpose**: Alternative JSON processing with Excel output
- **Key Features**:
  - Similar data extraction to gemini-script.py
  - Outputs to Excel format instead of CSV
  - Includes progress tracking and error handling
  - Organizes columns by static vs dynamic fields

**Usage**:
```bash
python jsonToExcel.py
```

### Data Directories

#### 1. `JSON-DATA-ALL/` - Main Data Storage
- **Purpose**: Primary storage for successfully fetched JSON files
- **Content**: Contains thousands of JSON files with officer data
- **Naming Convention**: `{Company-Name}-{OrgID}.json`
- **Example**: `Apple-Inc-123456789.json`

#### 2. `first-results/` - Initial Data Collection
- **Purpose**: Contains the first batch of successfully fetched data
- **Content**: Sample of processed JSON files
- **Status**: Historical data from initial scraping runs

#### 3. `remaining-json/` - Recovery Data Storage
- **Purpose**: Storage for re-fetched data from failed API calls
- **Content**: JSON files recovered using `find_and_fetch_remaining.py`
- **Naming Convention**: Same as main directory

### Virtual Environment

#### `venv/` - Python Virtual Environment
- **Purpose**: Isolated Python environment with project dependencies
- **Key Packages**:
  - `requests`: HTTP client for API calls
  - `pandas`: Data manipulation and analysis
  - `openpyxl`: Excel file reading/writing
  - `numpy`: Numerical computing support

## Data Flow

```
Excel File (main.xlsx)
    ↓
request.py (API Fetching)
    ↓
JSON-DATA-ALL/ (Raw Data Storage)
    ↓
gemini-script.py OR jsonToExcel.py (Data Processing)
    ↓
CSV/Excel Output Files
```

## Authentication

The project uses Refinitiv's authentication system with:
- **Cookies**: Session tokens and authentication cookies
- **Headers**: User agent and API-specific headers
- **Tokens**: STS tokens for API access

**Note**: Authentication tokens expire and need periodic updates.

## Key Features

### 1. Robust Error Handling
- Graceful handling of API failures
- Retry mechanisms for failed requests
- Fallback naming for unknown companies

### 2. Data Normalization
- Consistent column ordering across all records
- Dynamic column generation for variable-length data
- Safe navigation through nested JSON structures

### 3. Scalable Processing
- Handles large datasets (thousands of companies)
- Progress tracking for long-running operations
- Memory-efficient processing

### 4. Flexible Output Formats
- CSV output with UTF-8 encoding
- Excel output with proper formatting
- Structured column organization

## Usage Instructions

### 1. Initial Setup
```bash
# Activate virtual environment
source venv/bin/activate

# Install dependencies (if needed)
pip install requests pandas openpyxl numpy
```

### 2. Data Fetching
```bash
# Fetch data from Excel file
python request.py

# Resume from specific row
python request.py --start-row 4162
```

### 3. Data Recovery
```bash
# Re-fetch failed entries
python find_and_fetch_remaining.py
```

### 4. Data Processing
```bash
# Convert to CSV (recommended)
python gemini-script.py

# Convert to Excel
python jsonToExcel.py
```

## Output Files

### CSV Output (`combined_officers_output_ordered.csv`)
- **Format**: UTF-8 encoded CSV
- **Structure**: Flat tabular format with consistent columns
- **Content**: All officer data from processed JSON files

### Excel Output (`officers_data.xlsx`)
- **Format**: Excel workbook
- **Structure**: Organized columns with static and dynamic fields
- **Content**: Same data as CSV but in Excel format

## Data Schema

The processed data includes the following field categories:

### Static Fields
- CompanyName, CompanyOrgID, SourceFile
- OfficerID, PersonID, Status
- Prefix, FirstName, LastName, Initial/Middle, Suffix
- Age, Sex, Biography

### Dynamic Fields (Multiple Instances)
- **Positions**: LongTitle_N, StartDate_N, EndDate_N
- **Affiliations**: AffiliationCompanyName_N, AffiliationTitle_N, etc.
- **Salaries**: SalaryYear_N, FYT_N, RSA_N
- **Education**: College_N, Degree_N, Major_N, Graduation_N
- **Committees**: CommitteeName_N, CommitteeTitle_N, CommitteeStartDate_N

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Update cookies and tokens in `request.py`
2. **API Rate Limits**: Add delays between requests
3. **Memory Issues**: Process data in smaller batches
4. **File Encoding**: Ensure UTF-8 encoding for international characters

### Error Recovery

1. **Failed API Calls**: Use `find_and_fetch_remaining.py` to retry
2. **Corrupted JSON**: Check file integrity and re-fetch if needed
3. **Missing Data**: Verify PermIDs in source Excel file

## Dependencies

- Python 3.10+
- requests >= 2.32.5
- pandas >= 2.3.3
- openpyxl >= 3.1.5
- numpy >= 2.2.6

## License

This project is for educational and research purposes. Please ensure compliance with Refinitiv's terms of service when using their API.

## Contributing

When modifying the project:
1. Test with small datasets first
2. Maintain backward compatibility
3. Update documentation for new features
4. Follow existing code style and structure
