# Vacation Rental Details Page Automation Testing



## Project Overview

  Automate testing of a vacation rental details page to ensure SEO compliance and functionality. This includes validating H1 tags, HTML tag sequence, image alt attributes, URL availability, currency filter accuracy, and script data integrity. Test results are consolidated into an Excel report for quick issue identification.
---

## Features
- Automated browser interactions using **Selenium WebDriver**.
- Currency filter testing with multiple currency options.
- Dynamic handling of dropdowns and web elements.
- Generates formatted Excel reports for results and summaries using **OpenPyXL**.
- Logging for debugging and monitoring test progress.

---

## Technologies Used
- **Python**: Selenium, Pandas, OpenPyXL, Logging
- **Automation Tool**: Selenium WebDriver with WebDriver Manager
- **Excel Handling**: OpenPyXL for creating and formatting Excel files
- **Frontend Interaction**: HTML and JavaScript elements using Selenium

---

## Prerequisites
1. Python 3.7+
2. Google Chrome browser installed
3. Required Python libraries:
   - `selenium`
   - `pandas`
   - `openpyxl`
   - `webdriver-manager`

Install dependencies using:
```bash
pip install selenium pandas openpyxl webdriver-manager
```

---

## How to Run
## Git Clone Instructions

To clone this project to your local machine, follow these steps:

1. **Open terminal (Command Prompt, PowerShell, or Terminal)**

2. **Clone the repository**:
   
       git clone https://github.com/M-E-U-E/Assignment_7.git or git clone git@github.com:M-E-U-E/Assignment_7.git
   
    Go to the Directory:
    ```
    cd Assignment_7
    ```
    Create a virtual environment to isolate dependencies:
    ```
    python3 -m venv env
    source env/bin/activate  # Linux/Mac
    env\Scripts\activate
    ```
   Install the requirements
    ```
    pip install -r requirements.txt
    ```
    **Run all tests using the main script:**
    ```
    python run_all_test.py
    ```

    **Run individual test scripts:**
    
    Currency Filtering Test:
    ```
    python Currency_Filtering_Test.py
    ```
    H1 Tag Existence Test:
    
    ```
    python H1_Tag_Existence_Test.py
    ```
    HTML Tag Sequence Test:
    
    ```
    python HTML_Tag_Sequence_Test.py
    ```
    Image Alt Attribute Test:
    
    ```
    python Image_Alt_Attribute_Test.py
    ```
    URL Status Code Test:
    
    ```
    python URL_Status_Code_Test.py
    ```
    Scrape Data from Script Tag Test:
    ```
    python Scrape_Data_from_Script_Tag.py
    ```

4. View the generated Excel reports in the `test_results` directory.

---
Generated Files Include:

- `currency_test_summary.xlsx`
- `h1_tag_results.xlsx`
- `h1_tag_summary.xlsx`
- `html_tag_results.xlsx`
- `image_alt_results.xlsx`
- `image_alt_summary.xlsx`
- `report_model_details.xlsx`
- `report_model.xlsx`
- `script_data_summary.xlsx`
- `script_tag_results.xlsx`
- `url_status_results.xlsx`
- `url_status_summary.xlsx`
---
**View**
  
After running all the code, the main report files are report_model_details.xlsx and report_model.xlsx. However, you can also check the individual files to review the detailed results for each aspect of the testing.

## File Structure
```
Assignment_7/
├── env/
├── test_results/
├── .gitignore
├── Currency_Filtering_Test.py
├── H1_Tag_Existence_Test.py
├── HTML_Tag_Sequence_Test.py
├── Image_Alt_Attribute_Test.py
├── report_model.py
├── requirements.txt
├── run_all_test.py
├── Scrape_Data_from_Script_Tag.py
└── URL_Status_Code_Test.py

```

---

## Outputs
1. **Currency Test Results**:
   - File: `test_results/currency_test_results.xlsx`
   - Details of each tested currency, including status (Pass/Fail) and reasons for failures.

2. **Summary**:
   - File: `test_results/currency_test_summary.xlsx`
   - Overall summary of the test case with Pass/Fail status and comments.

---

## Configuration
- **URL**: Update the `url` variable in the `main()` function to point to the desired page.
- **Currencies**: Modify the `currency_list` dictionary to include the currencies and their symbols for testing.

---

## Security Best Practices
- Use virtual environments to isolate dependencies.
- Avoid hardcoding sensitive information (e.g., URLs, credentials).
- Regularly update dependencies to patch vulnerabilities.

---

## Known Issues
- Dropdown selectors and element locators may need adjustments for different websites.
- Timeout errors may occur if the page load time is too long. Adjust wait times as needed.

---

## Contributing
1. Fork the repository.
2. Create a new branch:
   ```bash
   git checkout -b feature-branch
   ```
3. Commit your changes:
   ```bash
   git commit -m "Add new feature"
   ```
4. Push to the branch:
   ```bash
   git push origin feature-branch
   ```
5. Open a Pull Request.



## Depencies
    pandas==<latest_version>         # For handling data and generating reports
    openpyxl==<latest_version>       # For working with Excel files
    requests==<latest_version>       # For making HTTP requests (URL status checks)
    bs4==<latest_version>            # BeautifulSoup for HTML parsing (HTML tag tests)
    lxml==<latest_version>           # XML and HTML parsing library
    pytest==<latest_version>         # For running test cases
    pytest-html==<latest_version>    # For generating HTML reports (optional)


