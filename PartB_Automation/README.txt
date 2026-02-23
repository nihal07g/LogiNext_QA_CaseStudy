Google Maps Route Extraction (Selenium)

Overview:
Automates a Google Maps driving route lookup, captures the step-by-step driving
instructions, saves them to an Excel file, and stores a full-page screenshot
of the route page.

Technologies Used:
- Python 3.x
- Selenium (Chrome WebDriver)
- openpyxl (Excel export)

Execution Steps:
1) Install dependencies:
   python -m pip install selenium openpyxl

2) Ensure ChromeDriver matches your Chrome version and is in PATH or placed
   beside the script.

3) Run the script from the PartB_Automation folder:
   python maps_route_instructions.py

Assumptions:
- Internet connectivity is available.
- Google Maps UI is accessible without manual interaction.
- Chrome is installed and ChromeDriver is compatible with the installed version.

Expected Outputs:
- driving_instructions.xlsx (in the same folder as the script)
- route_screenshot.png (in the same folder as the script)
