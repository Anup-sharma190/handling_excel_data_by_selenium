# handling_excel_data_by_selenium
# Excel Update Automation Demo

## Project Overview
This project demonstrates **Excel automation using Python**. The script reads, searches, and updates Excel files dynamically using the `openpyxl` library.  
It is designed to showcase practical Python skills, file handling, and automation logic in a clean, reusable, and recruiter-friendly format.

---

## Skills Demonstrated
- Python programming
- Excel file handling and automation (`openpyxl`)
- Searching and updating data dynamically
- Reusable and modular function design
- Error handling for safe updates
- Logging and confirmation messages
- GitHub-ready project structure

---

## Tools Used
- Python 3.x
- openpyxl library (`pip install openpyxl`)
- Any Excel file (.xlsx format)

---

## Workflow
1. **Load the Excel file** using `openpyxl`.
2. **Identify the target column** based on the header name.
3. **Search for a row** containing a specific value.
4. **Update the cell** with a new value.
5. **Save the Excel file** safely.
6. **Print logs** to confirm the update or report errors.

---

## File Description

- `excel_update_demo.py` – Main Python script demonstrating Excel automation.
    - Contains a reusable function: `update_excel_cell(file_path, search_value, column_name, new_value)`
    - Includes a demo example updating `"Apple"` price to `500`.

---

## How to Use

1. Clone the repository:
```bash
git clone <repository_url>
