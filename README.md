# Excel File Creator using Python and XlsxWriter

This Python script creates a structured Excel file using the `xlsxwriter` library. It populates a worksheet with detailed part information such as assembly components, revision numbers, material specifications, and approvals. This is useful for managing and exporting Bill of Materials (BoM), design specifications, or manufacturing records in a clean and professional Excel format.

## 🚀 Features

* Creates a new Excel workbook and worksheet
* Writes custom headers and formats for documentation
* Supports structured entry of assembly part details
* Handles optional/missing fields gracefully
* Easy to customize or extend with more fields

## 📦 Dependencies

* Python 3.x
* [xlsxwriter](https://pypi.org/project/XlsxWriter/)

Install the required package using pip:

```bash
pip install xlsxwriter
```

## 📂 How to Use

1. Clone or download the script.
2. Run the script:

```bash
python create_excel.py
```

3. An Excel file named `Creating Excel Task.xlsx` will be generated in the current directory with a worksheet named "Sheet 1".

## 📊 Output Preview

| S.No | Assembly Parts | Standard Part | Revision Number | Status       | Part Number | Weight | Material | Quantity | Notes                             | Date       | Designed By | Detailed | Approved By | Custom Scale | Custom Paper Size | Orientation |
| ---- | -------------- | ------------- | --------------- | ------------ | ----------- | ------ | -------- | -------- | --------------------------------- | ---------- | ----------- | -------- | ----------- | ------------ | ----------------- | ----------- |
| 1    | Lever SA       | No            | 2               | Not Released | L1254       | 120    | SS316    | 1        | Remove Burns, Chamfer sharp edges | 22/10/2023 | James       | John     | Michael     | 0.5          | A1                |             |
| ...  |                |               |                 |              |             |        |          |          |                                   |            |             |          |             |              |                   |             |

## 📘 Notes

* Empty fields in the entry dictionary will appear as blank cells.
* The column headers can be adjusted to suit your specific document structure.
* Useful for documentation in engineering, manufacturing, or design workflows.

