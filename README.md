# BatchXMLPrepperMacro
Preparatory process for batch XML uploading to Vernon to flatten System IDs with multiple rows - Macro VBA version

Overview

This VBA macro transforms a long/RAW input table of file references into a wide output table suitable for use in Vernon CMS or other batch-processing workflows.

It takes rows in the format:

System_ID	URL	Category
12345	https://example.com/img1.jpg
	Photo
12345	https://example.com/img2.jpg
	Photo
67890	https://example.com/img3.jpg
	Poster

And produces a wide table like:

SystemID	Category1	ExternalFileField1	Category2	ExternalFileField2
id\12345	Photo	https://example.com/img1.jpg
	Photo	https://example.com/img2.jpg

id\67890	Poster	https://example.com/img3.jpg
		
Features

Cleans and validates input (System_ID, URL, optional Category)

Drops blanks and "nan" values

Assigns slot numbers per System_ID in input order

Produces CategoryN + ExternalFileFieldN pairs

Adds SystemID prefix (id\...)

Outputs to a new sheet named Output, formatted as an Excel Table

Repository Structure
your-repo/
│
├─ README.md                   # This file
├─ src/
│   └─ BuildWideFromLong.bas   # VBA macro code (exported module)
├─ example/
│   └─ sample_input.xlsx       # Example input data (optional)
└─ workbook/
    └─ macro_workbook.xlsm     # Example macro-enabled workbook (optional)

Installation

Open your workbook in Excel.

Press Alt + F11 to open the VBA editor.

Go to File → Import File… and select BuildWideFromLong.bas.

Save the workbook as .xlsm (macro-enabled).

Usage

Place your input data on a sheet named Data with headers:

System_ID (required)

URL (required)

Category (optional)

Run the macro:

Press Alt + F8, select BuildWideFromLong, click Run.

The wide output will appear on a sheet named Output.

Notes

Input must be in long format (one row per file).

The macro will overwrite any existing Output sheet.

To keep sensitive data out of GitHub, use .gitignore to exclude working .xlsm files.

Contributing

Fork the repository and submit pull requests with improvements.

If you add new macros, export them as .bas files and place them in the /src/ folder.
