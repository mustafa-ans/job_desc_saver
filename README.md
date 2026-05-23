# Job Description Saver

A local Python CLI tool for saving job descriptions into job-specific folders, copying the right CV, archiving cover letters, and updating a local Excel application tracker.

## What it does

This tool helps organize job applications by creating one folder per job and storing the related materials in one place.

Main features:
- Saves a pasted job description as a formatted document.
- Converts the document to PDF when PDF conversion is available.
- Saves the raw pasted job description as a `.txt` archive.
- Copies a selected CV into the job-specific folder.
- Moves the current working cover letter into the same job folder.
- Recreates a fresh working cover letter from a master template.
- Updates a local Excel tracker with job metadata.
- Copies the final created file path to the clipboard.

## Project structure

Run the project from the repository root.

```text
job_desc_saver/
├── app/
│   ├── __init__.py
│   ├── __main__.py
│   ├── config.py
│   ├── document_service.py
│   ├── excel_service.py
│   ├── file_service.py
│   ├── prompts.py
│   └── workflow.py
├── .env
├── .env.example
├── .gitignore
├── requirements.txt
├── run.py
└── README.md
```

## Required local files

Before using the tool, make sure these files already exist on your machine:

- Excel tracker workbook.
- At least one CV PDF.
- A working cover letter `.docx`.
- Optionally a working cover letter PDF.
- A master cover letter `.docx` used to recreate the working copy.

These files can live anywhere on your machine, but their exact paths must be added to the `.env` file.

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/<your-username>/job_desc_saver.git
cd job_desc_saver
```

### 2. Create and activate a virtual environment

#### Windows PowerShell

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

If `requirements.txt` does not yet include everything, these are the important packages:

```bash
pip install python-docx openpyxl pyperclip python-dotenv docx2pdf
```

## `.env` setup

Create a file named `.env` in the project root.

Important:
- The `.env` file must be in the root of the project, next to `run.py`.
- Do not commit `.env` to GitHub.
- Commit only `.env.example`.

### Example `.env`

```env
OUTPUT_DIR=C:\Users\YOUR_NAME\OneDrive\Employment\job_desc
EXCEL_PATH=C:\Users\YOUR_NAME\OneDrive\Employment\Job List Mustafa.xlsx

COVER_LETTER_SOURCE_PATH=C:\Users\YOUR_NAME\OneDrive\Employment\Cover_Letter_Mustafa_Ansari.docx
COVER_LETTER_SOURCE_PATH_PDF=C:\Users\YOUR_NAME\OneDrive\Employment\Cover_Letter_Mustafa_Ansari.pdf
COVER_LETTER_MASTER_PATH=C:\Users\YOUR_NAME\OneDrive\Employment\Cover_Letter_Mustafa_Ansari Master.docx

CV_BACKEND_PATH=C:\Users\YOUR_NAME\OneDrive\Employment\1. Backend Engineer CV\CV_Mustafa_Ansari.pdf
CV_PYTHON_AI_PATH=C:\Users\YOUR_NAME\OneDrive\Employment\2. Python AI Dev CV\CV_Mustafa_Ansari.pdf
```

## Where files should be pasted or placed

This tool does not require you to paste files into the repository.

Instead:

- Put your real local file paths into `.env`.
- Keep the actual Excel, CV, and cover letter files in your preferred folders on your computer.
- Run the tool from the project root.
- Paste the **job description text into the terminal** when prompted.

The only file you manually create in the repository is `.env`.

## How to run

From the project root:

```bash
python run.py
```

or:

```bash
python -m app
```

## How to use it

When you start the program:

1. It asks for company name, company website, and job title.
2. It asks you to paste the job description text into the terminal.
3. It creates a job-specific output folder under `OUTPUT_DIR`.
4. It asks which CV variant to copy.
5. It saves:
   - the raw text,
   - the formatted document,
   - the PDF if conversion succeeds.
6. It moves the current working cover letter into the job folder.
7. It recreates a fresh working cover letter from the master file.
8. It asks for Excel tracker fields.
9. It appends a new row to the Excel tracker.
10. It copies the final created file path to your clipboard.

## Example output folder

If the company is `Example GmbH` and the role is `Backend Engineer`, the tool may create:

```text
C:\Users\YOUR_NAME\OneDrive\Employment\job_desc\Example_GmbH_Backend_Engineer\
├── Example_GmbH_Backend_Engineer.txt
├── Example_GmbH_Backend_Engineer.pdf
├── CV_Mustafa_Ansari.pdf
├── Cover_Letter_Mustafa_Ansari.docx
└── Cover_Letter_Mustafa_Ansari.pdf
```

If PDF conversion fails, the temporary `.docx` may remain instead of the final PDF.

## Commands supported in prompts

Some prompts support special commands:

- `:back` → go back to the previous question.
- `:restart` → restart the current input section.
- `:cls` → clear the console and restart that section.
- `:quit` → exit the program.

## Important cautions

### 1. Do not commit `.env`

Your `.env` contains private local paths and should remain ignored by Git.

Your `.gitignore` should include:

```gitignore
.env
__pycache__/
*.pyc
venv/
```

### 2. PDF conversion may require Microsoft Word

`docx2pdf` typically requires Microsoft Word on Windows for DOCX-to-PDF conversion.

If conversion fails:
- the script keeps the generated `.docx`,
- and you can convert it manually.

### 3. Excel workbook should not be open

The script tries to close the Excel tracker if it is open, but it is safer to keep the workbook closed before running the tool.

### 4. OneDrive can sometimes lock files

If your files are inside OneDrive, syncing may occasionally interfere with file moves, Excel access, or Git operations.

### 5. Paths must be valid

If any path in `.env` is missing or wrong, the related step may fail:
- CV copy,
- cover letter move/reset,
- Excel update,
- output creation.

## Extending CV variants

The current design can easily support more CV versions.

### Current idea

Your `config.py` likely has a `CV_OPTIONS` structure similar to this:

```python
CV_OPTIONS = {
    "1": {
        "label": "Backend Engineer CV",
        "path": os.getenv("CV_BACKEND_PATH", ""),
    },
    "2": {
        "label": "Python AI Dev CV",
        "path": os.getenv("CV_PYTHON_AI_PATH", ""),
    },
}
```

### To add another CV

#### 1. Add a new variable to `.env`

Example:

```env
CV_DATA_ENGINEER_PATH=C:\Users\YOUR_NAME\OneDrive\Employment\3. Data Engineer CV\CV_Mustafa_Ansari.pdf
```

#### 2. Add it to `.env.example`

```env
CV_DATA_ENGINEER_PATH=
```

#### 3. Add a new option in `app/config.py`

```python
CV_OPTIONS = {
    "1": {
        "label": "Backend Engineer CV",
        "path": os.getenv("CV_BACKEND_PATH", ""),
    },
    "2": {
        "label": "Python AI Dev CV",
        "path": os.getenv("CV_PYTHON_AI_PATH", ""),
    },
    "3": {
        "label": "Data Engineer CV",
        "path": os.getenv("CV_DATA_ENGINEER_PATH", ""),
    },
}
```

That is enough for the menu-based selection flow if the prompt already loops through `CV_OPTIONS`.

## Extending cover letter variants

You can support multiple cover-letter templates in the same way as CVs.

### Recommended pattern

Instead of only one working/master pair, define a cover-letter options mapping:

```python
COVER_LETTER_OPTIONS = {
    "1": {
        "label": "Backend Cover Letter",
        "working_docx": os.getenv("CL_BACKEND_WORKING_DOCX", ""),
        "working_pdf": os.getenv("CL_BACKEND_WORKING_PDF", ""),
        "master_docx": os.getenv("CL_BACKEND_MASTER_DOCX", ""),
    },
    "2": {
        "label": "Python AI Cover Letter",
        "working_docx": os.getenv("CL_PYTHON_AI_WORKING_DOCX", ""),
        "working_pdf": os.getenv("CL_PYTHON_AI_WORKING_PDF", ""),
        "master_docx": os.getenv("CL_PYTHON_AI_MASTER_DOCX", ""),
    },
}
```

Then:
- prompt the user to choose a cover-letter variant,
- pass the selected paths into the archive/reset function,
- move and recreate the correct working version.

### Example `.env` additions

```env
CL_BACKEND_WORKING_DOCX=C:\Users\YOUR_NAME\OneDrive\Employment\Backend_CL\Cover_Letter_Mustafa_Ansari.docx
CL_BACKEND_WORKING_PDF=C:\Users\YOUR_NAME\OneDrive\Employment\Backend_CL\Cover_Letter_Mustafa_Ansari.pdf
CL_BACKEND_MASTER_DOCX=C:\Users\YOUR_NAME\OneDrive\Employment\Backend_CL\Cover_Letter_Mustafa_Ansari Master.docx

CL_PYTHON_AI_WORKING_DOCX=C:\Users\YOUR_NAME\OneDrive\Employment\Python_AI_CL\Cover_Letter_Mustafa_Ansari.docx
CL_PYTHON_AI_WORKING_PDF=C:\Users\YOUR_NAME\OneDrive\Employment\Python_AI_CL\Cover_Letter_Mustafa_Ansari.pdf
CL_PYTHON_AI_MASTER_DOCX=C:\Users\YOUR_NAME\OneDrive\Employment\Python_AI_CL\Cover_Letter_Mustafa_Ansari Master.docx
```

## Recommended future improvements

Possible future extensions:
- Auto-suggest CV variant based on job title keywords.
- Auto-suggest cover letter variant based on role category.
- Save the selected CV/cover-letter variant into the Excel tracker.
- Add validation for missing `.env` variables at startup.
- Add logging to a local file.
- Add unit tests for path validation and filename generation.
- Add a `settings validator` command that checks whether all required files exist before running the main flow.

## Troubleshooting

### `Could not copy to clipboard`
Install or verify `pyperclip`:

```bash
pip install pyperclip
```

### PDF conversion failed
Check:
- Microsoft Word is installed.
- `docx2pdf` is installed.
- the DOCX file is not open.

### Excel update failed
Check:
- the workbook path in `.env`,
- the workbook is not locked,
- `openpyxl` is installed.

### CV or cover letter file not found
Check the file paths in `.env` carefully. A single wrong folder name or missing file will cause the copy/move step to fail.

## Public repository note

This project is designed to be safe for a public GitHub repository only if:
- all private machine-specific paths are stored in `.env`,
- `.env` is ignored,
- no personal local paths remain in Git history,
- generated files and caches are ignored.

## License
MIT