import os
import sys
from datetime import datetime

from .config import OUTPUT_DIR, CV_SOURCE_PATH
from .document_service import save_raw_text, save_to_word, save_to_pdf
from .file_service import (
    clear_console,
    copy_to_clipboard,
    copy_cv_to_folder,
    archive_and_reset_cover_letter,
    build_safe_base_name,
)
from .excel_service import close_excel_tracker_if_open, append_to_excel, get_portal_name
from .prompts import (
    collect_job_metadata,
    confirm_metadata,
    collect_excel_fields,
    confirm_excel_fields,
    edit_excel_field,
)


def main():
    clear_console()
    close_excel_tracker_if_open()

    print("=" * 80)
    print("JOB DESCRIPTION SAVER")
    print("=" * 80)
    print()

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("Commands available while answering metadata:")
    print("  :back    -> redo previous answer")
    print("  :restart -> restart all metadata fields")
    print("  :cls     -> clear console")
    print("  :quit    -> exit program")
    print()

    while True:
        company_name, company_website, job_title = collect_job_metadata()
        if confirm_metadata(company_name, company_website, job_title):
            break
        print("\nLet's enter them again.\n")

    print()
    print("Paste the job description below.")
    print("When finished, type 'END' on a new line and press Enter")
    print("-" * 80)

    try:
        lines = []
        while True:
            try:
                line = input()
                if line.strip().upper() == 'END':
                    break
                lines.append(line)
            except EOFError:
                break
        job_text = '\n'.join(lines)
    except KeyboardInterrupt:
        print("\n\nOperation cancelled.")
        sys.exit(0)

    if not job_text.strip():
        print("\nError: No text provided!")
        sys.exit(1)

    print()
    print("-" * 80)

    base_filename = build_safe_base_name(OUTPUT_DIR, company_name, job_title)

    job_folder = os.path.join(OUTPUT_DIR, base_filename)
    os.makedirs(job_folder, exist_ok=True)

    copied_cv_path = copy_cv_to_folder(CV_SOURCE_PATH, job_folder)

    if copied_cv_path:
        print(f"Copied CV: {copied_cv_path}")
        print("")

    txt_filename = os.path.join(job_folder, f"{base_filename}.txt")
    try:
        save_raw_text(job_text, txt_filename)
    except Exception as e:
        print(f"✗ Error saving text archive: {e}")

    docx_filename = os.path.join(job_folder, f"temp_{base_filename}.docx")
    pdf_filename = os.path.join(job_folder, f"{base_filename}.pdf")

    try:
        save_to_word(job_text, docx_filename, job_title, company_name, company_website)
    except Exception as e:
        print(f"✗ Error creating document: {e}")
        sys.exit(1)

    pdf_success = save_to_pdf(docx_filename, pdf_filename)

    final_created_file = ""

    if pdf_success:
        final_created_file = pdf_filename
        try:
            os.remove(docx_filename)
        except Exception:
            pass
    else:
        final_created_file = docx_filename
        print(f"⚠ Temporary Word file kept: {docx_filename}")
        print("  You can manually convert it to PDF")

    copy_to_clipboard(final_created_file)

    print("\n--- Cover Letter Archive & Reset ---")
    moved_cover_letter_path, recreated_cover_letter_path = archive_and_reset_cover_letter(job_folder)

    print("\n--- Update Excel Tracker ---")
    excel_data = collect_excel_fields()

    while True:
        action = confirm_excel_fields(excel_data)

        if action == "confirm":
            break

        if action == "restart":
            print("\nLet's enter the Excel tracker fields again.\n")
            excel_data = collect_excel_fields()
            continue

        if isinstance(action, int):
            result = edit_excel_field(excel_data, action)

            if result == "restart":
                print("\nRestarting all Excel tracker fields.\n")
                excel_data = collect_excel_fields()
            else:
                excel_data = result

    row_data = [
        job_title,
        excel_data["job_id"],
        datetime.now().strftime("%Y-%m-%d"),
        "",
        "Yes",
        get_portal_name(excel_data["portal_choice"]),
        excel_data["job_link"],
        company_name,
        excel_data["job_type"],
        excel_data["industry"],
        excel_data["location"],
        excel_data["status"],
        excel_data["docs_sent"],
        excel_data["todo_done"],
        excel_data["hr_name"],
        excel_data["hr_contact"],
        excel_data["comments"],
        excel_data["id_pw"],
        job_folder
    ]

    if append_to_excel(row_data):
        print("✓ Excel tracker updated successfully.")

    print()
    print("=" * 80)
    print("DONE!")
    print(f"Job folder: {job_folder}")
    print("")
    print(f"Main output file: {final_created_file}")
    print("")

    if moved_cover_letter_path:
        print(f"Moved cover letter: {moved_cover_letter_path}")
        print("")

    if recreated_cover_letter_path:
        print(f"Recreated working cover letter: {recreated_cover_letter_path}")
        print("")

    print(f"Text archive file: {txt_filename}")
    print("=" * 80)
