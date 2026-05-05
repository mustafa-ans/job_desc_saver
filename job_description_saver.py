#!/usr/bin/env python3
"""
Job Description Saver
Saves pasted job descriptions to PDF and archives raw text
Supports:
  :back    -> go to previous question
  :restart -> restart metadata entry
  :quit    -> exit program
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os
import sys
import pyperclip

OUTPUT_DIR = r"REMOVED_LOCAL_PATH"


def clear_console():
    """Clear terminal screen once at script start"""
    os.system('cls' if os.name == 'nt' else 'clear')


def format_document(doc, job_text, job_title, company_name, company_website=""):
    """Add formatted content to the Word document"""

    title = doc.add_heading(job_title, 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    metadata = doc.add_paragraph()
    metadata.add_run(f"Company: {company_name}").bold = True

    if company_website:
        metadata.add_run("\n")
        run = metadata.add_run(f"Website: {company_website}")
        run.font.color.rgb = RGBColor(0, 102, 204)

    metadata.add_run(f"\nSaved: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}").italic = True

    doc.add_paragraph("_" * 80)

    paragraphs = job_text.strip().split('\n\n')

    for para_text in paragraphs:
        para_text = para_text.strip()
        if not para_text:
            continue

        is_heading = False
        heading_keywords = [
            'DEINE ROLLE', 'DEIN PROFIL', 'ABOUT THE JOB', 'YOUR ROLE',
            'YOUR PROFILE', 'REQUIREMENTS', 'RESPONSIBILITIES', 'QUALIFICATIONS',
            'BENEFITS', 'ABOUT US', 'WIR BIETEN', 'AUFGABEN', 'ANFORDERUNGEN',
            'WAS WIR BIETEN', 'DAS BIETEN WIR', 'UNSER VERSPRECHEN'
        ]

        for keyword in heading_keywords:
            if para_text.upper().startswith(keyword):
                is_heading = True
                break

        if is_heading:
            heading = doc.add_heading(para_text, level=2)
            heading.runs[0].font.color.rgb = RGBColor(0, 102, 204)
        else:
            p = doc.add_paragraph(para_text)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(6)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)

    return doc


def save_to_word(job_text, filename, job_title, company_name, company_website=""):
    """Save job description to Word document"""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    format_document(doc, job_text, job_title, company_name, company_website)
    doc.save(filename)
    return filename


def save_to_pdf(docx_filename, pdf_filename):
    """Convert Word document to PDF"""
    try:
        from docx2pdf import convert
        convert(docx_filename, pdf_filename)
        print(f"✓ PDF saved: {pdf_filename}")
        return True
    except ImportError:
        print("⚠ docx2pdf not available. Install with: pip install docx2pdf")
        print("  Note: Requires Microsoft Word on Windows/Mac")
        return False
    except Exception as e:
        print(f"⚠ PDF conversion failed: {e}")
        print("  You can manually open the temp Word file and save as PDF")
        return False


def save_raw_text(job_text, filename):
    """Save raw job description text as-is to archive"""
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(job_text)
    print(f"✓ Raw text archived: {filename}")


def copy_to_clipboard(text):
    """Copy plain text to clipboard"""
    try:
        pyperclip.copy(text)
        print(f"✓ Copied to clipboard: {text}")
    except Exception as e:
        print(f"⚠ Could not copy to clipboard: {e}")


def collect_job_metadata():
    """
    Collect company name, website, and job title with support for:
    :back    -> go to previous question
    :restart -> restart metadata entry
    :quit    -> exit program
    """
    fields = [
        ("company_name", "Enter company name (required): ", True),
        ("company_website", "Enter company website (optional, press Enter to skip): ", False),
        ("job_title", "Enter job title (required): ", True),
    ]

    answers = {
        "company_name": "",
        "company_website": "",
        "job_title": "",
    }

    index = 0

    while index < len(fields):
        field_key, prompt, required = fields[index]
        current_value = answers[field_key]

        if current_value:
            user_input = input(f"{prompt}[current: {current_value}] ").strip()
        else:
            user_input = input(prompt).strip()

        cmd = user_input.lower()

        if cmd == ":quit":
            print("Exiting...")
            sys.exit(0)

        if cmd == ":restart":
            print("Restarting metadata entry...\n")
            answers = {k: "" for k in answers}
            index = 0
            continue

        if cmd == ":back":
            if index == 0:
                print("Already at the first question.")
            else:
                index -= 1
                print("Going back to previous question.")
            continue

        if required and not user_input:
            print("This field is required!")
            continue

        answers[field_key] = user_input
        index += 1

    return answers["company_name"], answers["company_website"], answers["job_title"]


def confirm_metadata(company_name, company_website, job_title):
    """Show entered metadata and ask for confirmation"""
    while True:
        print("\nPlease confirm:")
        print(f"Company name   : {company_name}")
        print(f"Company website: {company_website or '(none)'}")
        print(f"Job title      : {job_title}")

        confirm = input("Is this correct? (y/n): ").strip().lower()
        if confirm == "y":
            return True
        elif confirm == "n":
            return False
        else:
            print("Please enter y or n.")


def main():
    clear_console()

    print("=" * 80)
    print("JOB DESCRIPTION SAVER")
    print("=" * 80)
    print()

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    archive_dir = os.path.join(OUTPUT_DIR, "job_desc_txt_archive")
    os.makedirs(archive_dir, exist_ok=True)

    print("Commands available while answering metadata:")
    print("  :back    -> redo previous answer")
    print("  :restart -> restart all metadata fields")
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

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    safe_company = "".join(c for c in company_name if c.isalnum() or c in (' ', '-', '_')).strip()
    safe_company = safe_company.replace(' ', '_')

    safe_job_title = "".join(c for c in job_title if c.isalnum() or c in (' ', '-', '_')).strip()
    safe_job_title = safe_job_title.replace(' ', '_')

    base_filename = f"{safe_company}_{safe_job_title}_{timestamp}"

    txt_filename = os.path.join(archive_dir, f"{base_filename}.txt")
    try:
        save_raw_text(job_text, txt_filename)
    except Exception as e:
        print(f"✗ Error saving text archive: {e}")

    docx_filename = os.path.join(OUTPUT_DIR, f"temp_{base_filename}.docx")
    pdf_filename = os.path.join(OUTPUT_DIR, f"{base_filename}.pdf")

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

    print()
    print("=" * 80)
    print("DONE!")
    print(f"Main output file: {final_created_file}")
    print(f"Text archive file: {txt_filename}")
    print(f"Output directory: {OUTPUT_DIR}")
    print("=" * 80)


if __name__ == "__main__":
    main()