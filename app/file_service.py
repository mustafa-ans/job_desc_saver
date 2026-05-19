import os
import shutil
import pyperclip
from .config import (
    COVER_LETTER_SOURCE_PATH,
    COVER_LETTER_SOURCE_PATH_PDF,
    COVER_LETTER_MASTER_PATH,
    COVER_LETTER_WORKING_NAME,
)


def clear_console():
    os.system('cls' if os.name == 'nt' else 'clear')


def copy_to_clipboard(text):
    try:
        pyperclip.copy(text)
        print(f"✓ Copied to clipboard: {text}")
    except Exception as e:
        print(f"⚠ Could not copy to clipboard: {e}")


def move_cover_letter_to_folder(source_path, destination_folder):
    try:
        if not source_path.strip():
            print("⚠ No cover letter path provided. Skipping move.")
            return ""

        source_path = source_path.strip().strip('"')

        if not os.path.isfile(source_path):
            print(f"⚠ Cover letter file not found: {source_path}")
            return ""

        destination_path = os.path.join(destination_folder, os.path.basename(source_path))
        moved_path = shutil.move(source_path, destination_path)
        print(f"✓ Cover letter moved to: {moved_path}")
        return moved_path
    except Exception as e:
        print(f"⚠ Error moving cover letter: {e}")
        return ""


def archive_and_reset_cover_letter(destination_folder):
    try:
        moved_cover_letter_path = ""
        recreated_cover_letter_path = ""
        moved_cover_letter_pdf_path = ""

        if not os.path.isfile(COVER_LETTER_SOURCE_PATH):
            print(f"⚠ Working cover letter not found: {COVER_LETTER_SOURCE_PATH}")
            return "", ""

        if not os.path.isfile(COVER_LETTER_MASTER_PATH):
            print(f"⚠ Master cover letter not found: {COVER_LETTER_MASTER_PATH}")
            return "", ""

        if os.path.isfile(COVER_LETTER_SOURCE_PATH_PDF):
            print(f"✓ Found existing cover letter PDF: {COVER_LETTER_SOURCE_PATH_PDF}")
            moved_cover_letter_pdf_path = move_cover_letter_to_folder(COVER_LETTER_SOURCE_PATH_PDF, destination_folder)
            if not moved_cover_letter_pdf_path:
                print("⚠ Failed to move existing cover letter PDF.")
        else:
            print(f"✓ No existing cover letter PDF found at: {COVER_LETTER_SOURCE_PATH_PDF}")

        destination_cover_letter_path = os.path.join(destination_folder, COVER_LETTER_WORKING_NAME)

        moved_cover_letter_path = shutil.move(COVER_LETTER_SOURCE_PATH, destination_cover_letter_path)
        print(f"✓ Cover letter moved to: {moved_cover_letter_path}")

        recreated_cover_letter_path = shutil.copy2(COVER_LETTER_MASTER_PATH, COVER_LETTER_SOURCE_PATH)
        print(f"✓ Fresh working cover letter recreated: {recreated_cover_letter_path}")

        return moved_cover_letter_path, recreated_cover_letter_path

    except Exception as e:
        print(f"⚠ Error archiving/resetting cover letter: {e}")
        return "", ""


def copy_cv_to_folder(source_path, destination_folder):
    try:
        if not os.path.isfile(source_path):
            print(f"⚠ CV file not found: {source_path}")
            return ""

        destination_path = os.path.join(destination_folder, os.path.basename(source_path))
        copied_path = shutil.copy2(source_path, destination_path)
        print(f"✓ CV copied to: {copied_path}")
        return copied_path
    except Exception as e:
        print(f"⚠ Error copying CV: {e}")
        return ""


def sanitize_name(text):
    text = "".join(c for c in text if c.isalnum() or c in (' ', '-', '_')).strip()
    text = "_".join(text.split())
    return text or "untitled"


def build_safe_base_name(output_dir, company_name, job_title, extra_suffix=""):
    MAX_FULL_PATH = 259

    safe_company = sanitize_name(company_name)
    safe_job_title = sanitize_name(job_title)

    separator = "_"
    candidate = f"{safe_company}{separator}{safe_job_title}"

    reserved_len = len(os.path.join(output_dir, "")) + len(extra_suffix)
    max_base_len = MAX_FULL_PATH - reserved_len

    if max_base_len < 20:
        raise ValueError("Output directory path is too long to safely create files.")

    if len(candidate) <= max_base_len:
        return candidate

    available = max_base_len - len(separator)
    min_each = 8
    company_share = max(min_each, available // 2)
    title_share = max(min_each, available - company_share)

    if company_share + title_share > available:
        title_share = available - company_share

    safe_company = safe_company[:company_share].rstrip("_- ")
    safe_job_title = safe_job_title[:title_share].rstrip("_- ")

    candidate = f"{safe_company}{separator}{safe_job_title}"

    if len(candidate) > max_base_len:
        candidate = candidate[:max_base_len].rstrip("_- ")

    return candidate
