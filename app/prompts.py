import sys, os
from .config import CV_OPTIONS, EXCEL_FIELDS
from .file_service import clear_console


def collect_job_metadata():
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

        if cmd == ":cls":
            clear_console()
            print("Metadata entry restarted (console cleared).\n")
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
            print("Please enter 'y' or 'n'.")


def collect_excel_fields():
    answers = {key: "" for key, _, _ in EXCEL_FIELDS}
    index = 0

    print("Portal: 1-Linkedin, 2-Direct, 3-Website, 4-Email, 5-Linkedin Easy Apply")

    while index < len(EXCEL_FIELDS):
        field_key, prompt, required = EXCEL_FIELDS[index]
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
            print("Restarting Excel field entry...\n")
            answers = {key: "" for key, _, _ in EXCEL_FIELDS}
            index = 0
            print("Portal: 1-Linkedin, 2-Direct, 3-Website, 4-Email, 5-Linkedin Easy Apply")
            continue

        if cmd == ":cls":
            clear_console()
            print("Excel field entry restarted (console cleared).\n")
            print("Portal: 1-Linkedin, 2-Direct, 3-Website, 4-Email, 5-Linkedin Easy Apply")
            answers = {key: "" for key, _, _ in EXCEL_FIELDS}
            index = 0
            continue

        if cmd == ":back":
            if index == 0:
                print("Already at the first Excel field.")
            else:
                index -= 1
                print("Going back to previous question.")
            continue

        if required and not user_input:
            print("This field is required!")
            continue

        if field_key == "portal_choice" and user_input and user_input not in {"1", "2", "3", "4", "5"}:
            print("Please enter 1, 2, 3, 4, or 5.")
            continue

        answers[field_key] = user_input
        index += 1

    return answers


def confirm_excel_fields(data):
    while True:
        print("\nPlease confirm Excel tracker data:")
        for i, (key, _, _) in enumerate(EXCEL_FIELDS, start=1):
            print(f"{i}. {key}: {data.get(key, '') or '(empty)'}")

        print("\nEnter:")
        print("  y    -> confirm")
        print("  n    -> restart all Excel fields")
        print(f"  1-{len(EXCEL_FIELDS)} -> edit a specific field")
        print("  :cls -> clear console")
        print("  :quit -> exit")

        choice = input("Your choice: ").strip().lower()

        if choice == "y":
            return "confirm"

        if choice == "n":
            return "restart"

        if choice == ":quit":
            print("Exiting...")
            sys.exit(0)

        if choice == ":cls":
            clear_console()
            continue

        if choice.isdigit():
            field_num = int(choice)
            if 1 <= field_num <= len(EXCEL_FIELDS):
                return field_num - 1

        print("Invalid input. Enter y, n, :cls, :quit, or a field number.")


def edit_excel_field(data, field_index):
    field_key, prompt, required = EXCEL_FIELDS[field_index]

    while True:
        current_value = data.get(field_key, "")
        user_input = input(f"{prompt}[current: {current_value}] ").strip()

        cmd = user_input.lower()

        if cmd == ":quit":
            print("Exiting...")
            sys.exit(0)

        if cmd == ":cls":
            clear_console()
            print("Editing selected Excel field again.\n")
            continue

        if cmd == ":back":
            print("Already editing a single selected field. Re-enter the value or press Enter to keep current.")
            continue

        if cmd == ":restart":
            return "restart"

        if not user_input:
            if current_value:
                return data
            if required:
                print("This field is required!")
                continue

        if field_key == "portal_choice" and user_input and user_input not in {"1", "2", "3", "4", "5"}:
            print("Please enter 1, 2, 3, 4, or 5.")
            continue

        if user_input:
            data[field_key] = user_input

        return data

def select_cv_source():
    """
    Ask which CV to copy.
    Supports:
    :cls  -> clear console
    :quit -> exit program
    """
    while True:
        print("\nSelect CV to copy:")
        for key, item in CV_OPTIONS.items():
            print(f"  {key}. {item['label']}")
            print(f"     {item['path']}")

        choice = input("Choose CV (1-2): ").strip().lower()

        if choice == ":quit":
            print("Exiting...")
            sys.exit(0)

        if choice == ":cls":
            clear_console()
            continue

        if choice not in CV_OPTIONS:
            print("Please enter 1 or 2.")
            continue

        selected = CV_OPTIONS[choice]
        selected_path = selected["path"]

        if not os.path.isfile(selected_path):
            print(f"⚠ Selected CV file not found: {selected_path}")
            continue

        print(f"✓ Selected CV: {selected['label']}")
        return selected_path, selected["label"]