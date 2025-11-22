import os
import win32com.client

# -----------------------------
# CONFIGURATION
# -----------------------------
ROOT = r"\\campwisefs\users"

# Shortcuts are no longer used for export folder replacement,
# but we keep this list in case needed later.
SHORTCUTS = ["GDrive.lnk", "OneDrive.lnk", "OneDriveBusiness.lnk", "Dropbox.lnk"]

# The INI filename to search for
INI_FILENAME = "DataLink_Viewer.ini"

# Safe testing mode: no files are modified when True
DRY_RUN = False
# -----------------------------


def resolve_shortcut(path):
    """Return the target path of a .lnk Windows shortcut."""
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(path)
    return shortcut.TargetPath


def update_ini(ini_path):
    """Update SMTP values + Last_Saved_Export_Folder prefix replacement."""
    with open(ini_path, "r") as f:
        lines = f.readlines()

    in_file_locations = False
    updated_export_folder = False
    new_lines = []

    for line in lines:
        stripped = line.strip()

        # Track when inside [File_Locations] section
        if stripped.lower() == "[file_locations]":
            in_file_locations = True
            new_lines.append(line)
            continue

        elif stripped.startswith("[") and stripped.endswith("]"):
            in_file_locations = False
            new_lines.append(line)
            continue

        # ---------------------------
        # Update SMTP server
        # ---------------------------
        if stripped.startswith("Email_Default_SMTP_Server="):
            new_lines.append("Email_Default_SMTP_Server=webmail.campwise.com\n")
            continue

        # Update SMTP port
        if stripped.startswith("Email_SMTP_Port="):
            new_lines.append("Email_SMTP_Port=465\n")
            continue

        # ---------------------------
        # Update export folder path prefix
        # ---------------------------
        if in_file_locations and stripped.lower().startswith("last_saved_export_folder="):
            original_value = stripped.split("=", 1)[1]  # everything after =

            # Replace only the prefix
            new_value = original_value.replace(
                r"\\rackstation\remoteSync",
                r"\\storage-gateway\campwise-remotesync"
            )

            new_lines.append(f"Last_Saved_Export_Folder={new_value}\n")
            updated_export_folder = True
            continue

        # Default: keep line unchanged
        new_lines.append(line)

    # ---------------------------
    # Write or dry-run output
    # ---------------------------
    if updated_export_folder:
        if DRY_RUN:
            print(f"\n>>> DRY RUN: Would update: {ini_path}")

            print("----- ORIGINAL -----")
            print("".join(lines))
            print("----- NEW -----")
            print("".join(new_lines))
            print("--------------------\n")

        else:
            with open(ini_path, "w") as f:
                f.writelines(new_lines)
            print(f"Updated INI: {ini_path}")
    else:
        print(f"No Last_Saved_Export_Folder found in: {ini_path}")


def process_folder(top_folder):
    """Walk through all subfolders and update INI files."""
    print(f"Scanning folder: {top_folder}")

    # Walk recursively for INI files
    for root, dirs, files in os.walk(top_folder):
        for file in files:
            if file.lower() == INI_FILENAME.lower():
                ini_path = os.path.join(root, file)
                update_ini(ini_path)


def main():
    # Process each user folder inside ROOT
    for folder in os.listdir(ROOT):
        folder_path = os.path.join(ROOT, folder)
        if os.path.isdir(folder_path):
            print(f"\n------------------------------")
            print(f"Processing user folder: {folder_path}")
            print(f"------------------------------")
            process_folder(folder_path)


if __name__ == "__main__":
    main()
