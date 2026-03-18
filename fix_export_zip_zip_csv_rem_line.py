import zipfile
import io
import os

MAIN_ZIP = "C:\\Users\\alex_\\Downloads\\out_HZ_745187_SuperPharm_18_03_2026-Copy.zip"        # orig file
OUTPUT_ZIP = "C:\\Users\\alex_\\Downloads\\out_HZ_745187_SuperPharm_18_03_2026-fixed.zip"

REMOVE_PREFIX = '"2026/03/16","12:20"' # problematic lines


def clean_csv_content(content: str) -> str:
    """Remove lines starting with the target prefix."""
    cleaned_lines = [
        line for line in content.splitlines()
        if not line.startswith(REMOVE_PREFIX)
    ]
    return "\n".join(cleaned_lines)


with zipfile.ZipFile(MAIN_ZIP, 'r') as main_zip:
    with zipfile.ZipFile(OUTPUT_ZIP, 'w', zipfile.ZIP_DEFLATED) as out_zip:

        for inner_name in main_zip.namelist():

            # If it's NOT a zip, just copy it through
            if not inner_name.lower().endswith(".zip"):
                out_zip.writestr(inner_name, main_zip.read(inner_name))
                continue

            # Process nested zip
            inner_bytes = main_zip.read(inner_name)
            inner_zip_file = io.BytesIO(inner_bytes)

            with zipfile.ZipFile(inner_zip_file, 'r') as inner_zip:
                # Create a new zip in memory
                new_inner_zip_bytes = io.BytesIO()
                with zipfile.ZipFile(new_inner_zip_bytes, 'w', zipfile.ZIP_DEFLATED) as new_inner_zip:

                    for csv_name in inner_zip.namelist():

                        if csv_name.lower().endswith(".csv"):
                            # Read CSV, clean it, write back
                            csv_data = inner_zip.read(csv_name).decode("utf-8", errors="ignore")
                            cleaned = clean_csv_content(csv_data)
                            new_inner_zip.writestr(csv_name, cleaned)
                        else:
                            # Copy non-CSV files unchanged
                            new_inner_zip.writestr(csv_name, inner_zip.read(csv_name))

                # Write the cleaned nested zip back into the main output zip
                out_zip.writestr(inner_name, new_inner_zip_bytes.getvalue())

print("Done! Cleaned ZIP saved as:", OUTPUT_ZIP)
