import os
import email
import re
import mimetypes
from email import policy
from email.parser import BytesParser
from pathlib import Path
from io import BytesIO # Needed for xhtml2pdf
from xhtml2pdf import pisa # Import pisa
import logging

# --- Configuration ---
EML_FOLDER = Path("./INPUT")  # <<<--- CHANGE THIS to your folder containing .eml files
OUTPUT_FOLDER = Path("./OUTPUT") # <<<--- CHANGE THIS to where you want the results
# --- End Configuration ---

# Configure logging for xhtml2pdf (optional, reduces console noise)
# logging.basicConfig(level=logging.WARNING) # Or INFO, DEBUG
# logging.getLogger("xhtml2pdf").setLevel(logging.WARNING)

# --- Helper Function to Sanitize Filenames ---
def sanitize_filename(filename):
    """Removes or replaces characters invalid in filenames."""
    filename = filename.strip()
    filename = re.sub(r'[\\/*?:"<>|]', "_", filename)
    filename = re.sub(r'_+', '_', filename)
    if not filename:
        filename = "unnamed_file"
    return filename

# --- Function to convert HTML to PDF using xhtml2pdf ---
def html_to_pdf(source_html, output_filename):
    """Converts HTML string to a PDF file using xhtml2pdf."""
    try:
        with open(output_filename, "w+b") as result_file:
            # Ensure HTML is UTF-8 encoded bytes for pisa
            html_bytes = source_html.encode('utf-8')
            pisa_status = pisa.CreatePDF(
                BytesIO(html_bytes), # Use BytesIO for input
                dest=result_file,
                encoding='utf-8'
            )
        return pisa_status.err
    except Exception as e:
        print(f"  [Error] xhtml2pdf conversion failed: {e}")
        # Create an empty file to indicate failure, prevent crashing later if file expected
        try:
            output_filename.touch()
        except OSError: pass
        return 1 # Indicate error

# --- Main Processing Logic ---
def process_eml_file(eml_path, output_base_dir):
    """Processes a single .eml file."""
    print(f"Processing: {eml_path.name}")
    try:
        # Create a dedicated output folder for this email
        email_output_folder_name = sanitize_filename(eml_path.stem)
        email_output_dir = output_base_dir / email_output_folder_name
        email_output_dir.mkdir(parents=True, exist_ok=True)

        # Read and parse the email file
        with open(eml_path, 'rb') as fp:
            msg = BytesParser(policy=policy.default).parse(fp)

        # --- Extract Email Body ---
        body_html = None
        body_text = None

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))

                if "attachment" in content_disposition or "inline" in content_disposition:
                    continue

                if content_type == "text/html" and not body_html:
                    try:
                        charset = part.get_charset() or 'utf-8'
                        body_html = part.get_payload(decode=True).decode(charset, errors='replace')
                    except (LookupError, ValueError, TypeError) as e:
                        print(f"  [Warning] Could not decode HTML part: {e}")
                        body_html = part.get_payload(decode=True).decode('utf-8', errors='replace')

                elif content_type == "text/plain" and not body_html and not body_text:
                     try:
                        charset = part.get_charset() or 'utf-8'
                        body_text = part.get_payload(decode=True).decode(charset, errors='replace')
                     except (LookupError, ValueError, TypeError) as e:
                        print(f"  [Warning] Could not decode text part: {e}")
                        body_text = part.get_payload(decode=True).decode('utf-8', errors='replace')
        else:
             content_type = msg.get_content_type()
             if content_type == "text/html":
                 try:
                    charset = msg.get_charset() or 'utf-8'
                    body_html = msg.get_payload(decode=True).decode(charset, errors='replace')
                 except (LookupError, ValueError, TypeError) as e:
                    print(f"  [Warning] Could not decode HTML part: {e}")
                    body_html = msg.get_payload(decode=True).decode('utf-8', errors='replace')
             elif content_type == "text/plain":
                 try:
                    charset = msg.get_charset() or 'utf-8'
                    body_text = msg.get_payload(decode=True).decode(charset, errors='replace')
                 except (LookupError, ValueError, TypeError) as e:
                    print(f"  [Warning] Could not decode text part: {e}")
                    body_text = msg.get_payload(decode=True).decode('utf-8', errors='replace')


        # --- Convert body to PDF using xhtml2pdf ---
        body_pdf_path = email_output_dir / "email_body.pdf"
        pdf_generated = False
        if body_html:
            # Add basic CSS within HTML as xhtml2pdf handles stylesheets differently
            # Basic wrapper to help with potential encoding issues and basic styling
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body {{ font-family: sans-serif; line-height: 1.4; }}
                    pre {{ white-space: pre-wrap; word-wrap: break-word; }}
                </style>
            </head>
            <body>
                {body_html}
            </body>
            </html>
            """
            print("  Attempting HTML body conversion...")
            error_code = html_to_pdf(html_content, body_pdf_path)
            if not error_code:
                print(f"  Saved body (HTML) to: {body_pdf_path.name}")
                pdf_generated = True
            else:
                print(f"  [Error] Failed to convert HTML body to PDF using xhtml2pdf.")
                # Fall through to try text version if HTML failed badly
        elif body_text:
             # Wrap plain text in <pre> tags for basic formatting preservation
            print("  Attempting Text body conversion...")
            html_from_text = f"<html><head><meta charset='utf-8'></head><body><pre>{body_text}</pre></body></html>"
            error_code = html_to_pdf(html_from_text, body_pdf_path)
            if not error_code:
                print(f"  Saved body (Text) to: {body_pdf_path.name}")
                pdf_generated = True
            else:
                 print(f"  [Error] Failed to convert Text body to PDF using xhtml2pdf.")

        if not pdf_generated and not body_html and not body_text:
            print("  [Warning] No suitable text/html or text/plain body found to convert.")


        # --- Extract PDF Attachments (Same as before) ---
        attachment_count = 0
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue

            if part.get('Content-Disposition') is None and part.get_content_type() != 'application/pdf':
                 continue

            if part.get_content_type() == 'application/pdf' or \
               (part.get_filename() and part.get_filename().lower().endswith('.pdf')):

                filename = part.get_filename()
                if filename:
                    filename = sanitize_filename(filename)
                    if part.get_content_type() == 'application/pdf' and not filename.lower().endswith('.pdf'):
                         filename += ".pdf"
                else:
                    ext = mimetypes.guess_extension(part.get_content_type()) or ".pdf"
                    attachment_count += 1
                    filename = f"attachment_{attachment_count}{ext}"
                    filename = sanitize_filename(filename)

                if filename.lower().endswith('.pdf'):
                    attachment_path = email_output_dir / filename
                    try:
                        payload = part.get_payload(decode=True)
                        if payload:
                            with open(attachment_path, 'wb') as f_attach:
                                f_attach.write(payload)
                            print(f"  Saved attachment: {attachment_path.name}")
                        else:
                            print(f"  [Warning] Attachment '{filename}' has no payload.")
                    except Exception as e:
                        print(f"  [Error] Could not save attachment '{filename}': {e}")
                # else: # Optional: Log skipped non-PDF attachments
                #     print(f"  Skipping non-PDF attachment: {part.get_filename() or 'unnamed'}")


    except FileNotFoundError:
        print(f"[Error] File not found: {eml_path}")
    except Exception as e:
        print(f"[Error] Failed to process {eml_path.name}: {e}")


# --- Script Execution ---
if __name__ == "__main__":
    if not EML_FOLDER.is_dir():
        print(f"Error: Input folder '{EML_FOLDER}' not found or is not a directory.")
        exit(1)

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    print(f"Input folder: {EML_FOLDER.resolve()}")
    print(f"Output folder: {OUTPUT_FOLDER.resolve()}")
    print("-" * 30)

    processed_files = 0
    for item in EML_FOLDER.iterdir():
        if item.is_file() and item.suffix.lower() == ".eml":
            process_eml_file(item, OUTPUT_FOLDER)
            processed_files += 1
            print("-" * 10)

    print(f"\nFinished processing. Processed {processed_files} .eml file(s).")