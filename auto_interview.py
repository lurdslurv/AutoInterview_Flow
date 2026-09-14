"""
AutoInterview - Igbinedion University, Okada (IUO)
Automated Interview Invitation Generator & WhatsApp Dispatcher

Supports:
1. Master Word Document (All letters in one .docx for internal tracking & mass printing)
2. Individual Word Documents
3. Individual PDFs named "[Name] - [Phone].pdf" for sending on WhatsApp
4. Root WhatsApp Dispatcher interactive dashboard (HTML) with 1-click WhatsApp links
"""

import os
import sys
import re
import csv
import copy
import shutil
import tempfile
import json
import urllib.parse
from datetime import datetime
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

try:
    import win32com.client
    WORD_COM_AVAILABLE = True
except ImportError:
    WORD_COM_AVAILABLE = False


def safe_save_doc(doc, target_path):
    """Saves document safely, falling back to an alternate name if open in Word."""
    try:
        doc.save(target_path)
        return target_path
    except PermissionError:
        base, ext = os.path.splitext(target_path)
        alt_path = f"{base}_updated{ext}"
        print(f"  [Note] '{os.path.basename(target_path)}' is open in another app. Saved as '{os.path.basename(alt_path)}' instead.")
        doc.save(alt_path)
        return alt_path


def safe_copy_template(src_path):
    """Safely make a copy of the template even if open in Word."""
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp_path = tmp.name
    tmp.close()
    
    try:
        shutil.copyfile(src_path, tmp_path)
        return tmp_path
    except Exception:
        pass

    import subprocess
    ps_cmd = f"""
    $src = [System.IO.File]::Open('{os.path.abspath(src_path)}', [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $dst = [System.IO.File]::Create('{os.path.abspath(tmp_path)}')
    $src.CopyTo($dst)
    $src.Close()
    $dst.Close()
    """
    res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], capture_output=True, text=True)
    if os.path.exists(tmp_path) and os.path.getsize(tmp_path) > 0:
        return tmp_path
    raise IOError(f"Could not read template file {src_path}: {res.stderr}")


def clean_phone_number(raw_phone):
    """Clean phone number and format for international WhatsApp (Nigeria 234 default)."""
    if not raw_phone:
        return "", ""
    digits = re.sub(r"[^\d]", "", str(raw_phone).strip())
    if not digits:
        return "", ""
    
    display_phone = str(raw_phone).strip()
    
    if digits.startswith("234"):
        wa_phone = digits
    elif digits.startswith("0") and len(digits) == 11:
        wa_phone = "234" + digits[1:]
    elif len(digits) == 10:
        wa_phone = "234" + digits
    else:
        wa_phone = digits
        
    return display_phone, wa_phone


def format_address_lines(address_str):
    """
    Intelligently splits and formats an address string into clean, formal postal lines.
    Handles single-line strings, multi-line strings, and varies line count intelligently.
    """
    if not address_str or not str(address_str).strip():
        return ["Warri, Nigeria."]
    
    raw = str(address_str).strip()
    
    # If already multi-line, respect existing line breaks
    if "\n" in raw:
        lines = [line.strip().rstrip(".,") for line in raw.split("\n") if line.strip()]
        if lines:
            lines[-1] = lines[-1] + "."
        return lines

    # Clean redundant spaces, trailing periods/commas, and split by commas
    raw = re.sub(r"\s+", " ", raw)
    parts = [p.strip().rstrip(".,") for p in raw.split(",") if p.strip()]
    
    if len(parts) <= 1:
        # Single line address (e.g. "Okada Town, Edo State.")
        return [raw.rstrip(".,") + "."]
    elif len(parts) == 2:
        # Two parts (e.g. "Warri", "Nigeria" -> "Warri, Nigeria.")
        return [f"{parts[0]}, {parts[1]}."]
    elif len(parts) == 3:
        # Three parts (e.g. "No. 15 Sapele Road", "Benin City", "Edo State")
        # Line 1: Street, Line 2: City, State
        return [parts[0], f"{parts[1]}, {parts[2]}."]
    elif len(parts) == 4:
        # Four parts (e.g. "12 Boundary Road", "GRA", "Benin City", "Edo State")
        # Line 1: Street & Area, Line 2: City & State
        return [f"{parts[0]}, {parts[1]}", f"{parts[2]}, {parts[3]}."]
    else:
        # Five or more parts (e.g. "Flat 4", "Palm View Estate", "Airport Road", "Effurun", "Delta State")
        # Line 1: Flat & Estate, Line 2: Street & Town, Line 3: State
        line1 = f"{parts[0]}, {parts[1]}"
        line2 = ", ".join(parts[2:-1])
        line3 = f"{parts[-1]}."
        return [line1, line2, line3]


def load_applicants(data_path):
    """Reads CSV file and returns normalized list of applicants."""
    applicants = []
    ext = os.path.splitext(data_path)[1].lower()
    
    if ext == ".csv":
        with open(data_path, mode="r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Please provide a CSV file.")
    
    for row in rows:
        clean_row = {}
        for k, v in row.items():
            if k:
                clean_row[k.strip().lower()] = str(v).strip() if v is not None else ""
        
        name = ""
        for key in ["name", "full name", "applicant name", "candidate", "candidate name"]:
            if key in clean_row and clean_row[key]:
                name = clean_row[key]
                break
        if not name:
            continue
            
        phone = ""
        for key in ["phone", "phone number", "mobile", "whatsapp", "tel", "contact"]:
            if key in clean_row and clean_row[key]:
                phone = clean_row[key]
                break
        display_phone, wa_phone = clean_phone_number(phone)
        
        address = ""
        for key in ["address", "full address", "residential address", "home address", "location"]:
            if key in clean_row and clean_row[key]:
                address = clean_row[key]
                break
                
        if not address:
            street = clean_row.get("street", "")
            city = clean_row.get("city", "")
            state = clean_row.get("state", "")
            addr_parts = [p for p in [street, city, state] if p]
            if addr_parts:
                address = ", ".join(addr_parts)
                
        address_lines = format_address_lines(address)
        
        interview_date = ""
        for key in ["interview_date", "interview date", "date"]:
            if key in clean_row and clean_row[key]:
                interview_date = clean_row[key]
                break
                
        interview_time = ""
        for key in ["interview_time", "interview time", "time"]:
            if key in clean_row and clean_row[key]:
                interview_time = clean_row[key]
                break
                
        applicants.append({
            "name": name,
            "phone": display_phone,
            "wa_phone": wa_phone,
            "address_lines": address_lines,
            "interview_date": interview_date,
            "interview_time": interview_time,
            "raw": row
        })
        
    return applicants


def format_ref_date_paragraph(p, doc, ref_no, letter_date):
    """
    Properly formats the Reference and Date header line using a single Right-Aligned Tab Stop.
    Eliminates clumsy multi-tabs and prevents longer dates (e.g. '14th September, 2026') from wrapping.
    """
    p.text = ""
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Dynamically calculate exact right margin width
    sec = doc.sections[0]
    content_width_inches = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches
    if content_width_inches <= 0:
        content_width_inches = 6.5
        
    p.paragraph_format.tab_stops.clear_all()
    p.paragraph_format.tab_stops.add_tab_stop(Inches(content_width_inches), WD_TAB_ALIGNMENT.RIGHT)
    
    r_ref = p.add_run(ref_no)
    r_ref.font.name = "Comic Sans MS"
    r_ref.font.bold = True
    
    # A single tab jumps flush to the right margin
    p.add_run("\t")
    
    r_date = p.add_run(letter_date)
    r_date.font.name = "Comic Sans MS"
    r_date.font.bold = True


def apply_candidate_to_document(doc, applicant, batch_settings):
    """
    Populates candidate information and batch details into a Word document.
    Prioritizes explicit placeholders {{NAME}}, {{ADDRESS}}, {{DATE}}, etc.
    """
    int_date = applicant["interview_date"] or batch_settings.get("interview_date", "Tuesday, 4th August, 2026")
    int_time = applicant["interview_time"] or batch_settings.get("interview_time", "10:00 AM")
    venue = batch_settings.get("venue", "Senate Chamber, Paul Dike Central Administration Building, Igbinedion University, Main Campus, Okada")
    ref_no = batch_settings.get("ref_no", "IUO/REG/PERS/IM/26")
    letter_date = batch_settings.get("letter_date", "3rd August, 2026")
    
    paragraphs = doc.paragraphs
    has_tags = any("{{" in p.text for p in paragraphs)
    
    if has_tags:
        # Template has explicit placeholders!
        for p in paragraphs:
            txt = p.text
            if "{{REF_NO}}" in txt or "{{LETTER_DATE}}" in txt:
                format_ref_date_paragraph(p, doc, ref_no, letter_date)
            elif "{{NAME}}" in txt:
                p.text = ""
                r = p.add_run(applicant["name"])
                r.font.name = "Comic Sans MS"
                r.font.bold = True
                r.font.underline = False # Explicitly NOT underlined
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
            elif "{{ADDRESS}}" in txt:
                p.text = ""
                for idx, line in enumerate(applicant["address_lines"]):
                    if idx > 0:
                        p.add_run().add_break() # Soft break
                    r = p.add_run(line)
                    r.font.name = "Comic Sans MS"
                    r.font.bold = False
                    r.font.underline = False
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.0
            elif "{{DATE}}" in txt:
                p.text = ""
                r_lbl = p.add_run("Date:\t \t")
                r_lbl.font.name = "Comic Sans MS"
                r_lbl.font.bold = True
                r_val = p.add_run(int_date)
                r_val.font.name = "Comic Sans MS"
                r_val.font.bold = False
            elif "{{TIME}}" in txt:
                p.text = ""
                r_lbl = p.add_run("Time: \t")
                r_lbl.font.name = "Comic Sans MS"
                r_lbl.font.bold = True
                r_val = p.add_run(int_time)
                r_val.font.name = "Comic Sans MS"
                r_val.font.bold = False
            elif "{{VENUE}}" in txt:
                p.text = ""
                r_lbl = p.add_run("Venue: \t")
                r_lbl.font.name = "Comic Sans MS"
                r_lbl.font.bold = True
                r_val = p.add_run(venue)
                r_val.font.name = "Comic Sans MS"
                r_val.font.bold = False
    else:
        # Fallback anchor-based replacement for untagged templates
        for p in paragraphs:
            if "IUO/REG" in p.text or ("/" in p.text and "August" in p.text):
                format_ref_date_paragraph(p, doc, ref_no, letter_date)
                break
                
        invitation_idx = -1
        for i, p in enumerate(paragraphs):
            if "INVITATION FOR INTERVIEW" in p.text.upper():
                invitation_idx = i
                break
                
        if invitation_idx > 0:
            # Clear all text between paragraph 0 and invitation_idx
            for i in range(1, invitation_idx):
                paragraphs[i].text = ""
            
            # Place name in paragraph 2 (bold, NO underline)
            name_p = paragraphs[min(2, invitation_idx - 1)]
            r_name = name_p.add_run(applicant["name"])
            r_name.font.name = "Comic Sans MS"
            r_name.font.bold = True
            r_name.font.underline = False
            name_p.paragraph_format.space_before = Pt(0)
            name_p.paragraph_format.space_after = Pt(2)
            
            # Place address in paragraph 3
            addr_p = paragraphs[min(3, invitation_idx - 1)]
            for idx, line in enumerate(applicant["address_lines"]):
                if idx > 0:
                    addr_p.add_run().add_break()
                r = addr_p.add_run(line)
                r.font.name = "Comic Sans MS"
                r.font.bold = False
                r.font.underline = False
            addr_p.paragraph_format.space_before = Pt(0)
            addr_p.paragraph_format.space_after = Pt(6)
            addr_p.paragraph_format.line_spacing = 1.0

        for i, p in enumerate(paragraphs):
            if p.text.strip().startswith("Date:"):
                p.text = ""
                r_label = p.add_run("Date:\t \t")
                r_label.font.name = "Comic Sans MS"
                r_label.font.bold = True
                r_val = p.add_run(int_date)
                r_val.font.name = "Comic Sans MS"
                r_val.font.bold = False
            elif p.text.strip().startswith("Time:"):
                p.text = ""
                r_label = p.add_run("Time: \t")
                r_label.font.name = "Comic Sans MS"
                r_label.font.bold = True
                r_val = p.add_run(int_time)
                r_val.font.name = "Comic Sans MS"
                r_val.font.bold = False
            elif p.text.strip().startswith("Venue:"):
                p.text = ""
                r_label = p.add_run("Venue: \t")
                r_label.font.name = "Comic Sans MS"
                r_label.font.bold = True
                r_val = p.add_run(venue)
                r_val.font.name = "Comic Sans MS"
                r_val.font.bold = False
                # If subsequent paragraph is an orphaned duplicate venue fragment, clear it
                if i + 1 < len(paragraphs) and "University, Main Campus, Okada" in paragraphs[i+1].text:
                    paragraphs[i+1].text = ""


def clone_letter_into_doc(src_doc, target_doc):
    """Clones all paragraphs from src_doc into target_doc preserving formatting."""
    for p in src_doc.paragraphs:
        new_p = target_doc.add_paragraph()
        if p._p.pPr is not None:
            new_p._p.get_or_add_pPr()
            new_p._p.replace(new_p._p.pPr, copy.deepcopy(p._p.pPr))
        for child in p._p:
            if child.tag.endswith('r'):
                new_p._p.append(copy.deepcopy(child))


def sanitize_filename(name):
    """Sanitize string for Windows filename."""
    clean = re.sub(r'[\\/*?:"<>|]', "", name).strip()
    return clean or "Applicant"


def convert_docx_folder_to_pdfs(docx_files_map):
    """
    Converts individual docx files to pdf using native Microsoft Word COM.
    docx_files_map: dict of {docx_abs_path: pdf_abs_path}
    """
    if not WORD_COM_AVAILABLE:
        print("[WARNING] pywin32 not available. Skipping automatic PDF export.")
        return False
        
    print(f"\n[Word Native Engine] Exporting {len(docx_files_map)} individual PDF letters...")
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        count = 0
        for docx_path, pdf_path in docx_files_map.items():
            if os.path.exists(docx_path):
                doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
                try:
                    doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=17) # 17 = wdFormatPDF
                except Exception:
                    base, ext = os.path.splitext(pdf_path)
                    alt_pdf = f"{base}_updated{ext}"
                    doc.SaveAs2(os.path.abspath(alt_pdf), FileFormat=17)
                doc.Close(False)
                count += 1
                print(f"  -> Generated PDF ({count}/{len(docx_files_map)}): {os.path.basename(pdf_path)}")
        print("[Word Native Engine] All PDFs exported successfully!")
        return True
    except Exception as e:
        print(f"[Word Native Engine] Error during PDF conversion: {e}")
        return False
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass


def generate_whatsapp_dispatcher_data(applicants_data, out_dir, base_dir):
    """Saves candidate data to Output/candidates.js for the dynamic WhatsApp Dispatcher."""
    candidates_list = []
    for app in applicants_data:
        name = app["name"]
        display_phone = app["phone"]
        wa_phone = app["wa_phone"]
        int_date = app["interview_date"] or "Tuesday, 4th August, 2026"
        int_time = app["interview_time"] or "10:00 AM"
        pdf_name = app["pdf_filename"]
        pdf_rel_path = f"Output/PDFs/{urllib.parse.quote(pdf_name)}"
        
        msg_text = (
            f"Dear {name},\n\n"
            f"Following your application for employment at Igbinedion University, Okada (IUO), "
            f"I am pleased to inform you that you have been shortlisted for an interview.\n\n"
            f"Date: {int_date}\n"
            f"Time: {int_time}\n"
            f"Venue: Senate Chamber, Paul Dike Central Administration Building, IUO Main Campus, Okada.\n\n"
            f"Please find attached your official invitation letter (PDF).\n\n"
            f"Best regards,\nCouncil, Personnel & General Administration\nIgbinedion University, Okada"
        )
        encoded_msg = urllib.parse.quote(msg_text)
        addr_preview = app['address_lines'][0] if app.get('address_lines') else ""
        
        candidates_list.append({
            "name": name,
            "phone": display_phone,
            "wa_phone": wa_phone,
            "address": addr_preview,
            "interview_date": int_date,
            "interview_time": int_time,
            "pdf_path": pdf_rel_path,
            "pdf_name": pdf_name,
            "encoded_msg": encoded_msg
        })
        
    js_data = {
        "title": "Igbinedion University, Okada - Interview Dispatcher",
        "subtitle": "Direct WhatsApp dispatching & PDF access for shortlisted applicants",
        "generated_at": datetime.now().strftime("%A, %d %B, %Y %I:%M %p"),
        "total_candidates": len(candidates_list),
        "candidates": candidates_list
    }
    
    candidates_js_path = os.path.join(out_dir, "candidates.js")
    with open(candidates_js_path, "w", encoding="utf-8") as f:
        f.write(f"window.BATCH_DATA = {json.dumps(js_data, indent=2)};\n")
    print(f"[Dispatcher] Active batch candidate data saved to:\n  -> {candidates_js_path}")
    
    root_dispatcher = os.path.join(base_dir, "WhatsApp_Dispatcher.html")
    return root_dispatcher


def select_applicants_file(base_dir):
    """Interactive file selector supporting CLI arg, drag-and-drop, Windows dialog, or numbered list."""
    # 1. Check if argument was passed (e.g. dragged onto .bat or CLI arg)
    if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        arg_path = sys.argv[1].strip('"')
        if os.path.exists(arg_path):
            return os.path.abspath(arg_path)
            
    csv_files = [f for f in os.listdir(base_dir) if f.lower().endswith(".csv")]
    
    if "--defaults" in sys.argv or "-y" in sys.argv:
        if csv_files:
            return os.path.join(base_dir, csv_files[0])
        return os.path.join(base_dir, "applicants_sample.csv")
        
    print("\n--- STEP 1: Select Applicants File ---")
    if csv_files:
        print("CSV files found in this folder:")
        for idx, f in enumerate(csv_files, 1):
            print(f"  [{idx}] {f}")
    print("  [B] Browse for another CSV / Excel file (opens Windows File Chooser)")
    print("  [T] Type custom file path")
    
    default_choice = "1" if csv_files else "B"
    user_choice = input(f"\nSelect an option [Default: {default_choice}]: ").strip()
    if not user_choice:
        user_choice = default_choice
        
    choice_upper = user_choice.upper()
    if choice_upper == "B":
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            selected = filedialog.askopenfilename(
                title="Select Applicants CSV File",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
            )
            root.destroy()
            if selected and os.path.exists(selected):
                return os.path.abspath(selected)
            print("No file selected in dialog. Falling back to default.")
        except Exception as e:
            print(f"Dialog unavailable ({e}). Falling back to default.")
    elif choice_upper == "T":
        custom_path = input("Enter file path: ").strip().strip('"')
        if os.path.exists(custom_path):
            return os.path.abspath(custom_path)
        print("File not found. Falling back to default.")
    elif user_choice.isdigit():
        idx = int(user_choice) - 1
        if 0 <= idx < len(csv_files):
            return os.path.join(base_dir, csv_files[idx])
            
    # Default fallback
    if csv_files:
        return os.path.join(base_dir, csv_files[0])
    return os.path.join(base_dir, "applicants_sample.csv")


def prompt_batch_settings():
    """Interactive prompt for batch interview settings, with optional ref number."""
    default_ref = "IUO/REG/PERS/IM/26"
    default_letter_date = "14th September, 2026"
    default_int_date = "Tuesday, 4th August, 2026"
    default_time = "10:00 AM"
    default_venue = "Senate Chamber, Paul Dike Central Administration Building, Igbinedion University, Main Campus, Okada"
    
    # Check if non-interactive mode requested
    if "--defaults" in sys.argv or "-y" in sys.argv:
        return {
            "ref_no": default_ref,
            "letter_date": default_letter_date,
            "interview_date": default_int_date,
            "interview_time": default_time,
            "venue": default_venue
        }
        
    print("\n--- STEP 2: Batch Settings (Press Enter to keep current default) ---")
    
    # Ref number is optional (defaults to current)
    ref_in = input(f"Reference Number [{default_ref}] (Press Enter to keep): ").strip()
    ref_no = ref_in if ref_in else default_ref
    
    ld_in = input(f"Letter Date      [{default_letter_date}] (Press Enter to keep): ").strip()
    letter_date = ld_in if ld_in else default_letter_date
    
    id_in = input(f"Interview Date   [{default_int_date}] (Press Enter to keep): ").strip()
    interview_date = id_in if id_in else default_int_date
    
    it_in = input(f"Default Time     [{default_time}] (Press Enter to keep): ").strip()
    interview_time = it_in if it_in else default_time
    
    venue_in = input(f"Venue (Press Enter to keep standard campus venue): ").strip()
    venue = venue_in if venue_in else default_venue
    
    return {
        "ref_no": ref_no,
        "letter_date": letter_date,
        "interview_date": interview_date,
        "interview_time": interview_time,
        "venue": venue
    }


def prompt_post_run_actions(master_docx_path, pdf_dir, dispatcher_html_path):
    """Presents a clean post-run menu so no unwanted tabs or folders pop up."""
    if "--defaults" in sys.argv or "-y" in sys.argv:
        return
        
    while True:
        print("\n" + "=" * 70)
        print("  WHAT WOULD YOU LIKE TO OPEN? (Select an option or press Enter to exit)")
        print("=" * 70)
        print("  [1] Open WhatsApp Dispatcher Dashboard in browser")
        print("  [2] Open Master Word Document (All Letters)")
        print("  [3] Open Output Folder in File Explorer")
        print("  [Enter] Finished / Exit")
        print("=" * 70)
        choice = input("Enter choice [1-3, or Enter]: ").strip()
        
        if choice == "1":
            os.system(f'start "" "{os.path.abspath(dispatcher_html_path)}"')
            print("-> Opened WhatsApp Dispatcher in your browser.")
        elif choice == "2":
            os.system(f'start "" "{os.path.abspath(master_docx_path)}"')
            print("-> Opened Master Word Document.")
        elif choice == "3":
            os.system(f'start "" "{os.path.abspath(os.path.dirname(master_docx_path))}"')
            print("-> Opened Output Folder in File Explorer.")
        else:
            print("\nDone! All files are generated and ready.")
            break


def main():
    print("=" * 70)
    print("  IUO INTERVIEW INVITATION AUTOMATION & WHATSAPP DISPATCHER")
    print("=" * 70)
    
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Priority: template_with_placeholders.docx > invite_template.docx > sample_template.docx
    candidate_template_with_tags = os.path.join(base_dir, "template_with_placeholders.docx")
    default_template = os.path.join(base_dir, "invite_template.docx")
    sample_template = os.path.join(base_dir, "sample_template.docx")
    
    if os.path.exists(candidate_template_with_tags):
        template_path = candidate_template_with_tags
    elif os.path.exists(default_template):
        template_path = default_template
    elif os.path.exists(sample_template):
        template_path = sample_template
    else:
        template_path = candidate_template_with_tags
        
    print(f"Template File: {template_path}")
    if not os.path.exists(template_path):
        print(f"[ERROR] Template file not found: {template_path}")
        return

    # 1. Interactive file selection
    csv_path = select_applicants_file(base_dir)
    print(f"Selected Applicants CSV: {csv_path}")
    if not os.path.exists(csv_path):
        print(f"[ERROR] Applicants file not found: {csv_path}")
        return
        
    # 2. Interactive batch settings
    batch_settings = prompt_batch_settings()
    
    # 3. Setup Output folders
    out_dir = os.path.join(base_dir, "Output")
    docx_dir = os.path.join(out_dir, "DOCX")
    pdf_dir = os.path.join(out_dir, "PDFs")
    os.makedirs(docx_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)
    
    # 4. Load applicants
    applicants = load_applicants(csv_path)
    print(f"\nLoaded {len(applicants)} applicants successfully.")
    if not applicants:
        print("[WARNING] No applicants found in file.")
        return
        
    clean_template_copy = safe_copy_template(template_path)
    master_doc = docx.Document(clean_template_copy)
    docx_to_pdf_map = {}
    
    print("\nProcessing applicants and generating documents...")
    for idx, applicant in enumerate(applicants):
        name = applicant["name"]
        phone = applicant["phone"]
        
        if phone:
            base_filename = f"{sanitize_filename(name)} - {sanitize_filename(phone)}"
        else:
            base_filename = f"{sanitize_filename(name)}"
            
        individual_docx_path = os.path.join(docx_dir, f"{base_filename}.docx")
        individual_pdf_path = os.path.join(pdf_dir, f"{base_filename}.pdf")
        
        # Build individual document
        indiv_doc = docx.Document(clean_template_copy)
        apply_candidate_to_document(indiv_doc, applicant, batch_settings)
        actual_docx_path = safe_save_doc(indiv_doc, individual_docx_path)
        
        docx_to_pdf_map[actual_docx_path] = individual_pdf_path
        applicant["pdf_filename"] = f"{base_filename}.pdf"
        
        # Build Master Document
        if idx == 0:
            apply_candidate_to_document(master_doc, applicant, batch_settings)
        else:
            master_doc.add_page_break()
            clone_letter_into_doc(indiv_doc, master_doc)
            
        print(f"  [{idx+1}/{len(applicants)}] Formatted letter for: {name} (Phone: {phone or 'N/A'})")
        
    master_docx_path = os.path.join(out_dir, "All_Interview_Invitations.docx")
    master_docx_path = safe_save_doc(master_doc, master_docx_path)
    print(f"\n[Master Document] All {len(applicants)} letters combined into:")
    print(f"  -> {master_docx_path}")
    
    # 5. Convert individual docx to pdf
    convert_docx_folder_to_pdfs(docx_to_pdf_map)
    
    # 6. Generate WhatsApp Dispatcher Data
    root_dispatcher_path = generate_whatsapp_dispatcher_data(applicants, out_dir, base_dir)
    
    if os.path.exists(clean_template_copy):
        try:
            os.remove(clean_template_copy)
        except Exception:
            pass
            
    print("\n" + "=" * 70)
    print("  BATCH GENERATION COMPLETED SUCCESSFULLY!")
    print("=" * 70)
    print(f"  Master Word File     : {master_docx_path}")
    print(f"  Individual PDFs (x{len(applicants)}): {pdf_dir}")
    print(f"  WhatsApp Dispatcher  : {root_dispatcher_path}")
    print("=" * 70)
    
    # 7. Clean post-run menu (no auto-popups!)
    prompt_post_run_actions(master_docx_path, pdf_dir, root_dispatcher_path)


if __name__ == "__main__":
    main()
